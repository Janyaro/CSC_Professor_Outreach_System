"""
Word -> Gmail Drafts  (second app, FREE - Gemini free API)
----------------------------------------------------------
1. You give it the Word file that contains your professor emails.
2. It reads every email (To, Subject, body).
3. Emails already in YOUR format are kept; the others are converted to your
   format (or kept exactly as they are - your choice).
4. ALL emails are saved as Gmail DRAFTS in ONE click (nothing is ever sent
   automatically).

Run:  streamlit run docx_to_drafts.py
"""
import hashlib
import imaplib
import io
import json
import os
import re
import time
from pathlib import Path
from email.message import EmailMessage

import requests
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from pypdf import PdfReader

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ---------------------------------------------------------------------------
# Gemini (free) helper
# ---------------------------------------------------------------------------


class FatalAPIError(Exception):
    pass


_last_call = [0.0]
MIN_INTERVAL = [6.5]


def llm(api_key, model, system, user, max_tokens=1500):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    body = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "user", "parts": [{"text": user}]}],
        "generationConfig": {"maxOutputTokens": max_tokens + 2000, "temperature": 0.7,
                             "responseMimeType": "application/json"},
    }
    for attempt in range(6):
        wait = MIN_INTERVAL[0] - (time.time() - _last_call[0])
        if wait > 0:
            time.sleep(wait)
        _last_call[0] = time.time()
        r = requests.post(url, json=body, headers={"x-goog-api-key": api_key}, timeout=90)
        if r.status_code == 200:
            try:
                parts = r.json()["candidates"][0]["content"]["parts"]
                return "".join(p.get("text", "") for p in parts)
            except (KeyError, IndexError):
                raise ValueError("Empty response (blocked or no output)")
        if r.status_code == 429 and re.search(r"per\s*day|daily|PerDay", r.text, re.I):
            raise FatalAPIError("Gemini DAILY free quota is finished. Try again tomorrow, "
                                "or use the 'keep emails as in the Word file' option.")
        if r.status_code in (429, 500, 503):
            m = re.search(r'"retryDelay":\s*"(\d+(?:\.\d+)?)s"', r.text)
            time.sleep(float(m.group(1)) + 2 if m else 20 * (attempt + 1))
            continue
        raise FatalAPIError(f"Gemini error {r.status_code}: {r.text[:300]}")
    raise ValueError("Gemini kept refusing (rate limit). Try again later.")


def parse_json(text):
    m = re.search(r"\{.*\}", text.strip(), re.S)
    if not m:
        raise ValueError("No JSON found in model output")
    return json.loads(m.group(0))


# ---------------------------------------------------------------------------
# Reading the Word file
# ---------------------------------------------------------------------------
PARSE_LINES = []
SIG_RE = re.compile(r"^(best regards|kind regards|warm regards|regards|sincerely|"
                    r"yours sincerely|yours faithfully|best wishes|respectfully|"
                    r"with regards|yours truly)\s*[,.]?\s*$", re.I)


def _para_text(p_el, doc):
    """Full text of a paragraph, including hyperlinks and mailto targets."""
    out = []
    for el in p_el.iter():
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
        elif el.tag == qn("w:tab"):
            out.append(" ")
    extra = []
    for h in p_el.iter(qn("w:hyperlink")):
        rid = h.get(qn("r:id"))
        try:
            tgt = doc.part.rels[rid].target_ref if rid else ""
        except Exception:
            tgt = ""
        if tgt and tgt.lower().startswith("mailto:"):
            extra.append(tgt[7:].split("?")[0])
    return "".join(out) + ((" " + " ".join(extra)) if extra else "")


def parse_docx(data: bytes):
    """Return a list of emails found in the Word file:
    [{name, uni, email, subject, body, dear_line}].
    Reads paragraphs AND tables (hyperlinks included). An email starts at a line
    'Dear ...'; the lines just before it (To:, Subject:, a bare address,
    'Name - University' headings) are its header. An email body ends after the
    signature (Regards, + name), so the header of the NEXT email is not mixed in."""
    global PARSE_LINES
    doc = Document(io.BytesIO(data))
    lines = []                       # (text, heading_style)
    body_el = doc.element.body
    paras = []
    for child in body_el.iterchildren():
        if child.tag == qn("w:p"):
            paras.append(child)
        elif child.tag == qn("w:tbl"):
            paras.extend(child.iter(qn("w:p")))
    for p_el in paras:
        try:
            style = Paragraph(p_el, doc).style.name or ""
        except Exception:
            style = ""
        text = _para_text(p_el, doc).replace("\r", "")
        for ln in text.split("\n"):
            lines.append((ln.rstrip(), style.startswith("Heading")))
    PARSE_LINES = [t for t, _ in lines]

    recs, cur = [], None
    head = re.compile(r"^\s*[AB]?\s*\d+\s*[\.\)]\s*(.+)$")
    plain_head = re.compile(r"^\s*[AB]?\s*\d+\s*[\.\)]\s+\S.{0,100} - \S.{0,100}$")

    def new(name="", uni="", at=0):
        nonlocal cur
        cur = {"name": name, "uni": uni, "cands": [], "subject": "", "lines": [],
               "in_body": False, "sig": False, "done": False, "dear_line": at}
        recs.append(cur)

    for idx, (ln, is_head) in enumerate(lines):
        s = ln.strip()
        if not s:
            if cur and cur["in_body"] and not cur["done"]:
                cur["lines"].append("")
            continue
        free = cur is None or not cur["in_body"] or cur["done"]
        # numbered heading: "A1. Name - University"
        if is_head or (free and plain_head.match(s)):
            m = head.match(s)
            if m:
                name, _, uni = m.group(1).partition(" - ")
                new(name.strip(), uni.strip(), idx)
            continue
        mt = re.match(r"^(To|Email|E-mail|Mail)\s*:\s*(.*)$", s, re.I)
        ms = re.match(r"^Subject\s*:\s*(.*)$", s, re.I)
        if mt and (free or re.match(r"^To\b", s, re.I)):
            if cur is None or cur["in_body"]:
                new(at=idx)
            found = EMAIL_RE.findall(mt.group(2))
            cur["cands"] = found + cur["cands"]
            continue
        if ms and (free or cur["in_body"]):
            if cur is None or cur["in_body"]:
                new(at=idx)
            cur["subject"] = ms.group(1).strip()
            continue
        if re.match(r"^Dear\b", s):
            if cur is None or cur["in_body"]:
                new(at=idx)
            cur["in_body"] = True
            cur["dear_line"] = idx
            cur["lines"].append(ln.rstrip())
            continue
        if cur is not None and cur["in_body"] and not cur["done"]:
            cur["lines"].append(ln.rstrip())
            if SIG_RE.match(s):
                cur["sig"] = True
            elif cur["sig"]:
                cur["done"] = True          # signature name line = end of the email
            continue
        # header line of the next email (before 'Dear')
        if cur is None or cur["in_body"]:
            new(at=idx)
        if re.match(r"^(Profile|Match score|WeChat)\b", s, re.I):
            continue
        found = EMAIL_RE.findall(s)
        if found:
            cur["cands"].extend(found)
            continue
        mn = re.match(r"^(name|professor|prof\.?)\s*[:\-]\s*(.+)$", s, re.I)
        if mn and not cur["name"]:
            cur["name"] = mn.group(2).strip()

    # addresses that repeat in many emails are the APPLICANT's own (signature) - ignore
    n = max(1, len(recs))
    cnt = {}
    for r in recs:
        for a in set(x.lower() for x in r["cands"]):
            cnt[a] = cnt.get(a, 0) + 1
    body_cnt = {}
    for r in recs:
        for a in set(x.lower() for x in EMAIL_RE.findall("\n".join(r["lines"]))):
            body_cnt[a] = body_cnt.get(a, 0) + 1
    common = {a for a, c in cnt.items() if c > max(2, 0.3 * n)} | \
             {a for a, c in body_cnt.items() if c > max(2, 0.3 * n)}
    out = []
    for r in recs:
        body = re.sub(r"\n{3,}", "\n\n", "\n".join(r["lines"]).strip())
        if not body:
            continue
        cands = [a for a in r["cands"] if a.lower() not in common]
        if not cands:   # last resort: an address written inside the email text
            cands = [a for a in EMAIL_RE.findall(body) if a.lower() not in common]
        out.append({"name": r["name"], "uni": r["uni"],
                    "email": cands[0] if cands else "",
                    "subject": r["subject"], "body": body,
                    "dear_line": r["dear_line"]})
    return out


# ---------------------------------------------------------------------------
# Your format
# ---------------------------------------------------------------------------
TEMPLATE = """\
Dear Professor (Professor's Name),

I hope this email finds you well.

My name is (Your Name), and I have completed my Bachelor's degree in (Your Field) from (University Name, Country).

During my Bachelor's studies, I conducted my research thesis on the topic "(Your Thesis Title)." In this research, I used (Research Methodologies/Techniques) to investigate (Brief Description of Your Research).

After reviewing your academic profile and research publications, I found that your research interests are closely related to my academic background and research interests.

Therefore, I am highly interested in pursuing my Master's/PhD degree under your supervision at (University Name). I believe that your academic guidance and research expertise would greatly contribute to my academic and research development.

I would be grateful if you could consider my application and provide me with an acceptance letter (if you are willing to supervise me) for the upcoming admission/scholarship intake.

I have attached my CV for your kind consideration.

Thank you very much for your time and consideration. I look forward to hearing from you.

Regards,

(Your Full Name)"""

MARKERS = [
    "i hope this email finds you well", "my name is", "during my bachelor",
    "after reviewing your academic profile and research publications",
    "therefore, i am highly interested",
    "i would be grateful if you could consider my application",
    "i have attached my", "thank you very much for your time and consideration",
    "regards,",
]


def is_in_format(body: str) -> bool:
    t = body.lower().replace("’", "'").replace("“", '"').replace("”", '"')
    t = re.sub(r"\s+", " ", t)
    return t.lstrip().startswith("dear") and all(m in t for m in MARKERS)


def needs_transcript(body: str) -> bool:
    """True if the email says it attaches the academic transcript."""
    t = body.lower()
    return bool(re.search(r"attach\w*[^.]{0,80}transcript|transcript[^.]{0,60}attach|"
                          r"transcripts?\b[^.]{0,40}(enclosed|attached)", t))


def _assemble(salutation, paras, me):
    return (
        f"{salutation},\n\nI hope this email finds you well.\n\n"
        + "\n\n".join(p for p in paras if p)
        + f"\n\nI have attached my {me['attach']} for your kind consideration.\n\n"
        "Thank you very much for your time and consideration. I look forward to "
        f"hearing from you.\n\nRegards,\n\n{me['name']}"
    )


def _salutation_of(rec):
    first = rec["body"].strip().split("\n", 1)[0].strip()
    return first.rstrip(",") if first.lower().startswith("dear") else \
        f"Dear Professor {rec['name']}".strip()


def quick_convert(rec, me):
    """FREE conversion without AI. Works on emails written by this system's
    earlier format (they contain 'Your research on ... particularly interests me'
    and '... at <University> through the Chinese Government Scholarship').
    Returns None when the needed pieces cannot be found."""
    b = rec["body"]
    m = re.search(r"Your research on (.+?) particularly interests me", b, re.S | re.I)
    mu = re.search(r"\bat (.+?) through the Chinese Government Scholarship", b, re.S)
    uni = (mu.group(1).strip() if mu else rec["uni"]).strip()
    if not (m and uni and me["name"] and me["field"] and me["thesis_title"]):
        return None
    topics = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(",")
    if re.search(r"broaden|transferable|keen to (explore|learn|expand)|new area|eager to", b, re.I):
        p3 = (f"After reviewing your academic profile and research publications, I found your "
              f"research on {topics} interesting, and I am keen to broaden my research "
              "experience in this area.")
    else:
        p3 = (f"After reviewing your academic profile and research publications, I found that "
              f"your research on {topics} is closely related to my academic background and "
              "research interests.")
    p2 = (f'During my Bachelor\'s studies, I conducted my research thesis on the topic '
          f'"{me["thesis_title"].rstrip(".")}."')
    if me["thesis_methods"] and me["thesis_summary"]:
        p2 += (f" In this research, I used {me['thesis_methods']} to investigate "
               f"{me['thesis_summary']}.")
    elif me["thesis_methods"]:
        p2 += f" In this research, I used {me['thesis_methods']}."
    paras = [
        f"My name is {me['name']}, and I have completed my Bachelor's degree in "
        f"{me['field']} from {me['university']}.",
        p2, p3,
        f"Therefore, I am highly interested in pursuing my {me['degree']} degree under your "
        f"supervision at {uni}. I believe that your academic guidance and research expertise "
        "would greatly contribute to my academic and research development.",
        "I would be grateful if you could consider my application and provide me with an "
        "acceptance letter (if you are willing to supervise me) for the upcoming "
        f"{me['intake']}.",
    ]
    return _assemble(_salutation_of(rec), paras, me)


def convert_email(api_key, model, rec, me, extra=""):
    """Rewrite one existing email into YOUR format. Professor facts come ONLY
    from the old email text."""
    degree = me["degree"]
    salutation = _salutation_of(rec)
    system = (
        "You rewrite an existing email from a prospective graduate student into the "
        "student's own template. Output ONLY valid JSON "
        '{"p1": "...", "p2": "...", "p3": "...", "p4": "...", "p5": "..."}.'
    )
    user = f"""APPLICANT DETAILS (the only facts you may use about the applicant):
{json.dumps(me, indent=1, ensure_ascii=False)}

EXISTING EMAIL (the only source of facts about the professor - his/her research
topics and university):
\"\"\"
{rec['body'][:3500]}
\"\"\"
Professor name: {rec['name'] or '(see email)'}   Professor university: {rec['uni'] or '(see email)'}

APPLICANT'S TEMPLATE (keep its wording and order):
\"\"\"
{TEMPLATE}
\"\"\"

Write the paragraphs:
p1 = "My name is {me['name']}, and I have completed my Bachelor's degree in {me['field']} from {me['university']}."
p2 = "During my Bachelor's studies, I conducted my research thesis on the topic
     "{me['thesis_title']}." In this research, I used {me['thesis_methods']} to
     investigate {me['thesis_summary']}."  (if the thesis title is empty, use the
     applicant's main project from the details; never invent one)
p3 = "After reviewing your academic profile and research publications, I found that
     your research on <1-2 SPECIFIC topics taken from the EXISTING EMAIL> is closely
     related to my academic background and research interests." (If the existing
     email names no specific topic, use the general sentence from the template.)
p4 = "Therefore, I am highly interested in pursuing my {degree} degree under your
     supervision at <professor university>. I believe that your academic guidance and
     research expertise would greatly contribute to my academic and research development."
p5 = "I would be grateful if you could consider my application and provide me with an
     acceptance letter (if you are willing to supervise me) for the upcoming
     {me['intake']}."
Rules: no greeting, no "I hope...", no attachment sentence, no thanks, no signature
(added automatically). No invented facts. Keep template wording.
{extra}"""
    d = parse_json(llm(api_key, model, system, user, 1500))
    paras = [d.get(k, "").strip() for k in ("p1", "p2", "p3", "p4", "p5")]
    return _assemble(salutation, paras, me)


# ---------------------------------------------------------------------------
# Local storage: settings (keys) and progress (what is already done)
# ---------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
SETTINGS = BASE / "settings.json"     # your keys + details (stays on THIS computer)
PROGRESS = BASE / "progress.json"     # which emails are converted / saved as drafts
FILES = BASE / "my_files"             # your CV / transcript copies
SAVE_KEYS = ["api_key", "model", "gmail", "app_pw", "d_name", "d_field", "d_uni",
             "d_title", "d_methods", "d_summary", "d_degree", "d_intake"]


def load_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    try:
        os.chmod(path, 0o600)      # only you can read it (Linux/Mac)
    except Exception:
        pass


def rec_key(rec):
    """Stable id of an email = hash of its ORIGINAL text in the Word file."""
    return hashlib.sha1(re.sub(r"\s+", " ", rec["body"]).strip().encode("utf-8")).hexdigest()[:16]


def saved_file(kind):
    for p in sorted(FILES.glob(f"{kind}__*")):
        return p.name.split("__", 1)[1], p.read_bytes()
    return None


def get_file(kind):
    """The uploaded CV/transcript, or the copy saved earlier on this computer."""
    up = st.session_state.get(f"{kind}_upl")
    if up:
        FILES.mkdir(exist_ok=True)
        for old in FILES.glob(f"{kind}__*"):
            old.unlink()
        (FILES / f"{kind}__{up.name}").write_bytes(up.getvalue())
        return up.name, up.getvalue()
    return saved_file(kind)


def prepare_records(recs, keep_all, force_all, use_quick, me, ai_limit, api_key, model,
                    progress, persist, tick=None):
    """Decide what happens to every email. Returns a summary (counts).
    Statuses: drafted_before / kept / in_format / converted_before /
    quick / ai / not_converted."""
    ai_used, stop = 0, False
    summary = {"total": len(recs), "drafted_before": 0, "kept": 0, "in_format": 0,
               "converted_before": 0, "quick": 0, "ai": 0, "not_converted": 0}
    drafted_emails = {v.get("email", "").lower() for v in progress.values()
                      if v.get("status") == "drafted" and v.get("email")}
    for k, r in enumerate(recs):
        r["key"] = rec_key(r)
        r["include"], r["drafted"] = True, False
        p = progress.get(r["key"], {})
        if not p and r["email"] and r["email"].lower() in drafted_emails:
            p = {"status": "drafted", "email": r["email"]}
        if p.get("status") == "drafted":
            r["status"], r["include"], r["drafted"] = \
                "already saved to Gmail drafts earlier - skipped", False, True
            if p.get("email") and not r["email"]:
                r["email"] = p["email"]
            summary["drafted_before"] += 1
        elif p.get("status") == "converted" and not keep_all:
            r["body"], r["status"] = p["body"], "converted earlier (saved result)"
            summary["converted_before"] += 1
        elif keep_all:
            r["status"] = "kept as in Word file"
            summary["kept"] += 1
        elif not force_all and is_in_format(r["body"]):
            r["status"] = "already in your format - kept"
            summary["in_format"] += 1
        else:
            new_body = quick_convert(r, me) if use_quick else None
            how = "quick"
            if not new_body:
                if stop or ai_used >= ai_limit:
                    r["status"] = ("NOT converted - AI quota finished" if stop else
                                   "NOT converted - AI limit for this run reached")
                elif not (api_key and me["name"] and me["field"]):
                    r["status"] = "NOT converted - needs Gemini key + your details"
                else:
                    try:
                        new_body = convert_email(api_key, model, r, me)
                        how = "ai"
                        ai_used += 1
                    except FatalAPIError as e:
                        stop = True
                        r["status"] = "NOT converted - AI quota finished"
                        summary["error"] = str(e)
                    except Exception as e:
                        r["status"] = f"NOT converted - failed: {str(e)[:50]}"
            if new_body:
                r["body"] = new_body
                r["status"] = "converted to your format (no AI)" if how == "quick" \
                    else "converted to your format (AI)"
                summary[how] += 1
                progress[r["key"]] = {"status": "converted", "body": new_body,
                                      "email": r["email"]}
                persist()
            else:
                r["include"] = False
                summary["not_converted"] += 1
        r["subject"] = r["subject"] or "Prospective graduate student enquiry"
        if tick:
            tick((k + 1) / len(recs))
    return summary


# ---------------------------------------------------------------------------
# CV helper (optional auto-fill)
# ---------------------------------------------------------------------------
def read_cv_text(name, data):
    n = name.lower()
    if n.endswith(".pdf"):
        return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
    if n.endswith(".docx"):
        return "\n".join(p.text for p in Document(io.BytesIO(data)).paragraphs)
    return data.decode("utf-8", errors="ignore")


def autofill_from_cv():
    """Button callback: fill the 'My details' form from the uploaded CV."""
    ss = st.session_state
    cvf = get_file("cv")
    if not (cvf and ss.get("api_key")):
        ss["fill_msg"] = "Upload your CV and enter the Gemini key first."
        return
    try:
        system = "You extract structured data from academic CVs. Output ONLY valid JSON."
        user = f"""From this CV return JSON with keys: name, field (field of the bachelor's
degree), university_country (bachelor university with country), thesis_title (bachelor
thesis or final year project title, "" if none), thesis_methods (techniques/tools used),
thesis_summary (what it investigated, one short phrase). Do not invent anything.

CV:
{read_cv_text(cvf[0], cvf[1])[:12000]}"""
        d = parse_json(llm(ss["api_key"].strip(), ss["model"], system, user, 800))
        ss["d_name"] = d.get("name", "")
        ss["d_field"] = d.get("field", "")
        ss["d_uni"] = d.get("university_country", "")
        ss["d_title"] = d.get("thesis_title", "")
        ss["d_methods"] = d.get("thesis_methods", "")
        ss["d_summary"] = d.get("thesis_summary", "")
        ss["fill_msg"] = "Details filled from your CV - please check them."
    except Exception as e:
        ss["fill_msg"] = f"Could not read CV: {e}"


# ---------------------------------------------------------------------------
# Gmail drafts  (reliable bulk saving: reconnects, retries, checks Gmail reply)
# ---------------------------------------------------------------------------
def _drafts_folder(imap):
    typ, data = imap.list()
    for raw in data or []:
        line = raw.decode("utf-8", "ignore")
        if "\\Drafts" in line:
            return line.rsplit(' "/" ', 1)[-1].strip()
    return '"[Gmail]/Drafts"'


def _connect(gmail, app_pw):
    imap = imaplib.IMAP4_SSL("imap.gmail.com", timeout=120)
    imap.login(gmail, app_pw)
    return imap, _drafts_folder(imap)


def save_drafts(gmail, app_pw, rows, cv, transcript, progress_cb=None,
                reconnect_every=8, pause=1.5, retries=4):
    """cv / transcript = (filename, bytes). The transcript is attached ONLY to
    emails that say they attach it (row['need_tr']).
    Returns (list of saved row keys, list of error strings).
    - a fresh Gmail connection is opened every `reconnect_every` drafts
    - every draft is retried up to `retries` times
    - Gmail's reply is checked, so a draft is counted only if Gmail said OK"""
    imap, folder = _connect(gmail, app_pw)
    since_connect = 0
    saved_keys, errors = [], []

    for i, r in enumerate(rows):
        if since_connect >= reconnect_every:
            try:
                imap.logout()
            except Exception:
                pass
            try:
                imap, folder = _connect(gmail, app_pw)
                since_connect = 0
            except Exception as e:
                errors.append(f"{r['email']}: reconnect failed: {e}")
                if progress_cb:
                    progress_cb((i + 1) / len(rows))
                continue

        msg = EmailMessage()
        msg["From"] = gmail
        msg["To"] = r["email"]
        msg["Subject"] = r["subject"]
        msg.set_content(r["body"])
        files = [cv] + ([transcript] if r.get("need_tr") and transcript else [])
        for fname, fbytes in files:
            msg.add_attachment(fbytes, maintype="application",
                               subtype="octet-stream", filename=fname)
        raw = msg.as_bytes()

        last_err = None
        for attempt in range(retries):
            try:
                typ, data = imap.append(folder, "\\Draft",
                                        imaplib.Time2Internaldate(time.time()), raw)
                if typ != "OK":
                    raise RuntimeError(f"Gmail said {typ}: {data}")
                saved_keys.append(r["key"])
                since_connect += 1
                last_err = None
                break
            except Exception as e:
                last_err = e
                time.sleep(5 * (attempt + 1))
                try:
                    imap.logout()
                except Exception:
                    pass
                try:
                    imap, folder = _connect(gmail, app_pw)
                    since_connect = 0
                except Exception as e2:
                    last_err = e2
        if last_err is not None:
            errors.append(f"{r['email']}: {last_err}")

        if progress_cb:
            progress_cb((i + 1) / len(rows))
        time.sleep(pause)
    try:
        imap.logout()
    except Exception:
        pass
    return saved_keys, errors


def make_docx(rows):
    doc = Document()
    doc.add_heading("Professor Emails (final)", 0)
    for i, r in enumerate(rows, 1):
        doc.add_heading(f"{i}. {r['name'] or r['email']} - {r['uni']}".strip(" -"), 2)
        doc.add_paragraph(f"To: {r['email']}")
        doc.add_paragraph(f"Subject: {r['subject']}")
        doc.add_paragraph(r["body"])
        doc.add_page_break()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Word to Gmail Drafts", layout="wide")
st.title("Word file -> Gmail Drafts")
ss = st.session_state
saved = load_json(SETTINGS)
for _k in SAVE_KEYS:
    if _k in saved:
        ss.setdefault(_k, saved[_k])
ss.setdefault("model", "gemini-3.5-flash-lite")
ss.setdefault("d_degree", "Master's")
ss.setdefault("d_intake", "CSC 2027 scholarship intake")
ss.setdefault("recs", [])
ss.setdefault("summary", {})
progress = load_json(PROGRESS)


def persist():
    save_json(PROGRESS, progress)


def do_save(rows):
    """Save ALL given rows as Gmail drafts (one run) and remember them as done."""
    cv, tr = get_file("cv"), get_file("tr")
    if not (ss.get("gmail") and ss.get("app_pw")):
        st.error("Enter Gmail address and App Password in the sidebar.")
        return
    if not cv:
        st.error("Upload your CV in step 1 - it is attached to every draft.")
        return
    if any(r["need_tr"] for r in rows) and not tr:
        st.error(f"{sum(r['need_tr'] for r in rows)} emails say the academic transcript is "
                 "attached, but it is not uploaded. Upload it in step 1.")
        return
    if not rows:
        st.info("Nothing new to save.")
        return
    st.write(f"Saving {len(rows)} drafts to Gmail... keep this page open until it finishes.")
    bar = st.progress(0.0)
    try:
        ok_keys, errs = save_drafts(ss["gmail"].strip(), ss["app_pw"].replace(" ", ""),
                                    rows, cv, tr, bar.progress)
    except imaplib.IMAP4.error as ex:
        st.error(f"Gmail login failed: {ex}. Check the App Password.")
        return
    except Exception as ex:
        st.error(f"Could not save drafts: {ex}")
        return
    ok_set = set(ok_keys)
    for r in rows:
        if r["key"] in ok_set:
            progress[r["key"]] = {"status": "drafted", "email": r["email"]}
    persist()
    st.success(f"{len(ok_keys)} of {len(rows)} drafts saved to Gmail. They are remembered, "
               "so next time they are skipped. Open Gmail > Drafts, review, then send.")
    if errs:
        st.warning(f"{len(errs)} drafts failed. Click the same button again - already saved "
                   "ones are skipped, only these will be retried.")
    for er in errs:
        st.error(er)


with st.sidebar:
    st.header("Settings")
    st.text_input("Free Gemini API key (only needed for AI conversion)", type="password",
                  key="api_key", help="https://aistudio.google.com/apikey")
    st.text_input("Model", key="model")
    MIN_INTERVAL[0] = st.slider("Seconds between AI calls", 2.0, 20.0, 6.5)
    st.divider()
    st.text_input("Your Gmail address", key="gmail")
    st.text_input("Gmail App Password", type="password", key="app_pw",
                  help="Google Account > Security > App passwords")
    remember = st.checkbox("Remember my keys and details on this computer", value=True,
                           key="remember")
    if remember:
        data = {k: ss.get(k, "") for k in SAVE_KEYS}
        if any(data.values()):
            save_json(SETTINGS, data)
        st.caption(f"Saved in {SETTINGS.name} next to this script (plain text). "
                   "Never share that file.")
    elif SETTINGS.exists():
        SETTINGS.unlink()
    st.caption("Emails are only saved as DRAFTS. Nothing is ever sent automatically.")

# ---- Step 1: my details --------------------------------------------------
st.header("1. My details and attachments")
c1, c2 = st.columns(2)
with c1:
    st.file_uploader("My CV (attached to every draft)", type=["pdf", "docx"], key="cv_upl")
    st.file_uploader("Academic transcript (attached to emails that mention it)",
                     type=["pdf", "docx", "jpg", "png"], key="tr_upl")
    cvf, trf = get_file("cv"), get_file("tr")
    st.caption(f"CV in use: {cvf[0] if cvf else 'none'}   |   "
               f"Transcript in use: {trf[0] if trf else 'none'}  (saved on this computer)")
    st.button("Fill my details from the CV (Gemini)", on_click=autofill_from_cv)
    if ss.get("fill_msg"):
        st.info(ss["fill_msg"])
with c2:
    st.text_input("Full name", key="d_name")
    st.text_input("Bachelor's field (e.g. Software Engineering)", key="d_field")
    st.text_input("University, Country", key="d_uni")
    st.text_input("Thesis / final year project title", key="d_title")
    st.text_input("Methods / techniques used", key="d_methods")
    st.text_input("What it investigated (short)", key="d_summary")
    st.text_input("Degree you apply for", key="d_degree")
    st.text_input("Intake", key="d_intake")

# ---- Step 2/3: file + options -------------------------------------------
st.header("2. Word file with the professor emails")
wf = st.file_uploader("Upload the .docx (the same file every time - progress is remembered)",
                      type=["docx"], key="wf")
st.header("3. What should happen to the emails?")
mode = st.radio(
    "Choose",
    ["Keep the emails EXACTLY as they are in the Word file (no changes)",
     "Convert to MY format - but keep emails that are already in my format (recommended)",
     "Convert ALL emails to MY format (rewrite every email)"],
    index=1,
)
method = st.radio(
    "Conversion method",
    ["Quick (no AI, free) - AI only for emails the quick method cannot handle",
     "AI (Gemini) for every email"], index=0)
ai_limit = st.number_input("Max AI conversions in one run (protects your free quota)",
                           0, 5000, 500)
st.checkbox("Also save emails that have NO address (the To field stays empty - you type "
            "the address in Gmail later)", key="allow_blank", value=False)
with st.expander("Show my format"):
    st.code(TEMPLATE)


# ---- Step 4: prepare (and optionally save everything in one go) ----------
def run_prepare(save_all):
    if not wf:
        st.error("Upload the Word file first.")
        return
    recs = parse_docx(wf.getvalue())
    if not recs:
        st.error("No emails found. Each email must start with 'Dear ...'.")
        return
    trf_ = get_file("tr")
    me = {
        "name": ss["d_name"].strip(), "field": ss["d_field"].strip(),
        "university": ss["d_uni"].strip(), "thesis_title": ss["d_title"].strip(),
        "thesis_methods": ss["d_methods"].strip(),
        "thesis_summary": ss["d_summary"].strip(),
        "degree": ss["d_degree"].strip(), "intake": ss["d_intake"].strip(),
        "attach": "CV and academic transcript" if trf_ else "CV",
    }
    st.write("Preparing emails...")
    bar = st.progress(0.0)
    summ = prepare_records(
        recs, mode.startswith("Keep"), mode.startswith("Convert ALL"),
        not method.startswith("AI"), me, int(ai_limit),
        (ss.get("api_key") or "").strip(), ss["model"], progress, persist, bar.progress)
    ss.recs, ss.summary = recs, summ
    ss["last_lines"] = list(PARSE_LINES)
    for key in [k for k in ss.keys() if str(k).startswith("body_")]:
        del ss[key]
    if summ.get("error"):
        st.error(summ["error"])
    if save_all:
        allow = bool(ss.get("allow_blank"))
        ready = [(i, r) for i, r in enumerate(recs) if r["include"]]
        bad = [(i, r) for i, r in ready if not EMAIL_RE.fullmatch(r["email"].strip())]
        good = [r for i, r in ready if allow or EMAIL_RE.fullmatch(r["email"].strip())]
        if bad and not allow:
            st.warning(f"{len(bad)} emails have no address and were skipped. First ones: "
                       + "; ".join(f"#{i+1} ({r['body'][:35]!r})" for i, r in bad[:6])
                       + ". Tick 'Also save emails that have NO address' in step 3 to save "
                         "them anyway, or open 'Debug' in step 5 and send me what it shows.")
        do_save([{"key": r["key"], "name": r["name"], "uni": r["uni"],
                  "email": r["email"].strip(), "subject": r["subject"],
                  "body": r["body"], "need_tr": needs_transcript(r["body"])}
                 for r in good])


st.header("4. Read, prepare and save")
st.caption("ONE CLICK: reads the Word file, converts (or keeps) every email according to "
           "step 3, and saves ALL of them as Gmail drafts in one run.")
if st.button("Prepare + save ALL as Gmail drafts (one click)", type="primary"):
    run_prepare(save_all=True)
pc1, pc2 = st.columns([3, 1])
if pc2.button("Forget saved progress"):
    PROGRESS.unlink(missing_ok=True)
    progress.clear()
    ss.recs, ss.summary = [], {}
    st.info("Progress cleared - all emails will be treated as new.")
if pc1.button("Only read and prepare (let me review first)"):
    run_prepare(save_all=False)

# ---- Step 5: review + save ----------------------------------------------
if ss.recs:
    sm = ss.summary
    st.header("5. Result")
    ready_n = sum(1 for r in ss.recs if r["include"])
    m = st.columns(6)
    m[0].metric("Emails in file", sm.get("total", len(ss.recs)))
    m[1].metric("Converted now", sm.get("quick", 0) + sm.get("ai", 0))
    m[2].metric("Already in format / kept",
                sm.get("in_format", 0) + sm.get("kept", 0) + sm.get("converted_before", 0))
    m[3].metric("NOT converted yet", sm.get("not_converted", 0))
    m[4].metric("Saved to drafts earlier", sm.get("drafted_before", 0))
    m[5].metric("Ready to save", ready_n)
    if sm.get("not_converted", 0):
        st.warning(f"{sm['not_converted']} emails are NOT converted (AI limit or quota). "
                   "They are unticked. Click the one-click button again later "
                   "- saved/converted ones are skipped and only these will be converted.")
    table = [{"include": r["include"], "name": r["name"], "email": r["email"],
              "subject": r["subject"], "status": r["status"]} for r in ss.recs]
    edited = st.data_editor(table, use_container_width=True, hide_index=True,
                            disabled=["name", "status"], key="tbl")
    missing = [i + 1 for i, e in enumerate(edited)
               if e["include"] and not EMAIL_RE.fullmatch(e["email"].strip())]
    if missing and not ss.get("allow_blank"):
        st.warning(f"No valid email address for {len(missing)} ticked rows "
                   f"(first ones: {missing[:10]}). Fix the 'email' column or untick them.")
    seen, dups = set(), []
    for e in edited:
        k = e["email"].strip().lower()
        if k and k in seen:
            dups.append(k)
        seen.add(k)
    if dups:
        st.warning("Duplicate addresses: " + ", ".join(sorted(set(dups))))
    nomail = [i for i, r in enumerate(ss.recs)
              if r["include"] and not EMAIL_RE.fullmatch(r["email"].strip())]
    if nomail and ss.get("last_lines"):
        with st.expander(f"Debug: why {len(nomail)} emails have no address"):
            st.caption("Lines of the Word file just before the first emails without an address. "
                       "If the address is visible here, copy this box and send it to your helper.")
            for i in nomail[:2]:
                ln = ss.recs[i].get("dear_line", 0)
                st.code("\n".join(ss["last_lines"][max(0, ln - 10): ln + 2]))
    for i, r in enumerate(ss.recs):
        with st.expander(f"{i+1}. {r['name'] or r['email'] or 'email'} - {r['status']}"):
            st.text_area("Email text (you can edit)", r["body"], height=320, key=f"body_{i}")

    def final_rows():
        rows = []
        for i, (r, e) in enumerate(zip(ss.recs, edited)):
            if e["include"]:
                body = ss.get(f"body_{i}", r["body"])
                rows.append({"key": r["key"], "name": r["name"], "uni": r["uni"],
                             "email": e["email"].strip(), "subject": e["subject"].strip(),
                             "body": body, "need_tr": needs_transcript(body)})
        return rows

    rows_now = [r for r in final_rows()
                if ss.get("allow_blank") or EMAIL_RE.fullmatch(r["email"])]
    n_tr = sum(r["need_tr"] for r in rows_now)
    st.info(f"Every draft gets your CV. {n_tr} of {len(rows_now)} ticked emails mention the "
            "academic transcript, so it is attached to those.")
    cdl, csave = st.columns(2)
    cdl.download_button("Download final emails (.docx)", make_docx(final_rows()),
                        "final_emails.docx")
    first_n = csave.number_input("Save only the first N ticked emails (0 = all)", 0, 5000, 0)
    if csave.button("Save ALL ticked emails as Gmail drafts"):
        do_save(rows_now[: int(first_n)] if first_n else rows_now)
