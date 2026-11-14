"""
CSC Professor Outreach System  (FREE - uses Google Gemini free API)
--------------------------------------------------------------------
CV upload -> profile extraction -> crawl CSC university faculty directories ->
match research -> personalized emails -> all exported to one .docx file.

Run:  streamlit run app.py
"""
import concurrent.futures as cf
import hashlib
import imaplib
import io
import json
import os
import re
import sqlite3
import threading
import time
from email.message import EmailMessage
from urllib.parse import urljoin, urlparse

import requests
import streamlit as st
from bs4 import BeautifulSoup
from ddgs import DDGS
from docx import Document
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Your CSC university list (names from new_list_uni.docx). The broken links
# are NOT used - each university is searched by its website domain instead.
# Editable in the "Universities" tab (you may also paste a working faculty URL).
# ---------------------------------------------------------------------------
UNI_LIST = [
    ('Southern Medical University', 'smu.edu.cn'),
    ('Guangxi Medical University', 'gxmu.edu.cn'),
    ('Nanning Normal University', 'nnnu.edu.cn'),
    ('Guizhou Minzu University', 'gzmu.edu.cn'),
    ('Hainan Normal University', 'hainnu.edu.cn'),
    ('Hebei Medical University', 'hebmu.edu.cn'),
    ('Hebei University of Economics and Business', 'heuet.edu.cn'),
    ('Harbin Medical University', 'hrbmu.edu.cn'),

]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
BAD_TLDS = {"png", "jpg", "jpeg", "gif", "css", "js", "svg", "webp"}



# ---------------------------------------------------------------------------
# Local storage: settings (password etc.), CV files, professors database.
# Everything lives in the folder "csc_data" next to app.py (plain files on YOUR computer).
# ---------------------------------------------------------------------------
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "csc_data")
os.makedirs(DATA_DIR, exist_ok=True)
SETTINGS_F = os.path.join(DATA_DIR, "settings.json")
PROFILE_F = os.path.join(DATA_DIR, "profile.json")
EMAILS_F = os.path.join(DATA_DIR, "emails.json")
DRAFTED_F = os.path.join(DATA_DIR, "drafted.json")
UNIS_F = os.path.join(DATA_DIR, "universities.txt")
DB_F = os.path.join(DATA_DIR, "professors.db")
FILES_DIR = os.path.join(DATA_DIR, "my_files")
os.makedirs(FILES_DIR, exist_ok=True)


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path, obj):
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    except Exception:
        pass


def save_file(kind, name, data):
    """kind = 'cv' or 'tr'. Keep exactly one saved copy of each."""
    for fn in os.listdir(FILES_DIR):
        if fn.startswith(kind + "__"):
            os.remove(os.path.join(FILES_DIR, fn))
    safe = re.sub(r"[^\w.\- ]", "_", name)
    with open(os.path.join(FILES_DIR, f"{kind}__{safe}"), "wb") as f:
        f.write(data)


def load_file(kind):
    for fn in os.listdir(FILES_DIR):
        if fn.startswith(kind + "__"):
            with open(os.path.join(FILES_DIR, fn), "rb") as f:
                return fn.split("__", 1)[1], f.read()
    return "", None


_db_lock = threading.Lock()


def db():
    con = sqlite3.connect(DB_F, timeout=30, check_same_thread=False)
    con.execute("CREATE TABLE IF NOT EXISTS cands (key TEXT PRIMARY KEY, url TEXT, uni TEXT, "
                "text TEXT, emails TEXT)")
    con.execute("CREATE TABLE IF NOT EXISTS results (key TEXT, phash TEXT, status TEXT, "
                "data TEXT, PRIMARY KEY (key, phash))")
    return con


def cand_key(c):
    return hashlib.sha1((c["url"] + "|" + ",".join(sorted(c["emails"]))).encode()).hexdigest()


def db_add_cands(cands):
    """Store candidate pages. Returns how many were NEW."""
    new = 0
    with _db_lock:
        con = db()
        for c in cands:
            k = cand_key(c)
            cur = con.execute("INSERT OR IGNORE INTO cands VALUES (?,?,?,?,?)",
                              (k, c["url"], c["uni"], c["text"][:9000],
                               json.dumps(sorted(c["emails"]))))
            new += cur.rowcount
        con.commit()
        con.close()
    return new


def db_pending(phash, unis=None):
    """Pages that still need an AI judgement (never analysed, or failed last time)."""
    with _db_lock:
        con = db()
        rows = con.execute(
            "SELECT c.key, c.url, c.uni, c.text, c.emails FROM cands c LEFT JOIN results r "
            "ON r.key = c.key AND r.phash = ? WHERE r.key IS NULL OR r.status = 'fail'",
            (phash,)).fetchall()
        con.close()
    out = [{"key": k, "url": u, "uni": n, "text": t, "emails": set(json.loads(e))}
           for k, u, n, t, e in rows if not unis or n in unis]
    return out


def db_set_result(key, phash, status, data=None):
    with _db_lock:
        con = db()
        con.execute("INSERT OR REPLACE INTO results VALUES (?,?,?,?)",
                    (key, phash, status, json.dumps(data, ensure_ascii=False) if data else ""))
        con.commit()
        con.close()


def db_stats(phash):
    with _db_lock:
        con = db()
        total = con.execute("SELECT COUNT(*) FROM cands").fetchone()[0]
        st_ = dict(con.execute("SELECT status, COUNT(*) FROM results WHERE phash=? GROUP BY status",
                               (phash,)).fetchall())
        con.close()
    done = sum(st_.values()) - st_.get("fail", 0)
    return {"total": total, "pending": total - done, "ok": st_.get("ok", 0),
            "notprof": st_.get("notprof", 0), "notcs": st_.get("notcs", 0),
            "fail": st_.get("fail", 0)}


def db_ok_profs(phash):
    with _db_lock:
        con = db()
        rows = con.execute("SELECT data FROM results WHERE phash=? AND status='ok'",
                           (phash,)).fetchall()
        con.close()
    return [json.loads(r[0]) for r in rows if r[0]]


def profile_hash(profile):
    keep = {k: profile.get(k) for k in ("research_interests", "search_keywords")}
    return hashlib.sha1(json.dumps(keep, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:12]


CS_RE = re.compile(
    r"computer|software|informatic|information (system|technolog|engineer)|artificial|intelligen|"
    r"machine learning|deep learning|neural|data (science|mining|engineer|analy)|big data|"
    r"cloud|network|algorithm|programming|database|cyber|security|blockchain|internet of things|"
    r"\biot\b|\bai\b|\bnlp\b|vision|robot|embedded|distributed|compiler|operating system|"
    r"human[- ]computer|multimedia|pattern recognition|"
    r"计算机|软件|信息|人工智能|智能|数据|算法|网络|机器学习|深度学习|云计算|区块链|物联网|"
    r"自动化|控制|通信|电子|模式识别|视觉|大数据", re.I)
CS_HOST_RE = re.compile(r"(^|[./-])(cs|scs|cse|scse|sse|se|soft|software|sei|jsj|rjxy|ai|cist|cic|"
                        r"ice|info|it|cyber|ds|bigdata)([./-]|\d|$)", re.I)


def is_cs_page(c):
    """FREE local filter: skip pages clearly unrelated to computing (no AI call needed)."""
    host = urlparse(c["url"]).netloc.lower()
    return bool(CS_RE.search(c["text"]) or CS_HOST_RE.search(host))


# ---------------------------------------------------------------------------
# CV + AI helpers
# ---------------------------------------------------------------------------
def read_cv(uploaded) -> str:
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in reader.pages)
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    return data.decode("utf-8", errors="ignore")


class FatalAPIError(Exception):
    """Wrong key / wrong model name etc. - no point retrying."""


class TransientAIError(Exception):
    """A temporary AI problem (rate limit, bad output): this page is retried next run."""


_last_call = [0.0]
MIN_INTERVAL = [6.5]  # seconds between calls (free tier is rate limited)


def llm(client, model, system, user, max_tokens=1500):
    """Calls the FREE Gemini API (Google AI Studio). `client` is the API key."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    body = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "user", "parts": [{"text": user}]}],
        "generationConfig": {
            "maxOutputTokens": max_tokens + 2000,
            "temperature": 0.7,
            "responseMimeType": "application/json",
        },
    }
    for attempt in range(6):
        wait = MIN_INTERVAL[0] - (time.time() - _last_call[0])
        if wait > 0:
            time.sleep(wait)
        _last_call[0] = time.time()
        r = requests.post(url, json=body, headers={"x-goog-api-key": client}, timeout=90)
        if r.status_code == 200:
            data = r.json()
            try:
                parts = data["candidates"][0]["content"]["parts"]
                return "".join(p.get("text", "") for p in parts)
            except (KeyError, IndexError):
                raise ValueError("Empty response (blocked or no output)")
        if r.status_code == 429 and re.search(r"per\s*day|daily|PerDay", r.text, re.I):
            raise FatalAPIError(
                "Gemini DAILY free quota is finished. Stop now and run again "
                "tomorrow (quota resets daily), or lower the number of emails.")
        if r.status_code in (429, 500, 503):  # rate limit / busy -> wait, retry
            m = re.search(r'"retryDelay":\s*"(\d+(?:\.\d+)?)s"', r.text)
            time.sleep(float(m.group(1)) + 2 if m else 20 * (attempt + 1))
            continue
        raise FatalAPIError(f"Gemini error {r.status_code}: {r.text[:300]}")
    raise ValueError("Gemini kept refusing (rate limit). Try again later.")


def parse_json(text):
    text = text.strip()
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        raise ValueError("No JSON found in model output")
    return json.loads(m.group(0))


def extract_profile(client, model, cv_text):
    system = "You extract structured data from academic CVs. Output ONLY valid JSON."
    user = f"""From this CV extract JSON with keys:
name, email, phone, nationality, current_degree, university, gpa,
target_degree (Master/PhD - guess from CV if not stated),
research_interests (list), search_keywords (list of 6-10 specific English
keywords for finding matching professors),
search_keywords_zh (the same keywords translated to Simplified Chinese, as
used on Chinese faculty pages), skills (list),
publications (list of short strings), projects (list of short strings),
bachelor_field (the field of the bachelor's degree, e.g. "Software Engineering"),
country (applicant's home country), thesis_title (bachelor thesis or final year
project title, "" if none), thesis_methods (techniques/tools/methodologies used
in it), thesis_summary (1-2 sentences on what it investigated),
experience_summary (2-3 sentences).
Use "" or [] if not present. Do not invent anything.

CV:
{cv_text[:12000]}"""
    return parse_json(llm(client, model, system, user, 2000))


# ---------------------------------------------------------------------------
# Web crawling helpers (no AI needed here)
# ---------------------------------------------------------------------------
def root_domain(url_or_host: str) -> str:
    host = urlparse(url_or_host).netloc or url_or_host
    host = host.split(":")[0].lower()
    parts = host.split(".")
    if len(parts) >= 3 and parts[-2] in ("edu", "ac", "com", "org", "gov"):
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


ERRS = {}          # url -> last download error (used by the Diagnose tool)
VERIFY_SSL = [True]


def get_html(url, timeout=15, retries=1):
    """Download a page. `retries` extra attempts make results stable on flaky networks."""
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, verify=VERIFY_SSL[0])
            r.raise_for_status()
            if "charset" not in r.headers.get("content-type", "").lower():
                r.encoding = r.apparent_encoding  # many Chinese sites omit charset
            ERRS.pop(url, None)
            return r.text
        except Exception as e:
            ERRS[url] = f"{type(e).__name__}: {str(e)[:140]}"
            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
    return ""


RENDER_JS = [False]
_pw_lock = threading.Lock()


def render_js(url, timeout=20):
    """Optional: open the page in a real headless browser (for sites whose professor list
    is loaded by JavaScript). Needs:  pip install playwright  &  playwright install chromium"""
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        ERRS[url] = "playwright is not installed"
        return ""
    with _pw_lock:
        try:
            with sync_playwright() as p:
                b = p.chromium.launch()
                pg = b.new_page(user_agent=HEADERS["User-Agent"], ignore_https_errors=not VERIFY_SSL[0])
                pg.goto(url, timeout=timeout * 1000, wait_until="networkidle")
                html = pg.content()
                b.close()
                return html
        except Exception as e:
            ERRS[url] = f"browser: {type(e).__name__}: {str(e)[:100]}"
            return ""


CATEGORY_RE = re.compile(r"教授|副教授|讲师|助理|研究员|教师|师资|导师|全部|更多|more|professor|"
                         r"associate|assistant|lecturer|faculty|staff|people|teacher", re.I)


def norm_text(txt):
    txt = txt.replace("＠", "@")
    txt = re.sub(r"\s*[\[\(（]\s*at\s*[\]\)）]\s*", "@", txt, flags=re.I)
    txt = re.sub(r"\s*[\[\(（]\s*dot\s*[\]\)）]\s*", ".", txt, flags=re.I)
    return txt


def html_to_text(html):
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "nav", "footer"]):
        t.decompose()
    return norm_text(re.sub(r"\s+", " ", soup.get_text(" ")))


def extract_emails(html, text):
    found = set(EMAIL_RE.findall(text))
    for m in re.findall(r'mailto:([^"\'?>\s]+)', html, re.I):
        found.update(EMAIL_RE.findall(m))
    return {e for e in found if e.rsplit(".", 1)[-1].lower() not in BAD_TLDS}


def email_blocks(html, url, uni):
    """Directory pages often list name + title + email in one row.
    Take the smallest HTML block that contains exactly one email."""
    soup = BeautifulSoup(html, "html.parser")
    best = {}
    for el in soup.find_all(["tr", "li", "dd", "dl", "p", "div"]):
        t = norm_text(re.sub(r"\s+", " ", el.get_text(" ", strip=True)))
        if not 30 <= len(t) <= 1500:
            continue
        ems = extract_emails(str(el), t)
        if len(ems) != 1:
            continue
        e = next(iter(ems))
        if e not in best or len(t) < len(best[e]):
            best[e] = t
    return [{"url": url, "uni": uni, "text": t, "emails": {e}} for e, t in best.items()]


def _name_like(txt):
    if re.fullmatch(r"[\u4e00-\u9fff·]{2,4}", txt):
        return True
    words = txt.split()
    return 2 <= len(words) <= 4 and all(re.fullmatch(r"[A-Za-z.\-']+", w) for w in words)


def _is_pagination(txt, href):
    if txt in ("下一页", "下页", "Next", "next", ">", "›", "»", "尾页", "末页"):
        return True
    return bool(txt.isdigit() and len(txt) <= 2)


def crawl_directory(url, max_links=60, max_list_pages=4, uni=""):
    """Open a faculty directory, follow pagination, return links to
    individual profile pages (name-like link texts first)."""
    base = root_domain(url)
    queue, seen_lists, links, blocks = [url], set(), {}, []
    pages = 0
    while queue and pages < max_list_pages:
        u = queue.pop(0)
        if u in seen_lists:
            continue
        seen_lists.add(u)
        html = get_html(u, retries=2)
        pages += 1
        if RENDER_JS[0] and html.lower().count("<a ") < 25:
            html = render_js(u) or html
        if not html:
            continue
        blocks += email_blocks(html, u, uni)
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.startswith(("javascript", "#", "tel:", "mailto:")):
                continue
            full = urljoin(u, href).split("#")[0]
            p = urlparse(full)
            if p.scheme not in ("http", "https") or root_domain(full) != base:
                continue
            if re.search(r"\.(pdf|docx?|xlsx?|jpe?g|png|zip|rar)$", p.path, re.I):
                continue
            txt = a.get_text(strip=True)
            if _is_pagination(txt, full):
                if full not in seen_lists:
                    queue.append(full)
                continue
            if (CATEGORY_RE.search(txt) and len(txt) <= 14 and full not in seen_lists
                    and full not in queue and len(queue) < 12):
                queue.append(full)      # category page (Professors / Associate Professors ...)
            if 1 < len(txt) <= 40 and full.rstrip("/") != u.rstrip("/"):
                links.setdefault(full, txt)
    ordered = sorted(links, key=lambda l: 0 if _name_like(links[l]) else 1)
    return ordered[:max_links], blocks[:300]


def fetch_candidate(url, uni):
    """Download a page; keep it only if it looks like ONE person's page
    (has an email, but not dozens like a list page)."""
    html = get_html(url, retries=2)
    if not html:
        return None
    text = html_to_text(html)
    emails = extract_emails(html, text)
    if not emails or len(emails) > 6:
        return None
    return {"url": url, "uni": uni, "text": text[:9000], "emails": emails}


SEARCH_STATE = {"fails": 0, "last_error": ""}


def ddg_text(query, n):
    """One web search with a HARD time limit (so the app can never hang)."""
    def run():
        with DDGS(timeout=10) as d:
            return list(d.text(query, max_results=n))
    ex = cf.ThreadPoolExecutor(1)
    fut = ex.submit(run)
    try:
        res = fut.result(timeout=25)
        SEARCH_STATE["fails"] = 0
        return res
    except Exception as e:
        SEARCH_STATE["fails"] += 1
        SEARCH_STATE["last_error"] = f"{type(e).__name__}: {str(e)[:150] or 'timeout'}"
        return []
    finally:
        ex.shutdown(wait=False, cancel_futures=True)


def search_fallback(domain, kws_en, kws_zh, per_query, budget=70):
    """Web search restricted to a university domain (English + Chinese queries)."""
    queries = [f"site:{domain} {kw} professor email" for kw in kws_en[:3]]
    queries += [f"site:{domain} {kw} 教授 邮箱" for kw in kws_zh[:2]]
    urls, start = [], time.time()
    for q in queries:
        if time.time() - start > budget or SEARCH_STATE["fails"] >= 8:
            break
        for r in ddg_text(q, per_query):
            u = r.get("href") or r.get("url")
            if u and u not in urls:
                urls.append(u)
        time.sleep(2.5)
    return urls


FACULTY_RE = re.compile(r"faculty|staff|people|teacher|professor|师资|教师|教授|导师|szdw|jsdw|directory", re.I)
DEPT_RE = re.compile(r"计算机|软件|信息|人工智能|网络空间|电子|自动化|computer|software|informat|artificial|intelligen|cyber|\bcs\b|\bai\b|electronic|automation", re.I)
SUBS = ["cs", "scs", "cse", "scse", "sse", "se", "soft", "software", "sei", "jsj",
        "rjxy", "it", "info", "ai", "cist", "ice", "cic", "cyber"]


def _links(html, base, root):
    out = []
    for a in BeautifulSoup(html, "html.parser").find_all("a", href=True):
        h = a["href"].strip()
        if h.startswith(("javascript", "#", "mailto:", "tel:")):
            continue
        full = urljoin(base, h).split("#")[0]
        if urlparse(full).scheme in ("http", "https") and root_domain(full) == root:
            out.append((a.get_text(" ", strip=True), full))
    return out


def _fetch_home(host):
    for scheme in ("https", "http"):
        h = get_html(f"{scheme}://{host}/", timeout=8)
        if h:
            return f"{scheme}://{host}/", h
    return None, ""


def discover_directories(domain, max_dirs=4):
    """Find faculty-list pages by browsing the university's own website
    (no search engine): homepage -> computer/software departments -> faculty links."""
    hosts = [f"www.{domain}", domain] + [f"{p}.{domain}" for p in SUBS]
    homes = {}
    with cf.ThreadPoolExecutor(6) as ex:
        for host, res in zip(hosts, ex.map(_fetch_home, hosts)):
            if res[1]:
                homes[host] = res
    for host in (f"www.{domain}", domain):          # departments linked from the main page
        if host in homes:
            base, html = homes[host]
            depts = []
            for txt, link in _links(html, base, domain):
                if DEPT_RE.search(txt) and link not in depts:
                    depts.append(link)
            for link in depts[:8]:
                h = urlparse(link).netloc
                if h not in homes:
                    dh = get_html(link, timeout=8)
                    if dh:
                        homes[h] = (link, dh)
            break
    dirs = []
    for host, (base, html) in homes.items():
        for txt, link in _links(html, base, domain):
            if (FACULTY_RE.search(txt) or FACULTY_RE.search(link)) and link not in dirs:
                dirs.append(link)
    main = {f"www.{domain}", domain}
    dirs.sort(key=lambda l: 1 if urlparse(l).netloc in main else 0)  # department sites first
    return dirs[:max_dirs]


def direct_crawl(uni, domain, per_uni):
    links, blocks = [], []
    dirs = discover_directories(domain, 6)
    for d in dirs:
        l, b = crawl_directory(d, per_uni, 8, uni)
        links += l
        blocks += b
    return list(dict.fromkeys(links))[: per_uni * 3], blocks[:600], len(dirs)


def resolve_domain(name):
    """Find a university's website domain from its name (web search)."""
    cn = ""
    for r in ddg_text(f"{name} China official website", 8):
        u = r.get("href") or r.get("url") or ""
        host = urlparse(u).netloc.lower()
        if host.endswith((".edu.cn", ".ac.cn")):
            return root_domain(u)
        if host.endswith(".cn") and not cn:
            cn = root_domain(u)
    return cn


def local_score(text, kws):
    """Cheap keyword overlap (no AI) used to decide which pages deserve an AI call."""
    t, s = text.lower(), 0.0
    for kw in kws:
        kw = (kw or "").lower().strip()
        if not kw:
            continue
        if kw in t:
            s += 2
        else:
            s += 0.5 * sum(1 for w in re.findall(r"[a-z]{4,}", kw) if w in t)
    return s


def analyze_page(client, model, profile, uni, url, page_text):
    """Return professor dict, "notprof", or raise TransientAIError (retry later).
    Email must literally appear on the page."""
    emails_on_page = set(EMAIL_RE.findall(page_text))
    if not emails_on_page:
        return "notprof"
    system = "You analyze faculty web pages. Output ONLY valid JSON."
    user = f"""Student research interests: {profile.get('research_interests')}
Student keywords: {profile.get('search_keywords')}

Web page from {uni} ({url}):
{page_text}

Emails found on page: {sorted(emails_on_page)}

If this page is about ONE specific professor/faculty member, return JSON:
{{"is_professor": true, "name": "...", "title": "...", "department": "...",
"email": "<one email from the list above that belongs to this professor>",
"research_interests": ["..."], "recent_work": ["specific paper/project titles
mentioned on the page"], "match_score": 0-10, "match_reason": "1-2 sentences",
"cs_related": true/false, "cs_field": "short name of the area",
"recruiting": "international" | "students" | "no" | "unknown",
"recruit_evidence": "short phrase copied from the page, or empty",
"wechat": "WeChat ID (微信/WeChat) copied EXACTLY as written on this page, or empty"}}
Never guess or invent a WeChat ID; leave it empty unless the page itself shows one.
"recruiting" rules (use ONLY what the page says):
- "international" = page explicitly mentions recruiting/welcoming international,
  foreign or overseas students, or 留学生 / CSC scholarship students.
- "students" = page says the professor is recruiting master/PhD students,
  has openings, or 招生/招收研究生, without mentioning international students.
- "no" = page says the professor is NOT recruiting.
- "unknown" = page says nothing about recruiting.
"cs_related" = true if the professor works in ANY of these areas (even if it does
NOT match the student's interests): software engineering (requirements, design,
architecture, testing, maintenance, DevOps, code analysis, program repair, formal
methods, software quality, open source), computer science in general, artificial
intelligence, machine learning, deep learning, NLP, large language models, computer
vision, multimodal learning, robotics/intelligent systems, data science, data mining,
big data, databases, knowledge graphs, cloud/edge/fog computing, distributed and
parallel/HPC systems, mobile computing and apps, web technologies, Internet of Things,
embedded/cyber-physical systems, operating systems, computer networks and network
software, blockchain, programming languages and compilers, cyber/software security and
privacy, human-computer interaction, information systems, game development, multimedia.
Set false ONLY for clearly non-computing fields (e.g. chemistry, law, pure physics,
civil engineering, medicine, literature).
"cs_field" = the one best-fitting area name from the list above (or similar).
Write names and research in English (translate from Chinese if needed).
Only use information present on the page. Never invent an email."""
    # Network / quota problems are raised (never hidden), so a failure is not mistaken
    # for "this page is not a professor".
    try:
        d = parse_json(llm(client, model, system, user, 900))
    except (FatalAPIError, requests.RequestException):
        raise
    except ValueError as e:
        if "rate limit" in str(e) or "Empty response" in str(e):
            raise TransientAIError(str(e))
        raise TransientAIError("Bad AI answer: " + str(e)[:80])
    if not d.get("is_professor"):
        return "notprof"
    if d.get("email") not in emails_on_page:  # anti-hallucination check
        return "notprof"
    if "cs_related" not in d:
        d["cs_related"] = d.get("se_related", True)
        d["cs_field"] = d.get("cs_field") or d.get("se_field", "")
    d["university"] = uni
    d["url"] = url
    return d


def _salutation(prof):
    t = (prof.get("title") or "").lower()
    if "associate" in t:
        role = "Associate Professor"
    elif "assistant" in t:
        role = "Assistant Professor"
    elif "lecturer" in t:
        role = "Lecturer"
    else:
        role = "Professor"
    return f"{role} {prof.get('name', '').strip()}".strip()


SAMPLE_EMAIL = """\
My name is Wasim Akram Janyaro, and I have completed my B.E. in Software Engineering from Mehran University of Engineering & Technology (MUET), Pakistan, with a CGPA of 3.40/4.00.

I am interested in pursuing a Master's degree at Harbin Institute of Technology (Shenzhen) through the Chinese Government Scholarship (CSC) 2027. My research interests include artificial intelligence, machine learning, deep learning, multimodal learning, and large language models.

For my Final Year Project (FYP), I developed GamerConnect, an all-in-one platform for gamers using Flutter, Dart, and Firebase. The platform included a marketplace, an AI chatbot, a gaming news module, and a community question-and-answer feature. This project strengthened my software development skills and encouraged me to explore intelligent applications.

I have also worked on machine learning projects involving medical image classification and predictive modelling. Your research on multimodal representation learning, multimodal large models, and multimodal agent applications particularly interests me. My experience developing an AI chatbot and working on machine learning projects has encouraged me to explore intelligent systems that can process and integrate information from multiple modalities.

I would be grateful if you would consider supervising me for a Master's degree under the CSC 2027 scholarship. I noticed that your research group is recruiting students for the 2027 cohort, and I would like to ask whether international applicants can also be considered for Master's admission through the CSC scholarship."""


TEMPLATE_THESIS = """\
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


def write_email_thesis(client, model, profile, prof, used_openings, extra,
                       csc_year="2027", attach_text="CV"):
    """Format 2 (thesis format). Fixed skeleton in code, AI fills 5 paragraphs."""
    target = profile.get("target_degree") or "Master"
    degree = "PhD" if "phd" in target.lower() or "doctor" in target.lower() else "Master's"
    system = (
        "You write emails from a prospective graduate student to a professor, "
        "following the applicant's template closely. Output ONLY valid JSON "
        '{"subject": "...", "p1": "...", "p2": "...", "p3": "...", "p4": "...", '
        '"p5": "..."}.'
    )
    user = f"""APPLICANT PROFILE (the only facts you may use about the applicant):
{json.dumps(profile, indent=1, ensure_ascii=False)[:4000]}

PROFESSOR (the only facts you may use about the professor):
{json.dumps({k: prof.get(k) for k in ['name','title','department','university','research_interests','recent_work','match_reason']}, indent=1, ensure_ascii=False)}

APPLICANT'S TEMPLATE (keep its wording and order; fill the placeholders):
\"\"\"
{TEMPLATE_THESIS}
\"\"\"

Write these paragraphs for Professor {prof.get('name')} at {prof.get('university')}:
p1 = "My name is <full name>, and I have completed my Bachelor's degree in
     <field> from <university, country>."  (use "am currently pursuing" if the
     profile shows the degree is not finished)
p2 = "During my Bachelor's studies, I conducted my research thesis on the topic
     "<title>." In this research, I used <methods> to investigate <description>."
     Use thesis_title/thesis_methods/thesis_summary from the profile. If the
     profile has no thesis, use the final year project or main research project
     and call it that (do not invent a thesis).
p3 = "After reviewing your academic profile and research publications, I found
     that your research on <1-2 SPECIFIC topics from the professor facts> is
     closely related to my academic background and research interests." Then ONE
     short sentence linking a specific part of the applicant's work to that topic.
p4 = "Therefore, I am highly interested in pursuing my {degree} degree under your
     supervision at {prof.get('university')}. I believe that your academic guidance
     and research expertise would greatly contribute to my academic and research
     development."
p5 = "I would be grateful if you could consider my application and provide me with
     an acceptance letter (if you are willing to supervise me) for the upcoming
     CSC {csc_year} scholarship intake."
Rules:
- Do NOT write greeting, "I hope this email finds you well", the attachment
  sentence, thanks, or signature - those are added automatically.
- Keep the template wording; only change what is needed. NO invented facts.
- Small wording variations between emails are fine; do not reuse the phrasing of
  these earlier emails in p3: {used_openings[-6:]}
- subject: specific, under 12 words, mentioning the professor's topic.
{extra}"""
    if prof.get("group") == "B":
        user += (
            "\nIMPORTANT: this professor's research does NOT closely match the "
            "applicant's. In p3 do NOT say it is closely related. Say honestly that "
            "after reviewing the professor's work on <topic>, the applicant found it "
            "interesting and is keen to broaden into this area using their "
            "<real transferable skills>."
        )
    d = parse_json(llm(client, model, system, user, 1500))
    paras = [d.get(k, "").strip() for k in ("p1", "p2", "p3", "p4", "p5")]
    body = (
        f"Dear {_salutation(prof)},\n\nI hope this email finds you well.\n\n"
        + "\n\n".join(p for p in paras if p)
        + f"\n\nI have attached my {attach_text} for your kind consideration.\n\n"
        "Thank you very much for your time and consideration. I look forward to "
        f"hearing from you.\n\nRegards,\n\n{profile.get('name', '')}"
    )
    return {"subject": d.get("subject", "Prospective graduate student enquiry"), "body": body}



def _topics(prof, n=2):
    ri = [t for t in (prof.get("research_interests") or []) if t]
    if not ri:
        ri = [prof.get("cs_field") or prof.get("department") or "your research area"]
    ri = ri[:n]
    return ri[0] if len(ri) == 1 else ", ".join(ri[:-1]) + " and " + ri[-1]


def _uni_line(profile):
    u = ", ".join(x for x in [profile.get("university"), profile.get("country")] if x)
    return u or "my university"


def write_email_template(profile, prof, csc_year="2027", attach_text="CV", fmt="thesis"):
    """FREE: no AI call. Uses your fixed format + this professor's real research topics."""
    target = profile.get("target_degree") or "Master"
    degree = "PhD" if "phd" in target.lower() or "doctor" in target.lower() else "Master's"
    topics = _topics(prof)
    matched = prof.get("group") == "A"
    mine = ", ".join((profile.get("research_interests") or [])[:4])
    done = "have completed" if re.search(r"bachelor|b\.?e|b\.?sc|b\.?s\b|master|m\.?sc",
                                         str(profile.get("current_degree", "")), re.I) \
        and not re.search(r"pursuing|currently|student", str(profile.get("current_degree", "")), re.I) \
        else "am currently pursuing"
    field = profile.get("bachelor_field") or profile.get("current_degree") or "my field"
    subj = f"CSC {csc_year} {degree} application - {topics}"[:110]
    if fmt == "thesis":
        p1 = (f"My name is {profile.get('name','')}, and I {done} my Bachelor's degree in "
              f"{field} from {_uni_line(profile)}.")
        if profile.get("thesis_title"):
            p2 = (f"During my Bachelor's studies, I conducted my research thesis on the topic "
                  f"\"{profile['thesis_title']}.\" In this research, I used "
                  f"{profile.get('thesis_methods') or 'relevant techniques'} to investigate "
                  f"{profile.get('thesis_summary') or 'the topic'}")
        else:
            proj = (profile.get("projects") or [""])[0]
            p2 = f"During my Bachelor's studies, I worked on {proj}" if proj else ""
        p2 = p2.rstrip(".") + "." if p2 else ""
        if matched:
            p3 = (f"After reviewing your academic profile and research publications, I found that "
                  f"your research on {topics} is closely related to my academic background and "
                  f"research interests.")
        else:
            p3 = (f"After reviewing your academic profile and research publications, I found your "
                  f"research on {topics} very interesting, and I am keen to broaden my knowledge "
                  f"into this area using the software development and research skills I have built.")
        p4 = (f"Therefore, I am highly interested in pursuing my {degree} degree under your "
              f"supervision at {prof.get('university')}. I believe that your academic guidance and "
              f"research expertise would greatly contribute to my academic and research development.")
        p5 = (f"I would be grateful if you could consider my application and provide me with an "
              f"acceptance letter (if you are willing to supervise me) for the upcoming CSC "
              f"{csc_year} scholarship intake.")
        paras = [p for p in (p1, p2, p3, p4, p5) if p]
        body = (f"Dear {_salutation(prof)},\n\nI hope this email finds you well.\n\n"
                + "\n\n".join(paras)
                + f"\n\nI have attached my {attach_text} for your kind consideration.\n\n"
                "Thank you very much for your time and consideration. I look forward to hearing "
                f"from you.\n\nRegards,\n\n{profile.get('name','')}")
        return {"subject": subj, "body": body}
    # Format 1
    gpa = f", with a CGPA of {profile['gpa']}" if profile.get("gpa") else ""
    projs = [p for p in (profile.get("projects") or []) if p]
    p1 = f"My name is {profile.get('name','')}, and I {done} my {profile.get('current_degree') or field} from {_uni_line(profile)}{gpa}."
    p2 = (f"I am interested in pursuing a {degree} degree at {prof.get('university')} through the "
          f"Chinese Government Scholarship (CSC) {csc_year}. My research interests include {mine}.")
    p3 = (f"My main project was {projs[0].rstrip('.')}." if projs else "")
    link = (f"Your research on {topics} particularly interests me, and my background connects well with it."
            if matched else
            f"Your research on {topics} interests me, and I am keen to broaden into this area using my "
            f"software development and machine learning skills.")
    p4 = (f"I have also worked on {projs[1]}. " if len(projs) > 1 else "") + link
    ask = ("I noticed that your research group is recruiting students, and I would like to ask whether "
           "international applicants can also be considered through the CSC scholarship."
           if prof.get("recruit_evidence") else
           "I would like to ask whether international applicants can be considered through the CSC scholarship.")
    p5 = f"I would be grateful if you would consider supervising me for a {degree} degree under the CSC {csc_year} scholarship. {ask}"
    sig = [profile.get("name", ""), profile.get("current_degree", ""),
           profile.get("university", ""), profile.get("email", "")]
    body = (f"Dear {_salutation(prof)},\n\nI hope you are doing well.\n\n"
            + "\n\n".join(p for p in (p1, p2, p3, p4, p5) if p)
            + f"\n\nI have attached my {attach_text} with this email for your kind reference.\n\n"
            "Thank you for your time and consideration. I look forward to your response.\n\n"
            "Best regards,\n" + "\n".join(x for x in sig if x))
    return {"subject": subj, "body": body}


def write_email(client, model, profile, prof, used_openings, extra,
                csc_year="2027", attach_text="CV", fmt="thesis"):
    if fmt == "thesis":
        return write_email_thesis(client, model, profile, prof, used_openings,
                                  extra, csc_year, attach_text)
    return write_email_projects(client, model, profile, prof, used_openings,
                                extra, csc_year, attach_text)


def write_email_projects(client, model, profile, prof, used_openings, extra,
                         csc_year="2027", attach_text="CV"):
    """Format 1 (projects + CSC format). Fixed skeleton in code, AI fills paragraphs."""
    target = profile.get("target_degree") or "Master"
    degree = "PhD" if "phd" in target.lower() or "doctor" in target.lower() else "Master's"
    system = (
        "You write emails from a prospective graduate student to a professor, "
        "copying the structure and tone of the applicant's own sample email. "
        'Output ONLY valid JSON {"subject": "...", "p1": "...", "p2": "...", '
        '"p3": "...", "p4": "...", "p5": "..."}.'
    )
    user = f"""APPLICANT PROFILE (the only facts you may use about the applicant):
{json.dumps(profile, indent=1, ensure_ascii=False)[:4000]}

PROFESSOR (the only facts you may use about the professor):
{json.dumps({k: prof.get(k) for k in ['name','title','department','university','research_interests','recent_work','match_reason','recruiting','recruit_evidence']}, indent=1, ensure_ascii=False)}

APPLICANT'S OWN SAMPLE EMAIL (follow this paragraph structure exactly; do NOT
copy its facts - use the facts from the profile above):
\"\"\"
{SAMPLE_EMAIL}
\"\"\"

Write the five paragraphs for Professor {prof.get('name')} at {prof.get('university')}:
p1 = "My name is ..., and I have completed/am pursuing <degree> from <university>,
     with a CGPA of ..." (omit CGPA if not in the profile; use completed vs
     pursuing correctly)
p2 = interested in pursuing a {degree} degree at {prof.get('university')} through the
     Chinese Government Scholarship (CSC) {csc_year}, then "My research interests
     include ..." (choose applicant interests that relate to this professor)
p3 = the applicant's best project (e.g. final year project) with technologies and
     what it strengthened
p4 = other relevant projects/experience, then "Your research on <2-3 SPECIFIC topics
     from the professor facts> particularly interests me", then how the applicant's
     experience connects to it
p5 = "I would be grateful if you would consider supervising me for a {degree}
     degree under the CSC {csc_year} scholarship." Then: if recruit_evidence is
     not empty, say you noticed the research group is recruiting students (for the
     {csc_year} cohort only if the page said so) and ask whether international
     applicants can be considered through CSC. If recruit_evidence is empty, do NOT
     claim they are recruiting; just ask whether international applicants can be
     considered through CSC.
Rules:
- Do NOT write greeting, "I hope you are doing well", attachments sentence,
  thanks, or signature - those are added automatically.
- Each paragraph 2-4 sentences, plain text, no flattery, NO invented facts.
- Vary the wording from earlier emails (do not reuse these phrasings):
  {used_openings[-6:]}
- subject: specific, under 12 words, e.g. "CSC {csc_year} {degree} application - <professor's topic>".
{extra}"""
    if prof.get("group") == "B":
        user += (
            "\nIMPORTANT: this professor's research does NOT closely match the "
            "applicant's. Do NOT pretend it does. In p4 mention the professor's work "
            "honestly and connect it only through the applicant's REAL transferable "
            "skills, saying the applicant is keen to broaden into this area. In p2 list "
            "the applicant's real interests plus interest in this area, without claiming "
            "prior research in it."
        )
    d = parse_json(llm(client, model, system, user, 1500))
    sig = [profile.get("name", ""), profile.get("current_degree", ""),
           profile.get("university", ""), profile.get("email", "")]
    paras = [d.get(k, "").strip() for k in ("p1", "p2", "p3", "p4", "p5")]
    body = (
        f"Dear {_salutation(prof)},\n\nI hope you are doing well.\n\n"
        + "\n\n".join(p for p in paras if p)
        + f"\n\nI have attached my {attach_text} with this email for your kind reference.\n\n"
        "Thank you for your time and consideration. I look forward to your response.\n\n"
        "Best regards,\n" + "\n".join(x for x in sig if x)
    )
    return {"subject": d.get("subject", "CSC scholarship enquiry"), "body": body}


def _drafts_folder(imap):
    """Find the Gmail Drafts folder name (works in any Gmail language)."""
    typ, data = imap.list()
    for raw in data or []:
        line = raw.decode("utf-8", "ignore")
        if "\\Drafts" in line:
            return line.rsplit(' "/" ', 1)[-1].strip()
    return '"[Gmail]/Drafts"'


def save_drafts(gmail, app_pw, rows, cv_name, cv_bytes, progress_cb=None, extra_files=(),
                on_saved=None):
    """Save emails as Gmail DRAFTS (nothing is sent). Returns (saved, errors)."""
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(gmail, app_pw)
    folder = _drafts_folder(imap)
    saved, errors = 0, []
    for i, r in enumerate(rows):
        try:
            msg = EmailMessage()
            msg["From"] = gmail
            msg["To"] = r["email"]
            msg["Subject"] = r["subject"]
            msg.set_content(r["body"])
            files = ([(cv_name or "CV.pdf", cv_bytes)] if cv_bytes else []) + list(extra_files)
            for fname, fbytes in files:
                msg.add_attachment(fbytes, maintype="application",
                                   subtype="octet-stream", filename=fname)
            imap.append(folder, "\\Draft", imaplib.Time2Internaldate(time.time()),
                        msg.as_bytes())
            saved += 1
            if on_saved:
                on_saved(r["email"])
        except Exception as e:
            errors.append(f"{r['email']}: {e}")
        if progress_cb:
            progress_cb((i + 1) / len(rows))
        time.sleep(0.3)
    imap.logout()
    return saved, errors


def make_docx(rows):
    doc = Document()
    doc.add_heading("Professor Emails for Review", 0)
    doc.add_paragraph(
        "Review each email, then copy it into your mail client and attach your CV.\n"
        "Group A = research matches yours.  Group B = other software / computer science / AI / "
        "data science professors whose research does not match closely."
    )
    groups = [("A", "Group A - Matching research"),
              ("B", "Group B - Software / CS / AI / Data Science (not matching)")]
    n = 0
    for g, title in groups:
        part = [r for r in rows if r.get("group") == g]
        if not part:
            continue
        doc.add_heading(title, 1)
        table = doc.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 1"
        for c, h in zip(table.rows[0].cells, ["#", "Professor", "University", "Email"]):
            c.text = h
        for i, r in enumerate(part, 1):
            cells = table.add_row().cells
            cells[0].text = str(i)
            cells[1].text = r["name"]
            cells[2].text = r["university"]
            cells[3].text = r["email"]
    for g, title in groups:
        part = [r for r in rows if r.get("group") == g]
        if not part:
            continue
        doc.add_page_break()
        doc.add_heading(title, 0)
        for i, r in enumerate(part, 1):
            doc.add_heading(f"{g}{i}. {r['name']} - {r['university']}", level=2)
            doc.add_paragraph(f"To: {r['email']}")
            if r.get("wechat"):
                doc.add_paragraph(f"WeChat (listed on the professor's page): {r['wechat']}")
            doc.add_paragraph(f"Profile: {r['url']}")
            doc.add_paragraph(f"Match score: {r['match_score']} - {r['match_reason']}")
            doc.add_paragraph(f"Subject: {r['subject']}")
            doc.add_paragraph(r["body"])
            doc.add_page_break()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="CSC Professor Outreach", layout="wide")
st.title("CSC Professor Outreach System")

ss = st.session_state
SETTING_KEYS = {"api_key": "", "model": "gemini-3.5-flash-lite", "gmail": "", "app_pw": "",
                "interval": 6.5}

if "init" not in ss:   # first run of this browser session: load everything saved on disk
    ss.init = True
    _s = load_json(SETTINGS_F, {})
    for _k, _d in SETTING_KEYS.items():
        ss[_k] = _s.get(_k, _d)
    ss.domcache = _s.get("domcache", {})
    ss.profile = load_json(PROFILE_F, None)
    ss.cv_name, ss.cv_bytes = load_file("cv")
    ss.emails = load_json(EMAILS_F, [])
    ss.drafted = set(load_json(DRAFTED_F, []))
    _u = []
    if os.path.exists(UNIS_F):
        for _l in open(UNIS_F, encoding="utf-8").read().splitlines():
            _a, _, _b = _l.partition("|")
            if _a.strip():
                _u.append((_a.strip(), _b.strip() or "auto"))
    ss.unis = _u or list(UNI_LIST)
    ss.profs = []

with st.sidebar:
    st.header("Settings")
    st.text_input("Free Gemini API key", type="password", key="api_key",
                  help="Get it free at https://aistudio.google.com/apikey (no card needed). "
                       "It is saved on your computer - you enter it only once.")
    st.text_input("Model", key="model",
                  help="If you get a 404 error, pick another free model name from AI Studio.")
    st.slider("Seconds between AI calls", 2.0, 20.0, key="interval",
              help="Free tier has strict limits. Raise this if you see rate-limit errors.")
    MIN_INTERVAL[0] = float(ss.interval)
    st.divider()
    st.subheader("Gmail drafts")
    st.text_input("Your Gmail address", key="gmail")
    st.text_input("Gmail App Password (16 letters)", type="password", key="app_pw",
                  help="NOT your normal password. Google Account > Security > "
                       "2-Step Verification > App passwords")
    st.caption("Emails are only saved as DRAFTS. Nothing is ever sent automatically.")
    st.divider()
    st.caption(f"Saved on this computer in: {DATA_DIR}  (settings, CV, professors, emails). "
               "The password is stored as plain text there - never share this folder.")
    def _forget_keys():
        for _k in ("api_key", "app_pw"):
            ss[_k] = ""

    def _forget_all():
        import shutil
        shutil.rmtree(DATA_DIR, ignore_errors=True)
        os.makedirs(FILES_DIR, exist_ok=True)
        for _k in list(ss.keys()):
            del ss[_k]

    with st.expander("Forget saved data"):
        st.button("Delete saved password & keys", on_click=_forget_keys)
        st.button("Delete ALL saved data (professors, emails, CV)", on_click=_forget_all)

# remember settings (so they are never asked again)
_cur = {k: ss[k] for k in SETTING_KEYS}
_cur["domcache"] = ss.domcache
if _cur != load_json(SETTINGS_F, None):
    save_json(SETTINGS_F, _cur)

client = (ss.api_key or "").strip() or None
model = ss.model.strip() or "gemini-3.5-flash-lite"
gmail, app_pw = ss.gmail, ss.app_pw


class _Up:   # lets read_cv work on a file saved on disk
    def __init__(self, name, data):
        self.name, self._d = name, data

    def getvalue(self):
        return self._d


t1, t2, t3, t4, t5 = st.tabs(
    ["1. CV", "2. Universities", "3. Find Professors", "4. Emails", "5. Export"]
)

# --- Tab 1 -----------------------------------------------------------------
with t1:
    if ss.cv_bytes:
        st.success(f"Saved CV: {ss.cv_name}")
    up = st.file_uploader("Upload CV / document (only needed once)", type=["pdf", "docx", "txt"])
    if up:
        ss.cv_bytes, ss.cv_name = up.getvalue(), up.name
        save_file("cv", up.name, up.getvalue())
    if st.button("Extract my details"):
        if not client:
            st.error("Enter your free Gemini API key in the sidebar.")
        elif not ss.cv_bytes:
            st.error("Upload your CV first.")
        else:
            with st.spinner("Reading CV..."):
                try:
                    ss.profile = extract_profile(client, model, read_cv(_Up(ss.cv_name, ss.cv_bytes)))
                    save_json(PROFILE_F, ss.profile)
                except Exception as ex:
                    st.error(f"Could not extract: {ex}")
    if ss.profile:
        st.subheader("Extracted profile (saved - edit if needed)")
        edited = st.text_area("JSON", json.dumps(ss.profile, indent=2, ensure_ascii=False),
                              height=400)
        if st.button("Save profile edits"):
            try:
                ss.profile = json.loads(edited)
                save_json(PROFILE_F, ss.profile)
                st.success("Saved")
            except Exception as e:
                st.error(f"Invalid JSON: {e}")

# --- Tab 2 -----------------------------------------------------------------
with t2:
    st.write(f"**{len(ss.unis)} universities loaded** (names only). Format per line: "
             "`Name | website domain` like `ustc.edu.cn`. Write just the university NAME and "
             "its website will be found automatically (and remembered).")
    txt = st.text_area("List", "\n".join(f"{n} | {u}" for n, u in ss.unis), height=400)
    if st.button("Update list"):
        items = []
        for line in txt.splitlines():
            line = line.strip()
            if not line:
                continue
            a, _, b = line.partition("|")
            a, b = a.strip(), b.strip()
            if a:
                items.append((a, b or "auto"))
        ss.unis = items
        open(UNIS_F, "w", encoding="utf-8").write("\n".join(f"{n} | {u}" for n, u in items))
        st.success(f"{len(items)} universities saved")
    names = [n for n, _ in ss.unis]
    chosen = st.multiselect("Search only these (empty = ALL)", names)


def build_profs(phash, min_score, include_b, recruit_mode, only):
    """Professors list = everything the AI has already judged (saved on disk), filtered."""
    seen, out = set(), []
    for p in db_ok_profs(phash):
        if only and p.get("university") not in only:
            continue
        rec = p.get("recruiting", "unknown")
        if recruit_mode.startswith("Only"):
            ok = rec == "international"
        elif recruit_mode.startswith("International OR"):
            ok = rec in ("international", "students")
        else:
            ok = rec != "no"
        em = (p.get("email") or "").lower()
        if not ok or not em or em in seen:
            continue
        if p.get("match_score", 0) >= min_score:
            p["group"] = "A"
        elif include_b and p.get("cs_related", True):
            p["group"] = "B"
        else:
            continue
        seen.add(em)
        out.append(p)
    out.sort(key=lambda x: (x["group"], -x.get("match_score", 0)))
    return out


# --- Tab 3 -----------------------------------------------------------------
with t3:
    st.info("Results are SAVED. Every run adds new professors to the saved list - it never "
            "starts from zero. If the AI quota ends, just press the button again later "
            "(or tomorrow); it continues where it stopped.")
    c1, c2, c3, c4 = st.columns(4)
    per_uni = c1.number_input("Profile pages per university", 10, 300, 80)
    min_score = c2.slider("Minimum match score for Group A", 0, 10, 6)
    max_new = c3.number_input("New pages per run (AI calls)", 10, 2000, 150,
                              help="Each page = 1 AI call. Press the button again to continue "
                                   "with the next pages.")
    threads = c4.number_input("Parallel downloads", 1, 32, 12)
    VERIFY_SSL[0] = not st.checkbox(
        "Skip SSL certificate check (use only if many universities fail to open; "
        "pages are public, nothing private is sent)", value=False)
    recruit_mode = st.radio(
        "Which professors to keep?",
        ["Only those whose page says they recruit INTERNATIONAL students",
         "International OR any student recruiting",
         "All software / CS / AI / data science professors (recommended)"],
        index=2,
        help="Most professor pages do not mention recruiting, so the strict options "
             "give few results.")
    discovery = st.radio(
        "How to find professor pages?",
        ["Direct crawl of university websites (no search engine - recommended)",
         "Web search (DuckDuckGo - often rate-limited)",
         "Both"], index=0)
    include_b = st.checkbox(
        "Group B: ALL other software / CS / AI / data science professors whose research does "
        "NOT match mine", value=True)
    RENDER_JS[0] = st.checkbox(
        "Use a real browser for JavaScript pages (slower; needs: pip install playwright, then "
        "playwright install chromium)", value=False)
    crawl_again = st.checkbox(
        "Look for new pages on the university websites first (untick = only continue "
        "analysing pages already saved)", value=True)

    phash = profile_hash(ss.profile) if ss.profile else ""
    only_set = set(chosen) if chosen else None

    if st.button("Find professors (continues where it stopped)"):
        if not (client and ss.profile):
            st.error("Need API key and extracted profile first (Tab 1).")
        else:
            targets = [(n, u) for n, u in ss.unis if not chosen or n in chosen]
            kws_en = ss.profile.get("search_keywords", []) or []
            kws_zh = ss.profile.get("search_keywords_zh", []) or []
            status = st.empty()
            bar = st.progress(0.0)
            if crawl_again:
                # Phase 1: open every faculty directory, collect profile links
                tasks, no_dir, cands, report = [], [], [], []
                with cf.ThreadPoolExecutor(int(threads)) as ex:
                    futs = {}
                    for n, u in targets:
                        if u.startswith("http"):
                            futs[ex.submit(crawl_directory, u, int(per_uni), 4, n)] = n
                        else:
                            no_dir.append((n, u))
                    for i, f in enumerate(cf.as_completed(futs)):
                        n = futs[f]
                        status.info(f"Reading faculty directories {i+1}/{len(futs)}: {n}")
                        try:
                            links, blocks = f.result()
                            tasks += [(l, n) for l in links]
                            cands += blocks
                        except Exception:
                            pass
                        bar.progress((i + 1) / max(len(futs), 1))

                # Phase 1b: universities given by domain/name
                doms, skipped = [], []
                for n, dom in no_dir:
                    if dom == "auto":
                        dom = ss.domcache.get(n)
                        if not dom and SEARCH_STATE["fails"] < 8:
                            status.info(f"Finding website of {n} ...")
                            dom = resolve_domain(n)
                            if dom:
                                ss.domcache[n] = dom
                        if not dom:
                            skipped.append(n)
                            continue
                    doms.append((n, dom))

                if not discovery.startswith("Web search"):
                    bar.progress(0.0)
                    with cf.ThreadPoolExecutor(max(2, int(threads) // 3)) as ex:
                        futs = {ex.submit(direct_crawl, n, d, int(per_uni)): n for n, d in doms}
                        for i, f in enumerate(cf.as_completed(futs)):
                            n = futs[f]
                            try:
                                links, blocks, nd = f.result()
                                report.append({"university": n, "faculty pages": nd,
                                               "profile links": len(links), "list rows": len(blocks)})
                                tasks += [(l, n) for l in links]
                                cands += blocks
                                status.info(f"Direct crawl {i+1}/{len(futs)}: {n} -> {nd} faculty "
                                            f"pages, {len(links)} profile links, {len(blocks)} list rows")
                            except Exception:
                                pass
                            bar.progress((i + 1) / max(len(futs), 1))

                if not discovery.startswith("Direct"):
                    bar.progress(0.0)
                    pauses = 0
                    for i, (n, dom) in enumerate(doms):
                        found_l = search_fallback(dom, kws_en, kws_zh, 8)
                        tasks += [(l, n) for l in found_l]
                        status.info(f"Searching web {i+1}/{len(doms)}: {n} -> {len(found_l)} links "
                                    f"(total {len(tasks)})")
                        bar.progress((i + 1) / len(doms))
                        if SEARCH_STATE["fails"] >= 8:
                            pauses += 1
                            if pauses > 3:
                                st.error("Web search keeps failing (blocked or rate-limited): "
                                         f"{SEARCH_STATE['last_error']}. Stopped searching.")
                                break
                            for sec in range(90, 0, -5):
                                status.warning(f"Search is rate-limited. Pausing {sec}s "
                                               f"(pause {pauses}/3) ...")
                                time.sleep(5)
                            SEARCH_STATE["fails"] = 0
                if skipped:
                    st.warning(f"Website not found for {len(skipped)} universities (skipped). "
                               "Add their domain in Tab 2 (Name | domain): " + "; ".join(skipped[:15])
                               + (" ..." if len(skipped) > 15 else ""))

                # Phase 2: download profile pages (skip pages already saved)
                with _db_lock:
                    _con = db()
                    known_urls = {r[0] for r in _con.execute("SELECT url FROM cands")}
                    _con.close()
                tasks = [(l, n) for l, n in dict.fromkeys(tasks) if l not in known_urls]
                with cf.ThreadPoolExecutor(int(threads)) as ex:
                    futs = [ex.submit(fetch_candidate, l, n) for l, n in tasks]
                    for i, f in enumerate(cf.as_completed(futs)):
                        try:
                            r = f.result()
                            if r:
                                cands.append(r)
                        except Exception:
                            pass
                        if i % 10 == 0:
                            status.info(f"Downloading profile pages {i+1}/{len(futs)} "
                                        f"- {len(cands)} pages with emails so far")
                        bar.progress((i + 1) / max(len(futs), 1))

                got = {c["uni"] for c in cands}
                miss = [(n, u) for n, u in targets if n not in got and u.startswith("http")]
                for n, u in miss:
                    status.info(f"No profiles from directory, trying web search: {n}")
                    for l in search_fallback(root_domain(u), kws_en, kws_zh, 6):
                        r = fetch_candidate(l, n)
                        if r:
                            cands.append(r)
                if report:
                    st.write("What was found on each university website (0 = problem, see tips below):")
                    st.dataframe(report, use_container_width=True)
                added = db_add_cands(cands)
                st.write(f"{len(cands)} pages with a professor email found, **{added} NEW** "
                         "(already-saved pages are not repeated).")

            # Phase 3: AI judges only pages that are new or failed before (free pre-filter first)
            pend = db_pending(phash, {n for n, _ in targets} if chosen else None)
            skipped_nc = 0
            for c in list(pend):
                if not is_cs_page(c):
                    db_set_result(c["key"], phash, "notcs")
                    pend.remove(c)
                    skipped_nc += 1
            for c in pend:
                c["lscore"] = local_score(c["text"], kws_en + kws_zh)
            pend.sort(key=lambda c: c["lscore"], reverse=True)
            todo = pend[: int(max_new)]
            st.write(f"Free filter removed {skipped_nc} non-computing pages. "
                     f"{len(pend)} pages wait for AI; analysing {len(todo)} now.")
            bar.progress(0.0)
            n_ok = n_notprof = n_fail = consec = 0
            last_err = ""
            for i, c in enumerate(todo):
                status.info(f"AI analyzing {i+1}/{len(todo)}: {c['uni']}  "
                            f"(professors found this run: {n_ok})")
                try:
                    p = analyze_page(client, model, ss.profile, c["uni"], c["url"], c["text"])
                    consec = 0
                except FatalAPIError as e:
                    st.error(f"{e}  -  progress is saved; press the button again later.")
                    break
                except (TransientAIError, requests.RequestException) as e:
                    n_fail += 1
                    consec += 1
                    last_err = str(e)[:160]
                    db_set_result(c["key"], phash, "fail")
                    if consec >= 4:
                        st.error("The AI keeps failing, so I stopped (nothing is lost): "
                                 f"{last_err}. Wait a few minutes (or raise 'Seconds between AI "
                                 "calls' in the sidebar) and press the button again.")
                        break
                    continue
                if p == "notprof":
                    n_notprof += 1
                    db_set_result(c["key"], phash, "notprof")
                else:
                    n_ok += 1
                    db_set_result(c["key"], phash, "ok", p)
                bar.progress((i + 1) / max(len(todo), 1))
            status.success(f"Run finished: {n_ok} professors found, {n_notprof} pages were not "
                           f"professor pages, {n_fail} failed (will be retried next time).")

    if ss.profile:
        ss.profs = build_profs(phash, min_score, include_b, recruit_mode, only_set)
        s_ = db_stats(phash)
        na_ = sum(p["group"] == "A" for p in ss.profs)
        st.success(f"**Group A (matching): {na_}  |  Group B (other CS/AI/data science): "
                   f"{len(ss.profs) - na_}**  -  total saved professors: {len(ss.profs)}")
        st.caption(f"Saved pages: {s_['total']} | analysed by AI: {s_['ok'] + s_['notprof']} | "
                   f"skipped as non-computing (free): {s_['notcs']} | still waiting: "
                   f"{s_['pending']} | failed (retry): {s_['fail']}")
    with st.expander("Diagnose one university (use if a university gives 0 professors)"):
        if st.button("Test web search (does it work on my internet?)"):
            SEARCH_STATE["fails"] = 0
            t0 = time.time()
            res = ddg_text("site:pku.edu.cn professor computer science", 5)
            if res:
                st.success(f"Web search works: {len(res)} results in {time.time()-t0:.1f}s")
                st.write([r.get("href") for r in res])
            else:
                st.error(f"Web search FAILED: {SEARCH_STATE['last_error']}")
        dn = st.selectbox("University", names, key="dn")
        if st.button("Run diagnosis"):
            durl = dict(ss.unis)[dn]
            if not durl.startswith("http"):
                dom_ = ss.domcache.get(dn) or (durl if durl != "auto" else "")
                if not dom_:
                    st.info("Website not known yet - run the search once or write `Name | domain`.")
                else:
                    dd = discover_directories(dom_, 6)
                    st.write(f"Faculty pages found automatically: **{len(dd)}**")
                    rows_ = []
                    for d_ in dd:
                        l_, b_ = crawl_directory(d_, 60, 8, dn)
                        rows_.append({"faculty page": d_, "profile links": len(l_),
                                      "list rows": len(b_), "error": ERRS.get(d_, "")})
                    st.dataframe(rows_)
                    st.caption("If 0 pages: copy the real faculty-list URL from your browser and "
                               "write it in Tab 2 as `Name | https://...`")
            else:
                dh = get_html(durl)
                if not dh:
                    st.error(f"Could not open the directory page.\n\n{ERRS.get(durl)}")
                else:
                    dl, db_ = crawl_directory(durl, 60, 4, dn)
                    st.write(f"Directory opened ({len(dh)} characters). "
                             f"Profile links found: **{len(dl)}**. "
                             f"List rows with exactly one email: **{len(db_)}**.")
    if ss.profs:
        st.dataframe(
            [{"group": p["group"], "name": p["name"], "university": p["university"],
              "email": p["email"], "score": p.get("match_score", 0),
              "area": p.get("cs_field") or p.get("se_field", ""),
              "recruiting": p.get("recruiting", ""),
              "wechat (public on page)": p.get("wechat", ""),
              "reason": p.get("match_reason", "")} for p in ss.profs],
            use_container_width=True,
        )

# --- Tab 4 -----------------------------------------------------------------
with t4:
    fmt_label = st.radio(
        "Email format",
        ["Format 2 - Bachelor thesis format (Dear Professor ... Regards)",
         "Format 1 - Projects + CSC enquiry format (Best regards)"], index=0)
    fmt = "thesis" if fmt_label.startswith("Format 2") else "projects"
    mode = st.radio(
        "How to write the emails?",
        ["Hybrid (recommended): AI for the first N emails, free template for the rest",
         "AI for all (best quality, needs a lot of free quota)",
         "Free template for all (no AI at all, unlimited)"], index=0,
        help="Free Gemini allows only a limited number of AI calls per day. The free "
             "template uses your same format and each professor's real research topics.")
    ai_n = st.number_input("N = emails written by AI in Hybrid mode", 0, 500, 40)
    csc_year = st.text_input("CSC intake year used in emails", value="2027")
    extra = st.text_input("Extra instruction for all emails (optional, AI only)",
                          placeholder="e.g. mention I can start September 2027")
    skip_done = st.checkbox("Skip professors who already have an email (add only new ones)",
                            value=True)
    n_a = sum(p.get("group") == "A" for p in ss.profs)
    n_b = sum(p.get("group") == "B" for p in ss.profs)
    ca, cb = st.columns(2)
    limit_a = ca.number_input(f"Group A emails (matching) - {n_a} available", 0, 5000, n_a)
    limit_b = cb.number_input(f"Group B emails (other CS/AI/DS) - {n_b} available", 0, 5000, n_b)
    if st.button("Generate personalized emails"):
        if not ss.profs or not ss.profile:
            st.error("Need professors first (Tab 3).")
        elif not client and not mode.startswith("Free"):
            st.error("Enter the Gemini key, or choose 'Free template for all'.")
        else:
            have = {e["email"].lower() for e in ss.emails} if skip_done else set()
            if not skip_done:
                ss.emails = []
            out = ss.emails
            openings, ai_used, ai_dead = [], 0, False
            attach = "CV"   # only the CV is attached / mentioned in the email
            targets = ([p for p in ss.profs if p["group"] == "A"][: int(limit_a)]
                       + [p for p in ss.profs if p["group"] == "B"][: int(limit_b)])
            targets = [p for p in targets if p["email"].lower() not in have]
            bar = st.progress(0.0)
            msg = st.empty()
            for i, p in enumerate(targets):
                use_ai = (not ai_dead and client and
                          (mode.startswith("AI") or (mode.startswith("Hybrid") and ai_used < int(ai_n))))
                try:
                    if use_ai:
                        try:
                            e = write_email(client, model, ss.profile, p, openings, extra,
                                            csc_year, attach, fmt)
                            ai_used += 1
                            openings.append(e["body"][:160].split("\n\n")[-1][:90])
                        except FatalAPIError as ex:
                            ai_dead = True
                            msg.warning(f"AI stopped ({ex}). Remaining emails use the free template.")
                            e = write_email_template(ss.profile, p, csc_year, attach, fmt)
                        except Exception:
                            e = write_email_template(ss.profile, p, csc_year, attach, fmt)
                    else:
                        e = write_email_template(ss.profile, p, csc_year, attach, fmt)
                    out.append({**p, **e})
                    save_json(EMAILS_F, out)
                except Exception as ex:
                    st.warning(f"Skipped {p.get('name')}: {ex}")
                bar.progress((i + 1) / max(len(targets), 1))
            st.success(f"{len(targets)} new emails written ({ai_used} by AI, the rest by free "
                       f"template). Total saved: {len(out)}")
    show_all = st.checkbox("Show all emails for editing (slow if there are hundreds)", value=False)
    for i, e in enumerate(ss.emails if show_all else ss.emails[:25]):
        with st.expander(f"[{e['group']}] {e['name']} ({e['university']}) - score {e.get('match_score', 0)}"):
            e["subject"] = st.text_input("Subject", e["subject"], key=f"s{i}")
            e["body"] = st.text_area("Body", e["body"], height=260, key=f"b{i}")
            st.caption(f"To: {e['email']} | {e['url']}")
    if ss.emails:
        save_json(EMAILS_F, ss.emails)

# --- Tab 5 -----------------------------------------------------------------
with t5:
    if not ss.emails:
        st.info("Generate emails first.")
    else:
        todo_e = [e for e in ss.emails if e["email"].lower() not in ss.drafted]
        st.write(f"Total emails ready: **{len(ss.emails)}**  |  already in Gmail drafts: "
                 f"**{len(ss.emails) - len(todo_e)}**  |  not drafted yet: **{len(todo_e)}**")
        st.download_button("Download ALL emails (.docx)",
                           make_docx(ss.emails), "all_professor_emails.docx")
        st.divider()
        st.subheader("Save as Gmail drafts")
        only_new = st.checkbox("Skip emails already saved as drafts", value=True)
        pool = todo_e if only_new else ss.emails
        ea = [e for e in pool if e.get("group") == "A"]
        eb = [e for e in pool if e.get("group") == "B"]
        d1, d2 = st.columns(2)
        na = d1.number_input(f"Group A drafts (of {len(ea)})", 0, 5000, len(ea))
        nb = d2.number_input(f"Group B drafts (of {len(eb)})", 0, 5000, len(eb))
        st.caption("Your CV is attached to every draft. Tip: try 2 drafts first.")
        if st.button("Test Gmail login"):
            try:
                _im = imaplib.IMAP4_SSL("imap.gmail.com")
                _im.login(gmail.strip(), app_pw.replace(" ", ""))
                _im.logout()
                st.success("Gmail login works.")
            except Exception as ex:
                st.error(f"Gmail login failed: {ex}. Check the App Password and that "
                         "2-Step Verification is on.")
        if st.button("Save as Gmail drafts"):
            if not (gmail and app_pw):
                st.error("Enter Gmail address and App Password in the sidebar (once).")
            elif not ss.cv_bytes:
                st.error("CV file not found - upload it again in Tab 1.")
            else:
                rows = ea[: int(na)] + eb[: int(nb)]
                bar = st.progress(0.0)

                def _saved(addr):
                    ss.drafted.add(addr.lower())
                    save_json(DRAFTED_F, sorted(ss.drafted))
                try:
                    # Only the CV is attached (no transcript)
                    ok, errs = save_drafts(gmail.strip(), app_pw.replace(" ", ""), rows,
                                           ss.cv_name, ss.cv_bytes, bar.progress,
                                           [],
                                           on_saved=_saved)
                    st.success(f"{ok} drafts saved. Open Gmail > Drafts, review, then send.")
                    for er in errs:
                        st.error(er)
                except imaplib.IMAP4.error as ex:
                    st.error(f"Gmail login failed: {ex}. Check the App Password and "
                             "that 2-Step Verification is on.")
                except Exception as ex:
                    st.error(f"Could not save drafts: {ex}")
        if st.button("Forget which emails were drafted (to draft them again)"):
            ss.drafted = set()
            save_json(DRAFTED_F, [])
            st.rerun()
