"""link_checker.py

Checks that each reference link works and leads to the paper that the
reference names. Deterministic, standard library only. Used by the AI agent's
``check_links`` tool, and usable on its own::

    python link_checker.py "input/My Essay.docx"

How a link is checked (most reliable source first)
--------------------------------------------------
1. DOI links - ``https://doi.org/10.xxxx/...`` or a DOI inside the URL (for
   example ``link.springer.com/article/10.1007/...``). The title registered
   for the DOI at Crossref, the official registry for journal DOIs, is
   compared with the reference text. Publisher websites often block automated
   checks; Crossref does not. A doi.org link whose DOI no registry knows is a
   PROBLEM (the link leads nowhere).
2. ScienceDirect-style links (``.../pii/S...``) are looked up at Crossref by
   their PII (Elsevier's article id).
3. PubMed Central links (``PMC1234567``) are looked up at Europe PMC.
4. Anything else: the page (or PDF) is downloaded and its title compared.

Verdicts
--------
OK          the link works and its title matches the reference
CHECK       the link works, but the title only partly matches or cannot be
            read (e.g. many PDFs) - compare it yourself
PROBLEM     broken (page not found, DOI does not exist, website gone) or it
            leads to a different paper
UNVERIFIED  could not be checked automatically (the site blocks automated
            requests, timed out, or there is no internet) - open it yourself
NO_LINK     the reference has no URL

Only registry metadata (Crossref / Europe PMC) can produce a "different
paper" PROBLEM: website and PDF titles are less reliable (site names, file
names), so a mismatch there is only CHECK. Nothing here ever changes a
document; the checker only reads.
"""
import html
import http.client
import json
import re
import socket
import ssl
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from html.parser import HTMLParser

# Endpoints are module attributes so tests can point them at a local server.
CROSSREF_API = "https://api.crossref.org"
EUROPEPMC_API = "https://www.ebi.ac.uk/europepmc/webservices/rest"
DOI_HANDLE_API = "https://doi.org/api/handles"

USER_AGENT = ("Mozilla/5.0 (compatible; citation-agent-linkcheck/1.0; "
              "+https://github.com/nerduser13-ux/citation_agent)")
TIMEOUT = 15             # seconds per request
MAX_BYTES = 1_000_000    # read at most this much of a page / PDF
MAX_WORKERS = 6          # links checked in parallel

OK = "OK"
CHECK = "CHECK"
PROBLEM = "PROBLEM"
UNVERIFIED = "UNVERIFIED"
NO_LINK = "NO_LINK"
VERDICTS = (OK, CHECK, PROBLEM, UNVERIFIED, NO_LINK)

MATCH_OK = 0.75        # share of the title's words found in the reference
MATCH_PARTIAL = 0.40   # below this, registry metadata = a different paper

# Statuses that usually mean "no robots", not "broken".
_BLOCKED_STATUSES = {401, 403, 405, 406, 418, 429, 451, 503, 999}
_BOT_CHECK_TITLE = re.compile(
    r"just a moment|attention required|access denied|are you a robot|captcha|"
    r"security check|checking your browser|verify you are human|"
    r"bot protection|request rejected|pardon our interruption", re.I)
_NOT_FOUND_TITLE = re.compile(
    r"\b(404|page not found|not found|no longer available|does not exist)\b",
    re.I)


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
class Response:
    """Result of one GET. ``error`` is set (and ``status`` None) when the
    server could not be reached at all."""

    def __init__(self, status=None, url="", content_type="", body=b"",
                 error=None, dns_failure=False):
        self.status = status
        self.url = url
        self.content_type = content_type or ""
        self.body = body or b""
        self.error = error
        self.dns_failure = dns_failure


def _network_error(exc):
    reason = getattr(exc, "reason", exc)
    if isinstance(reason, socket.gaierror):
        return "the website address could not be found", True
    if isinstance(reason, (socket.timeout, TimeoutError)):
        return "the website took too long to respond", False
    if isinstance(reason, ssl.SSLError):
        return "a secure (https) connection could not be made", False
    if isinstance(reason, ConnectionRefusedError):
        return "the website refused the connection", False
    return f"the website could not be reached ({reason})", False


def http_get(url, accept="*/*", max_bytes=MAX_BYTES, timeout=TIMEOUT):
    """GET ``url`` (redirects are followed). Never raises for network or
    HTTP errors - they are returned in the Response."""
    req = urllib.request.Request(url, headers={
        "User-Agent": USER_AGENT, "Accept": accept, "Accept-Language": "en"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return Response(resp.status, resp.geturl(),
                            resp.headers.get("Content-Type", ""),
                            resp.read(max_bytes))
    except urllib.error.HTTPError as e:
        try:
            body = e.read(max_bytes)
        except Exception:
            body = b""
        ctype = e.headers.get("Content-Type", "") if e.headers else ""
        return Response(e.code, e.geturl() or url, ctype, body)
    except (urllib.error.URLError, OSError, http.client.HTTPException,
            ValueError) as e:
        msg, dns = _network_error(e)
        return Response(url=url, error=msg, dns_failure=dns)


def _get_json(url):
    """(parsed JSON or None, Response)."""
    r = http_get(url, accept="application/json")
    if r.error or r.status != 200:
        return None, r
    try:
        return json.loads(r.body.decode("utf-8", "replace")), r
    except ValueError:
        return None, r


# --------------------------------------------------------------------------
# Identifiers inside URLs
# --------------------------------------------------------------------------
_DOI_RE = re.compile(r"10\.\d{4,9}/\S+", re.I)
# Path pieces publishers append after a DOI (…/10.1108/X-1/full/html).
_DOI_SUFFIXES = ("/full/html", "/full", "/abstract", "/abs", "/pdf", "/epdf",
                 "/html", "/fulltext", "/references", "/meta", "/summary",
                 "/figures", "/supplemental")
_PII_RE = re.compile(r"/pii/([A-Za-z0-9()\-]+)", re.I)
_PMC_RE = re.compile(r"\b(PMC\d{4,})\b", re.I)


def _is_doi_host(host):
    return host == "doi.org" or host.endswith(".doi.org")


def doi_candidates(url):
    """DOIs that ``url`` may contain, most specific first. For publisher URLs
    trailing path pieces (``/full/html``) are trimmed off as fallbacks."""
    parsed = urllib.parse.urlparse((url or "").strip())
    host = parsed.netloc.lower()
    path = urllib.parse.unquote(parsed.path)
    texts = [path.lstrip("/")] if _is_doi_host(host) else (
        [path] + [v for vs in urllib.parse.parse_qs(parsed.query).values()
                  for v in vs])
    out = []
    for text in texts:
        m = _DOI_RE.search(text)
        if not m:
            continue
        doi = m.group(0).rstrip(".,;:)]}>\"'")
        if doi.lower().endswith(".pdf"):
            doi = doi[:-4]
        cands = [doi]
        if not _is_doi_host(host):
            trimmed = doi
            for _ in range(3):
                low = trimmed.lower()
                suffix = next((s for s in _DOI_SUFFIXES if low.endswith(s)), None)
                if suffix:
                    trimmed = trimmed[:-len(suffix)]
                elif trimmed.count("/") > 1:
                    trimmed = trimmed.rsplit("/", 1)[0]
                else:
                    break
                cands.append(trimmed)
        for c in cands:
            if c and c not in out:
                out.append(c)
    return out


def extract_pii(url):
    m = _PII_RE.search(urllib.parse.urlparse(url or "").path)
    return re.sub(r"[^A-Za-z0-9]", "", m.group(1)).upper() if m else None


def extract_pmcid(url):
    m = _PMC_RE.search(url or "")
    return m.group(1).upper() if m else None


# --------------------------------------------------------------------------
# Registries
# --------------------------------------------------------------------------
def lookup_doi(doi):
    """("found", crossref_record) | ("missing", None) | ("error", message)"""
    data, r = _get_json(f"{CROSSREF_API}/works/{urllib.parse.quote(doi, safe='/')}")
    if data is not None and isinstance(data.get("message"), dict):
        return "found", data["message"]
    if r.error:
        return "error", r.error
    if r.status == 404:
        return "missing", None
    return "error", f"Crossref answered HTTP {r.status}"


def doi_exists(doi):
    """Ask doi.org whether the DOI is registered anywhere (Crossref, DataCite,
    mEDRA, ...). True / False / None (could not tell)."""
    data, r = _get_json(f"{DOI_HANDLE_API}/{urllib.parse.quote(doi, safe='/')}")
    if data is not None:
        return data.get("responseCode") == 1
    if not r.error and r.status == 404:
        return False
    return None


def lookup_pii(pii):
    query = urllib.parse.urlencode({"filter": f"alternative-id:{pii}", "rows": 2})
    data, r = _get_json(f"{CROSSREF_API}/works?{query}")
    if data is not None:
        items = (data.get("message") or {}).get("items") or []
        return ("found", items[0]) if items else ("missing", None)
    return "error", r.error or f"Crossref answered HTTP {r.status}"


def lookup_pmcid(pmcid):
    query = urllib.parse.urlencode(
        {"query": f"PMCID:{pmcid}", "format": "json", "resultType": "lite"})
    data, r = _get_json(f"{EUROPEPMC_API}/search?{query}")
    if data is not None:
        hits = (data.get("resultList") or {}).get("result") or []
        return ("found", hits[0]) if hits else ("missing", None)
    return "error", r.error or f"Europe PMC answered HTTP {r.status}"


def _crossref_year(rec):
    for key in ("published-print", "published-online", "issued", "published"):
        parts = (rec.get(key) or {}).get("date-parts") or []
        if parts and parts[0] and parts[0][0]:
            return parts[0][0]
    return None


# --------------------------------------------------------------------------
# Titles
# --------------------------------------------------------------------------
def clean_text(s):
    s = html.unescape(s or "")
    s = re.sub(r"<[^>]+>", " ", s)          # Crossref titles contain <i>, <sub>
    return re.sub(r"\s+", " ", s).strip()


class _HeadParser(HTMLParser):
    KEYS = ("citation_title", "dc.title", "og:title", "twitter:title")

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.meta = {}
        self.title = None
        self._parts = None
        self._in_body = False

    def handle_starttag(self, tag, attrs):
        if self._in_body:
            return
        a = {k.lower(): (v or "") for k, v in attrs}
        if tag == "meta":
            key = (a.get("name") or a.get("property") or "").lower()
            if key in self.KEYS and key not in self.meta and a.get("content"):
                self.meta[key] = a["content"]
        elif tag == "title" and self.title is None:
            self._parts = []
        elif tag == "body":
            self._in_body = True

    def handle_endtag(self, tag):
        if tag == "title" and self._parts is not None:
            self.title = "".join(self._parts)
            self._parts = None

    def handle_data(self, data):
        if self._parts is not None:
            self._parts.append(data)


def _decode(body, content_type):
    m = re.search(r"charset=([\w\-]+)", content_type or "", re.I)
    enc = m.group(1) if m else None
    if not enc:
        m = re.search(rb"<meta[^>]+charset=[\"']?([\w\-]+)", body[:4096], re.I)
        enc = m.group(1).decode("ascii", "ignore") if m else "utf-8"
    try:
        return body.decode(enc, "replace")
    except LookupError:
        return body.decode("utf-8", "replace")


def html_title(body, content_type=""):
    """Best title of an HTML page: citation_title (what journals publish for
    Google Scholar) > dc.title > og:title > twitter:title > <title>."""
    parser = _HeadParser()
    try:
        parser.feed(_decode(body, content_type))
    except Exception:
        pass
    for key in _HeadParser.KEYS:
        if parser.meta.get(key):
            return clean_text(parser.meta[key])
    return clean_text(parser.title or "")


def pdf_title(body):
    """Best-effort title from PDF metadata (XMP dc:title, else /Title)."""
    m = re.search(rb"<dc:title>\s*<rdf:Alt>\s*<rdf:li[^>]*>(.*?)</rdf:li>",
                  body, re.S)
    if m:
        return clean_text(m.group(1).decode("utf-8", "replace"))
    m = re.search(rb"/Title\s*\((.*?)(?<!\\)\)", body, re.S)
    if m:
        raw = m.group(1)
        if raw.startswith(b"\xfe\xff"):
            text = raw[2:].decode("utf-16-be", "replace")
        else:
            text = raw.decode("latin-1", "replace")
        return clean_text(re.sub(r"\\([()\\])", r"\1", text))
    return ""


_STOPWORDS = frozenset(
    "a an and are as at be by for from has in into is it its of on or over "
    "than that the their this to under using via vs with within without".split())


def _words(text):
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).lower()
    return [w for w in re.findall(r"[a-z0-9]+", text) if w not in _STOPWORDS]


def title_match(title, reference_text):
    """Share (0..1) of the title's words that also appear in the reference.
    Containment, not similarity: the reference also holds authors, journal,
    year and URL, which must not count against the title."""
    t = set(_words(title))
    if not t:
        return 0.0
    return round(len(t & set(_words(reference_text))) / len(t), 2)


def _page_title_match(title, reference_text):
    """Like title_match, but page titles often carry a site name
    ("Paper title | Journal"), so the best segment counts too."""
    best = title_match(title, reference_text)
    for seg in re.split(r"\s+[|\-\u2013\u2014:]{1,2}\s+", title):
        if len(_words(seg)) >= 3:
            best = max(best, title_match(seg, reference_text))
    return best


# --------------------------------------------------------------------------
# Checking
# --------------------------------------------------------------------------
def _judge_registry(res, title, reference_text, source, subtitle="",
                    year=None, doi=None):
    title, subtitle = clean_text(title), clean_text(subtitle)
    score = title_match(title, reference_text)
    if subtitle:
        score = max(score, title_match(f"{title} {subtitle}", reference_text))
    res.update(found_title=title + (f": {subtitle}" if subtitle else ""),
               match=score, source=source, reached_internet=True)
    if year:
        res["found_year"] = year
    if doi:
        res["doi"] = doi
    if not title:
        res.update(verdict=CHECK, detail=(
            f"The link is registered at {source}, but no title is recorded - "
            "compare it yourself."))
    elif score >= MATCH_OK:
        res.update(verdict=OK, detail=f"Registered at {source}; the title matches.")
    elif score >= MATCH_PARTIAL:
        res.update(verdict=CHECK, detail=(
            f"Registered at {source}, but the title only partly matches the "
            "reference - compare them."))
    else:
        res.update(verdict=PROBLEM, detail=(
            f"This link leads to a different paper: \"{title}\"."))
    return res


def _check_page(res, url, reference_text):
    r = http_get(url)
    if r.error:
        res.update(verdict=UNVERIFIED, detail=f"Could not open the link: {r.error}.",
                   dns_failure=r.dns_failure, network_error=True)
        return res
    res["reached_internet"] = True
    res["http_status"] = r.status
    if r.status in (404, 410):
        res.update(verdict=PROBLEM, detail=f"The page was not found (HTTP {r.status}).")
        return res
    if r.status in _BLOCKED_STATUSES:
        res.update(verdict=UNVERIFIED, detail=(
            f"The website blocks automated checks (HTTP {r.status}) - open the "
            "link in your browser."))
        return res
    if r.status >= 400:
        res.update(verdict=UNVERIFIED, detail=(
            f"The website answered with an error (HTTP {r.status}) - try the "
            "link in your browser."))
        return res
    is_pdf = "pdf" in r.content_type.lower() or r.body[:5] == b"%PDF-"
    title = pdf_title(r.body) if is_pdf else html_title(r.body, r.content_type)
    res["source"] = "PDF" if is_pdf else "web page"
    if title and _BOT_CHECK_TITLE.search(title):
        res.update(verdict=UNVERIFIED, detail=(
            "The website showed a robot check instead of the page - open the "
            "link in your browser."))
        return res
    res["found_title"] = title
    if not title:
        res.update(verdict=CHECK, detail=(
            "The PDF opens, but its title can't be read automatically - take a "
            "quick look." if is_pdf else
            "The page opens, but it has no readable title - take a quick look."))
        return res
    score = _page_title_match(title, reference_text)
    res["match"] = score
    if score >= MATCH_OK:
        res.update(verdict=OK, detail="The link works and the title matches.")
    elif score < MATCH_PARTIAL and _NOT_FOUND_TITLE.search(title):
        res.update(verdict=PROBLEM, detail=f"The page says: \"{title}\".")
    else:
        res.update(verdict=CHECK, detail=(
            f"The link works, but the title (\"{title}\") doesn't clearly match "
            "the reference - compare it yourself."))
    return res


def check_link(url, reference_text=""):
    """Check one link against its reference text. Returns a plain dict with
    ``verdict`` (see module docstring), ``detail`` (plain-English reason),
    ``found_title``, ``match`` (0..1 or None) and ``source``."""
    url = (url or "").strip()
    res = {"url": url, "verdict": UNVERIFIED, "detail": "", "found_title": "",
           "match": None, "source": ""}
    if not url:
        res.update(verdict=NO_LINK, detail="This reference has no link.")
        return res
    host = urllib.parse.urlparse(url).netloc.lower()

    # 1. DOI -> Crossref
    registry_down = False
    candidates = doi_candidates(url)
    for doi in candidates:
        state, rec = lookup_doi(doi)
        if state == "found":
            return _judge_registry(res, " ".join(rec.get("title") or []),
                                   reference_text, "Crossref",
                                   subtitle=" ".join(rec.get("subtitle") or []),
                                   year=_crossref_year(rec), doi=doi)
        if state == "error":
            registry_down = True
            break
        res["reached_internet"] = True
    if candidates and _is_doi_host(host) and not registry_down:
        # Not a Crossref DOI. Does it exist at all (DataCite, mEDRA, ...)?
        if doi_exists(candidates[0]) is False:
            res.update(verdict=PROBLEM, detail=(
                f"The DOI {candidates[0]} does not exist, so this link leads "
                "nowhere - check the reference."))
            return res
        # Registered elsewhere, or unknown: fall through to the page itself.

    # 2. ScienceDirect PII -> Crossref
    pii = extract_pii(url)
    if pii and not registry_down:
        state, rec = lookup_pii(pii)
        if state == "found":
            return _judge_registry(res, " ".join(rec.get("title") or []),
                                   reference_text, "Crossref",
                                   subtitle=" ".join(rec.get("subtitle") or []),
                                   year=_crossref_year(rec), doi=rec.get("DOI"))

    # 3. PubMed Central -> Europe PMC
    pmcid = extract_pmcid(url)
    if pmcid:
        state, rec = lookup_pmcid(pmcid)
        if state == "found":
            year = rec.get("pubYear")
            return _judge_registry(res, rec.get("title") or "", reference_text,
                                   "Europe PMC", year=int(year) if str(year).isdigit() else None,
                                   doi=rec.get("doi"))

    # 4. The page itself
    return _check_page(res, url, reference_text)


def check_references(refs, numbers=None, max_workers=MAX_WORKERS):
    """Check the links of a reference map ``{number: {"text", "url"}}``.

    Returns ``{"checked", "summary": {verdict: count}, "note", "results"}``
    with one result per reference (sorted by number)."""
    items = sorted(refs.items())
    if numbers:
        wanted = {int(n) for n in numbers}
        items = [(n, r) for n, r in items if n in wanted]

    def one(item):
        num, ref = item
        text = ref.get("text") or ""
        out = check_link(ref.get("url") or "", text)
        return {"reference": num, "reference_text": text[:160], **out}

    if items:
        with ThreadPoolExecutor(max_workers=max(1, min(max_workers, len(items)))) as ex:
            results = list(ex.map(one, items))
    else:
        results = []

    online = any(r.get("reached_internet") for r in results)
    offline = not online and any(r.get("network_error") for r in results)
    for r in results:
        # "Address not found" only means "website gone" when other links
        # worked - otherwise this computer is probably offline.
        if r.pop("dns_failure", False) and online and r["verdict"] == UNVERIFIED:
            r.update(verdict=PROBLEM, detail=(
                "The website address doesn't exist (any more) - the link is broken."))
        r.pop("reached_internet", None)
        r.pop("network_error", None)
    summary = {v: sum(1 for r in results if r["verdict"] == v) for v in VERDICTS}
    note = ("No link could be reached - is this computer connected to the "
            "internet?") if offline else ""
    return {"checked": len(results), "summary": summary, "note": note,
            "results": results}


def main(argv=None):
    """``python link_checker.py <document.docx>`` - print a link report."""
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 1:
        print('Usage: python link_checker.py "input/My Essay.docx"')
        return 2
    from config import Config
    from document_reader import DocumentReader
    from reference_parser import ReferenceParser

    cfg = Config()
    paras = DocumentReader(argv[0]).all_paragraphs()
    found = ReferenceParser(cfg).find_reference_section(paras, cfg)
    if found is None:
        print("Could not find the reference list in this document.")
        return 1
    refs, _ = ReferenceParser().parse_references(paras, found[0])
    report = check_references(refs)
    for r in report["results"]:
        print(f"[{r['reference']}] {r['verdict']:<10} {r['url'] or '-'}")
        print(f"      {r['detail']}")
        if r["found_title"] and r["verdict"] != OK:
            print(f"      found title: {r['found_title']}")
    print("\n" + ", ".join(f"{k}: {v}" for k, v in report["summary"].items() if v))
    if report["note"]:
        print(report["note"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
