"""tests/test_agent.py

Tests for the AI-agent layer: link_checker.py, agent_tools.Toolbox and
agent.py. Everything runs offline:

* registries (Crossref, Europe PMC, doi.org) and web pages are served by a
  local fake server;
* "Gemini" is a local server speaking the real Gemini REST format, driven
  through Google's real google-genai SDK (skipped if it isn't installed).

Run:  python -m pytest tests/ -v
"""
import hashlib
import json
import socket
import sys
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest
from docx import Document

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import link_checker as lc                                   # noqa: E402
from agent_tools import Toolbox                             # noqa: E402

try:
    import google.genai  # noqa: F401
    HAVE_GENAI = sys.version_info >= (3, 10)
except ImportError:
    HAVE_GENAI = False
needs_genai = pytest.mark.skipif(
    not HAVE_GENAI, reason="google-genai not installed (pip install -r requirements-agent.txt)")


# --------------------------------------------------------------------------
# Fake HTTP server (registries, web pages, Gemini)
# --------------------------------------------------------------------------
class FakeServer:
    """GET: ``routes`` {path: (status, content_type, body[, headers]) or
    callable(query) -> tuple}. POST: answered from ``post_script`` in order
    (item = response dict, or (status, dict))."""

    def __init__(self, routes=None, post_script=None):
        self.routes = routes or {}
        self.post_script = list(post_script or [])
        self.gets, self.posts = [], []
        server = self

        class Handler(BaseHTTPRequestHandler):
            def _send(self, status, ctype, body, headers=None):
                if isinstance(body, (dict, list)):
                    body = json.dumps(body)
                if isinstance(body, str):
                    body = body.encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                for k, v in (headers or {}).items():
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                parsed = urllib.parse.urlparse(self.path)
                server.gets.append(self.path)
                route = server.routes.get(urllib.parse.unquote(parsed.path))
                if callable(route):
                    route = route(urllib.parse.parse_qs(parsed.query))
                if route is None:
                    return self._send(404, "text/plain", "not found")
                self._send(*route)

            def do_POST(self):
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length) or b"{}")
                server.posts.append({"path": self.path, "headers": dict(self.headers),
                                     "body": body})
                if not server.post_script:
                    return self._send(500, "application/json", {"error": {
                        "code": 500, "message": "script exhausted", "status": "INTERNAL"}})
                item = server.post_script.pop(0)
                status, payload = item if isinstance(item, tuple) else (200, item)
                self._send(status, "application/json", payload)

            def log_message(self, *args):
                pass

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.base = f"http://127.0.0.1:{self.httpd.server_port}"
        threading.Thread(target=self.httpd.serve_forever, kwargs={"poll_interval": 0.05},
                         daemon=True).start()

    def close(self):
        self.httpd.shutdown()
        self.httpd.server_close()


def _closed_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


JSON = "application/json"
HTML = "text/html; charset=utf-8"
AI_TITLE = "Artificial intelligence adoption in SMEs"
REF_AI = ("3. S\u00e1nchez E, Calder\u00f3n R. Artificial intelligence adoption in SMEs: "
          "survey based on TOE\u2013DOI framework. Appl Sci. 2025;15(12):6465.")


def crossref(title, subtitle=None):
    rec = {"title": [title], "issued": {"date-parts": [[2025, 6]]}}
    if subtitle:
        rec["subtitle"] = [subtitle]
    return (200, JSON, {"status": "ok", "message": rec})


def _crossref_filter(query):
    if query.get("filter") == ["alternative-id:S2444569X25000320"]:
        return (200, JSON, {"message": {"items": [
            {"title": [AI_TITLE], "DOI": "10.1016/j.jik.2025.100687"}]}})
    return (200, JSON, {"message": {"items": []}})


def _europepmc(query):
    if query.get("query") == ["PMCID:PMC12749562"]:
        return (200, JSON, {"hitCount": 1, "resultList": {"result": [
            {"title": AI_TITLE + ".", "pubYear": "2025", "doi": "10.1/x"}]}})
    return (200, JSON, {"hitCount": 0, "resultList": {"result": []}})


PDF_XMP = (b"%PDF-1.6\n<x:xmpmeta><rdf:RDF><rdf:Description><dc:title><rdf:Alt>"
           b"<rdf:li xml:lang=\"x-default\">Artificial intelligence adoption in SMEs"
           b"</rdf:li></rdf:Alt></dc:title></rdf:Description></rdf:RDF></x:xmpmeta>\n")

ROUTES = {
    "/crossref/works/10.1234/good": crossref(
        "Artificial Intelligence Adoption in SMEs", "Survey Based on TOE-DOI Framework"),
    "/crossref/works/10.1234/other": crossref("Deep learning for protein structure prediction"),
    "/crossref/works/10.1234/partial": crossref(
        "Artificial intelligence and blockchain adoption in large manufacturing firms"),
    "/crossref/works/10.1108/IJSE-12-2021-0752": crossref(AI_TITLE),
    "/crossref/works": _crossref_filter,
    "/handles/10.9999/nothing": (404, JSON, {"responseCode": 100}),
    "/europepmc/search": _europepmc,
    "/page/good": (200, HTML, "<html><head><title>Journal site</title>"
                   f"<meta name=\"citation_title\" content=\"{AI_TITLE}\"></head>"
                   "<body><title>not this</title></body></html>"),
    "/page/sitename": (200, HTML, f"<html><head><title>{AI_TITLE} | Journal of "
                       "Things</title></head></html>"),
    "/page/unrelated": (200, HTML, "<title>Welcome to our homepage</title>"),
    "/page/soft404": (200, HTML, "<title>Page not found</title>"),
    "/page/blocked": (403, HTML, "<title>Forbidden</title>"),
    "/page/robot": (200, HTML, "<title>Just a moment...</title>"),
    "/page/pdf-xmp": (200, "application/pdf", PDF_XMP),
    "/page/pdf-plain": (200, "application/pdf", b"%PDF-1.4\n1 0 obj<<>>endobj\n"),
    "/page/moved": (301, HTML, "", {"Location": "/page/good"}),
}


@pytest.fixture
def web(monkeypatch):
    srv = FakeServer(dict(ROUTES))
    monkeypatch.setattr(lc, "CROSSREF_API", srv.base + "/crossref")
    monkeypatch.setattr(lc, "EUROPEPMC_API", srv.base + "/europepmc")
    monkeypatch.setattr(lc, "DOI_HANDLE_API", srv.base + "/handles")
    monkeypatch.setattr(lc, "TIMEOUT", 5)
    yield srv
    srv.close()


# --------------------------------------------------------------------------
# link_checker: pure functions
# --------------------------------------------------------------------------
@pytest.mark.parametrize("url, expected", [
    ("https://doi.org/10.3390/app15126465", "10.3390/app15126465"),
    ("http://dx.doi.org/10.1080/09537287.2022.2131620", "10.1080/09537287.2022.2131620"),
    ("https://link.springer.com/article/10.1007/s11156-024-01278-0", "10.1007/s11156-024-01278-0"),
    ("https://www.tandfonline.com/doi/full/10.1080/09537287.2022.2131620",
     "10.1080/09537287.2022.2131620"),
    ("https://link.springer.com/content/pdf/10.1007/s11156-024-01278-0.pdf",
     "10.1007/s11156-024-01278-0"),
    ("https://example.org/search?doi=10.5267%2Fj.ijdns.2023.12.021", "10.5267/j.ijdns.2023.12.021"),
    ("https://www.mdpi.com/2076-3417/15/12/6465", None),
    ("https://pmc.ncbi.nlm.nih.gov/articles/PMC12749562/", None),
    ("https://www.growingscience.com/ijds/Vol7/ijdns_2022_120.pdf", None),
])
def test_doi_extraction(url, expected):
    cands = lc.doi_candidates(url)
    assert (cands[0] if cands else None) == expected


def test_doi_trailing_publisher_path_is_trimmed():
    cands = lc.doi_candidates(
        "https://www.emerald.com/insight/content/doi/10.1108/IJSE-12-2021-0752/full/html")
    assert cands[0] == "10.1108/IJSE-12-2021-0752/full/html"
    assert "10.1108/IJSE-12-2021-0752" in cands
    # doi.org links are taken literally - a DOI may contain slashes
    assert lc.doi_candidates("https://doi.org/10.1002/a/b") == ["10.1002/a/b"]


def test_pii_and_pmcid_extraction():
    assert lc.extract_pii(
        "https://www.sciencedirect.com/science/article/pii/S2444569X25000320") == "S2444569X25000320"
    assert lc.extract_pii("https://www.cell.com/fulltext/S0092-8674(20)30230-2") is None
    assert lc.extract_pmcid("https://pmc.ncbi.nlm.nih.gov/articles/PMC12749562/") == "PMC12749562"
    assert lc.extract_pmcid("https://example.com/paper") is None


def test_html_title_priority_and_entities():
    page = (b"<html><head><title>Site | Home</title>"
            b"<meta property=\"og:title\" content=\"OG title\">"
            b"<meta name=\"citation_title\" content=\"Caf&eacute; &amp; <i>AI</i>\">"
            b"</head><body><meta name=\"citation_title\" content=\"late\"></body></html>")
    assert lc.html_title(page, HTML) == "Caf\u00e9 & AI"
    assert lc.html_title(b"<title>  Only\n title </title>") == "Only title"
    assert lc.html_title(b"<p>no title</p>") == ""


def test_pdf_title():
    assert lc.pdf_title(PDF_XMP) == AI_TITLE
    assert lc.pdf_title(b"%PDF-1.4 << /Title (Paper \\(draft\\)) >>") == "Paper (draft)"
    utf16 = "\u00c9tude".encode("utf-16-be")
    assert lc.pdf_title(b"%PDF << /Title (\xfe\xff" + utf16 + b") >>") == "\u00c9tude"
    assert lc.pdf_title(b"%PDF-1.4 nothing") == ""


def test_title_match():
    assert lc.title_match("Artificial Intelligence Adoption in SMEs", REF_AI) == 1.0
    # accents, dashes and case are normalised
    assert lc.title_match("Survey based on TOE-DOI framework", REF_AI) == 1.0
    assert lc.title_match("Deep learning for protein structure prediction", REF_AI) == 0.0
    assert lc.title_match("", REF_AI) == 0.0
    assert lc._page_title_match(f"{AI_TITLE} | Journal of Things", REF_AI) == 1.0


# --------------------------------------------------------------------------
# link_checker: against the fake registries / websites
# --------------------------------------------------------------------------
@pytest.mark.parametrize("url, verdict, source", [
    ("https://doi.org/10.1234/good", lc.OK, "Crossref"),
    ("https://doi.org/10.1234/other", lc.PROBLEM, "Crossref"),
    ("https://doi.org/10.1234/partial", lc.CHECK, "Crossref"),
    ("https://www.emerald.com/insight/content/doi/10.1108/IJSE-12-2021-0752/full/html",
     lc.OK, "Crossref"),
    ("https://www.sciencedirect.com/science/article/pii/S2444569X25000320", lc.OK, "Crossref"),
    ("https://pmc.ncbi.nlm.nih.gov/articles/PMC12749562/", lc.OK, "Europe PMC"),
])
def test_registry_checks(web, url, verdict, source):
    res = lc.check_link(url, REF_AI)
    assert res["verdict"] == verdict, res
    assert res["source"] == source
    assert res["found_title"]


def test_registry_result_details(web):
    res = lc.check_link("https://doi.org/10.1234/good", REF_AI)
    assert res["found_title"] == ("Artificial Intelligence Adoption in SMEs: "
                                  "Survey Based on TOE-DOI Framework")
    assert res["match"] == 1.0 and res["found_year"] == 2025 and res["doi"] == "10.1234/good"
    wrong = lc.check_link("https://doi.org/10.1234/other", REF_AI)
    assert "different paper" in wrong["detail"]
    assert "Deep learning for protein structure prediction" in wrong["detail"]


def test_nonexistent_doi_is_a_problem(web):
    res = lc.check_link("https://doi.org/10.9999/nothing", REF_AI)
    assert res["verdict"] == lc.PROBLEM
    assert "does not exist" in res["detail"]


@pytest.mark.parametrize("page, verdict", [
    ("good", lc.OK),           # citation_title meta wins over <title>
    ("sitename", lc.OK),       # "Title | Journal" - best segment
    ("moved", lc.OK),          # redirects are followed
    ("pdf-xmp", lc.OK),
    ("pdf-plain", lc.CHECK),   # opens, title unreadable
    ("unrelated", lc.CHECK),   # website titles never prove a "different paper"
    ("soft404", lc.PROBLEM),
    ("missing", lc.PROBLEM),   # HTTP 404
    ("blocked", lc.UNVERIFIED),
    ("robot", lc.UNVERIFIED),
])
def test_page_checks(web, page, verdict):
    res = lc.check_link(f"{web.base}/page/{page}", REF_AI)
    assert res["verdict"] == verdict, res


def test_check_references_batch(web):
    refs = {
        1: {"text": REF_AI, "url": "https://doi.org/10.1234/good"},
        2: {"text": REF_AI, "url": f"{web.base}/page/blocked"},
        3: {"text": "3. No link here. 2020.", "url": None},
        4: {"text": REF_AI, "url": "http://no-such-host.invalid/paper"},
    }
    report = lc.check_references(refs)
    verdicts = {r["reference"]: r["verdict"] for r in report["results"]}
    # the .invalid address can't resolve while other links worked -> gone
    assert verdicts == {1: lc.OK, 2: lc.UNVERIFIED, 3: lc.NO_LINK, 4: lc.PROBLEM}
    assert report["summary"] == {"OK": 1, "CHECK": 0, "PROBLEM": 1,
                                 "UNVERIFIED": 1, "NO_LINK": 1}
    assert report["checked"] == 4 and report["note"] == ""
    assert all(set(r) >= {"reference", "reference_text", "url", "verdict", "detail"}
               for r in report["results"])
    only = lc.check_references(refs, numbers=[3])
    assert [r["reference"] for r in only["results"]] == [3]


def test_offline_is_unverified_not_broken(monkeypatch):
    dead = f"http://127.0.0.1:{_closed_port()}"
    for attr in ("CROSSREF_API", "EUROPEPMC_API", "DOI_HANDLE_API"):
        monkeypatch.setattr(lc, attr, dead)
    refs = {1: {"text": REF_AI, "url": dead + "/paper"},
            2: {"text": REF_AI, "url": "http://no-such-host.invalid/paper"}}
    report = lc.check_references(refs)
    assert [r["verdict"] for r in report["results"]] == [lc.UNVERIFIED, lc.UNVERIFIED]
    assert "internet" in report["note"]


# --------------------------------------------------------------------------
# Toolbox
# --------------------------------------------------------------------------
def make_doc(path, body, refs, heading="References"):
    doc = Document()
    for text in body:
        doc.add_paragraph(text)
    if heading:
        doc.add_heading(heading, level=1)
    for ref in refs:
        doc.add_paragraph(ref)
    doc.save(str(path))
    return path


REFS = ["1. Alpha A. First paper. 2021. https://example.com/one",
        "2. Beta B. Second paper. 2022. https://example.com/two",
        "3. Gamma C. Third paper. 2023. https://example.com/three"]


@pytest.fixture
def box(tmp_path):
    inp = tmp_path / "input"
    inp.mkdir()
    make_doc(inp / "My Essay.docx", ["Intro [1] and [2].", "More [3]."], REFS)
    make_doc(inp / "Report.docx", ["Only [1]."], REFS)
    (inp / "~$My Essay.docx").write_bytes(b"lock")       # Word's lock file
    (inp / "notes.txt").write_text("x")
    return Toolbox(inp, tmp_path / "output")


def _sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def test_list_documents_skips_lock_and_other_files(box):
    res = box.call("list_documents")
    assert res["documents"] == ["My Essay.docx", "Report.docx"]
    assert res["output_files"] == []


@pytest.mark.parametrize("name, expected", [
    ("My Essay.docx", "My Essay.docx"),
    ("my essay", "My Essay.docx"),
    ("essay", "My Essay.docx"),
    ("REPORT.DOCX", "Report.docx"),
    ("C:\\Users\\me\\citation_agent\\input\\Report.docx", "Report.docx"),
])
def test_resolve_names(box, name, expected):
    assert box.resolve(name).name == expected


@pytest.mark.parametrize("name, fragment", [
    ("../../etc/passwd", "no document called"),
    ("..\\main.py", "no document called"),
    ("~$My Essay.docx", "no document called"),
    ("e", "matches several"),
    ("", "no document name"),
])
def test_resolve_refuses_everything_else(box, name, fragment):
    res = box.call("analyze_document", {"name": name})
    assert fragment in res["error"].lower()


def test_analyze_is_a_preview_that_saves_nothing(box):
    before = _sha(box.input_dir / "My Essay.docx")
    res = box.call("analyze_document", {"name": "essay"})
    assert res["document"] == "My Essay.docx" and res["preview_only"] is True
    assert res["citations_found"] == 3
    assert res["footnotes_that_would_be_added"] == 3
    assert res["problems"] == []
    assert [r["number"] for r in res["references"]] == [1, 2, 3]
    assert res["references"][0]["url"] == "https://example.com/one"
    assert not box.output_dir.exists()                  # nothing written
    assert _sha(box.input_dir / "My Essay.docx") == before


def test_add_footnotes_creates_copy_and_keeps_original(box):
    src = box.input_dir / "My Essay.docx"
    before = _sha(src)
    res = box.call("add_footnotes", {"name": "essay"})
    assert res["done"] is True and res["validation"] == "PASS"
    assert res["footnotes_added"] == 3 and res["original_unchanged"] is True
    assert res["output_file"] == "My Essay_with_footnotes.docx"
    assert (box.output_dir / res["output_file"]).is_file()
    assert _sha(src) == before
    again = box.call("add_footnotes", {"name": "essay", "keep_marker": True})
    assert again["output_file"] == "My Essay_with_footnotes_1.docx"   # never overwrites
    assert box.call("list_documents")["output_files"] == [
        "My Essay_with_footnotes.docx", "My Essay_with_footnotes_1.docx",
        "citation_review_report.csv"]


def test_problems_are_explained(box):
    make_doc(box.input_dir / "Broken.docx", ["A [1]. B [5]. C [2]."],
             ["1. Alpha A. First paper. 2021. https://example.com/one",
              "2. Beta B. Second paper without a link. 2022.",
              "3. Gamma C. Third paper. 2023. https://example.com/three"])
    res = box.call("add_footnotes", {"name": "Broken"})
    assert res["footnotes_added"] == 1 and res["validation"] == "PASS"
    assert res["references_without_link"] == [2]
    by_marker = {p["citation"]: p for p in res["problems"]}
    assert set(by_marker) == {"[5]", "[2]"}
    assert all(p["status"] == "ERROR" and p["meaning"] for p in res["problems"])
    assert "not found" in by_marker["[5]"]["why"].lower()
    assert "section" in by_marker["[5]"]["where"]


def test_no_reference_list_is_explained(box):
    make_doc(box.input_dir / "NoRefs.docx", ["Text [1]."], [], heading=None)
    res = box.call("add_footnotes", {"name": "NoRefs"})
    assert res["reference_list_found"] is False and "References" in res["message"]
    assert not list(box.output_dir.glob("NoRefs*.docx"))
    assert box.call("check_links", {"name": "NoRefs"})["reference_list_found"] is False


def test_check_links_tool(box, web):
    make_doc(box.input_dir / "Links.docx", ["Text [1], [2] and [3]."], [
        f"1. S\u00e1nchez E. {AI_TITLE}. 2025. {web.base}/page/good",
        "2. Other A. Some unrelated title. 2024. https://doi.org/10.1234/other",
        "3. No Link B. A reference without any link. 2023."])
    res = box.call("check_links", {"name": "Links"})
    assert res["document"] == "Links.docx"
    assert [(r["reference"], r["verdict"]) for r in res["results"]] == [
        (1, "OK"), (2, "PROBLEM"), (3, "NO_LINK")]
    # JSON numbers may arrive as floats; unknown numbers are reported
    some = box.call("check_links", {"name": "Links", "numbers": [2.0, 9]})
    assert [r["reference"] for r in some["results"]] == [2]
    assert some["not_in_reference_list"] == [9]


def test_call_rejects_unknown_tools_and_bad_arguments(box):
    assert "Unknown tool" in box.call("delete_everything", {})["error"]
    assert "error" in box.call("check_links", {"name": "essay", "numbers": ["x"]})
    # unexpected arguments are ignored, not passed through
    assert box.call("list_documents", {"path": "/etc"})["documents"]
    assert "error" in box.call("analyze_document", {})


def test_tool_results_are_json_serialisable(box, web):
    for name, args in [("list_documents", {}), ("analyze_document", {"name": "essay"}),
                       ("add_footnotes", {"name": "Report"})]:
        json.dumps(box.call(name, args))


# --------------------------------------------------------------------------
# agent.py with a fake Gemini
# --------------------------------------------------------------------------
def fc(*calls, sig=None):
    """A model turn with function calls: fc(("name", {args}, "id"), ...)."""
    parts = []
    for i, call in enumerate(calls):
        name, args, cid = (tuple(call) + (None, None))[:3]
        part = {"functionCall": {"name": name, "args": args or {}}}
        if cid:
            part["functionCall"]["id"] = cid
        if sig and i == 0:
            part["thoughtSignature"] = sig
        parts.append(part)
    return {"candidates": [{"content": {"role": "model", "parts": parts},
                            "finishReason": "STOP"}]}


def say(text):
    return {"candidates": [{"content": {"role": "model", "parts": [{"text": text}]},
                            "finishReason": "STOP"}]}


def gemini_error(code, status, message, reason=None):
    err = {"code": code, "message": message, "status": status}
    if reason:
        err["details"] = [{"@type": "type.googleapis.com/google.rpc.ErrorInfo",
                           "reason": reason}]
    return (code, {"error": err})


@pytest.fixture
def gemini():
    srv = FakeServer()
    yield srv
    srv.close()


def _agent(box, gemini, **kw):
    import agent
    from google import genai
    from google.genai import types
    client = genai.Client(api_key="test-key", http_options=types.HttpOptions(
        base_url=gemini.base + "/"))                     # no retries in tests
    return agent.CitationAgent(client, box, **kw)


@needs_genai
def test_agent_uses_tools_and_answers(box, gemini):
    gemini.post_script = [
        fc(("list_documents",), sig="c2lnLTE="),
        fc(("add_footnotes", {"name": "essay"}, "call-2"), sig="c2lnLTI="),
        say("Done! 3 footnotes added."),
    ]
    seen = []
    bot = _agent(box, gemini, on_tool=lambda name, args: seen.append((name, args)))
    assert bot.ask("please add footnotes to my essay") == "Done! 3 footnotes added."
    assert seen == [("list_documents", {}), ("add_footnotes", {"name": "essay"})]
    assert (box.output_dir / "My Essay_with_footnotes.docx").is_file()

    first = gemini.posts[0]
    assert first["path"].endswith("/models/gemini-flash-latest:generateContent")
    assert first["headers"].get("x-goog-api-key") == "test-key"
    body = first["body"]
    assert "Never invent" in body["systemInstruction"]["parts"][0]["text"]
    assert [f["name"] for f in body["tools"][0]["functionDeclarations"]] == [
        "list_documents", "analyze_document", "add_footnotes", "check_links"]

    turns = gemini.posts[2]["body"]["contents"]
    assert [t["role"] for t in turns] == ["user", "model", "user", "model", "user"]
    # Gemini 3 requires its thought signatures back, unchanged
    assert turns[1]["parts"][0]["thoughtSignature"] == "c2lnLTE="
    assert turns[3]["parts"][0]["thoughtSignature"] == "c2lnLTI="
    reply = turns[4]["parts"][0]["functionResponse"]
    assert reply["name"] == "add_footnotes" and reply["id"] == "call-2"
    assert reply["response"]["footnotes_added"] == 3
    assert reply["response"]["validation"] == "PASS"
    assert len(bot.history) == 6


@needs_genai
def test_agent_answers_parallel_calls_in_one_turn(box, gemini):
    gemini.post_script = [
        fc(("list_documents",), ("analyze_document", {"name": "Report"})),
        say("ok"),
    ]
    assert _agent(box, gemini).ask("what do I have?") == "ok"
    replies = gemini.posts[1]["body"]["contents"][-1]["parts"]
    assert [p["functionResponse"]["name"] for p in replies] == [
        "list_documents", "analyze_document"]
    assert replies[1]["functionResponse"]["response"]["citations_found"] == 1


@needs_genai
def test_tool_errors_go_back_to_the_model(box, gemini):
    gemini.post_script = [fc(("add_footnotes", {"name": "thesis"})),
                          say("I couldn't find a document called thesis.")]
    bot = _agent(box, gemini)
    assert "couldn't find" in bot.ask("footnotes for my thesis")
    reply = gemini.posts[1]["body"]["contents"][-1]["parts"][0]["functionResponse"]
    assert "no document called 'thesis'" in reply["response"]["error"].lower()


@needs_genai
@pytest.mark.parametrize("error, advice", [
    (gemini_error(429, "RESOURCE_EXHAUSTED", "Resource has been exhausted"), "free-tier limit"),
    (gemini_error(400, "INVALID_ARGUMENT", "API key not valid. Please pass a valid API key.",
                  "API_KEY_INVALID"), "--forget-key"),
    (gemini_error(404, "NOT_FOUND", "models/gemini-x is not found"), "--model"),
    (gemini_error(503, "UNAVAILABLE", "The model is overloaded."), "busy"),
])
def test_gemini_errors_become_advice_and_roll_back(box, gemini, error, advice):
    import agent
    gemini.post_script = [error]
    bot = _agent(box, gemini)
    with pytest.raises(Exception) as exc:
        bot.ask("hello")
    assert advice in agent.friendly_error(exc.value, "gemini-x")
    assert bot.history == []                             # conversation still valid


def test_network_errors_become_advice():
    import agent
    assert "internet" in agent.friendly_error(ConnectionError("down"))
    assert "internet" in agent.friendly_error(TimeoutError())


@needs_genai
def test_quota_message_mentioning_api_key_is_still_a_rate_limit():
    import agent
    from google.genai import errors
    exc = errors.ClientError(429, {"error": {"code": 429, "status": "RESOURCE_EXHAUSTED",
                                             "message": "Quota exceeded for this API key."}})
    assert "free-tier limit" in agent.friendly_error(exc)


def test_plain_removes_markdown():
    import agent
    md = "## Result\n**3 footnotes** added to __Essay__.\n* one\n  * two\n\u2022 three\n- kept"
    assert agent.plain(md) == ("Result\n3 footnotes added to Essay.\n- one\n  - two\n"
                               "- three\n- kept")
    assert agent.plain("2 * 3 = 6, a*b") == "2 * 3 = 6, a*b"


@needs_genai
def test_empty_answer_and_step_limit_roll_back(box, gemini):
    gemini.post_script = [{"candidates": [{"finishReason": "SAFETY"}]}]
    bot = _agent(box, gemini)
    assert "empty answer" in bot.ask("hi") and bot.history == []
    gemini.post_script = [fc(("list_documents",))] * 3
    bot = _agent(box, gemini, max_steps=3)
    assert "too many steps" in bot.ask("loop") and bot.history == []


@needs_genai
def test_main_once_end_to_end(box, gemini, monkeypatch, capsys):
    import agent
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.setenv("GOOGLE_GEMINI_BASE_URL", gemini.base + "/")
    gemini.post_script = [fc(("add_footnotes", {"name": "Report"})), say("All done.")]
    code = agent.main(["--once", "footnotes for Report", "--model", "gemini-test",
                       "--input-dir", str(box.input_dir), "--output-dir", str(box.output_dir)])
    out = capsys.readouterr().out
    assert code == 0
    assert "... adding footnotes to \"Report\"" in out and "All done." in out
    assert (box.output_dir / "Report_with_footnotes.docx").is_file()
    assert "/models/gemini-test:generateContent" in gemini.posts[0]["path"]


def test_api_key_sources(tmp_path, monkeypatch, capsys):
    import agent
    key_file = tmp_path / "home" / "gemini_api_key.txt"
    monkeypatch.setattr(agent, "KEY_FILE", key_file)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    assert agent.get_api_key(interactive=False) is None
    key_file.parent.mkdir()
    key_file.write_text("  saved-key\n", encoding="utf-8")
    assert agent.get_api_key(interactive=False) == "saved-key"
    monkeypatch.setenv("GEMINI_API_KEY", "env-key")
    assert agent.get_api_key(interactive=False) == "env-key"
    assert agent.main(["--forget-key"]) == 0 and not key_file.exists()
    assert "Deleted" in capsys.readouterr().out


def test_default_key_location_is_outside_the_project():
    import agent
    assert agent.HERE not in Path(agent.KEY_FILE).resolve().parents
