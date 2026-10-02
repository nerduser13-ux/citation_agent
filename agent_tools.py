"""agent_tools.py

Thin, narrow wrappers around the deterministic pipeline. An AI agent (future
stage) can call these tools instead of manipulating the document directly. Each
tool has a single, explicit responsibility and returns plain data - the agent
is never given unrestricted filesystem or shell control. It cannot create,
delete or rewrite arbitrary files: the only file operations available to it are
"read this .docx and report" (inspect_document) and "run the deterministic
pipeline on one .docx into the standard output path" (run_pipeline).

``Toolbox`` (bottom of this file) is what the AI agent in ``agent.py``
actually gets: four tools with JSON-friendly inputs/outputs -
list_documents, analyze_document, add_footnotes, check_links. File access is
limited to .docx files that exist in the input folder (looked up by listing
the folder, never by opening a path the model supplied); the only write is
the pipeline's new copy in the output folder.

Future Mendeley stage can build on the same shape:
    citation_marker -> reference_url -> identify publication -> Mendeley ref
    -> Mendeley Cite citation -> updated bibliography
(currently stops at: citation_marker -> reference_url -> Word footnote)
"""
import contextlib
import csv
import io
import re
import tempfile
from pathlib import Path

import link_checker
from config import Config
from document_reader import DocumentReader, is_heading
from reference_parser import ReferenceParser
from citation_detector import CitationDetector
from validator import file_hash, validate_output, ValidationContext
from report import Report


def inspect_document(path):
    """Return a high-level description of the document (counts, headings,
    existing footnotes). Read-only; no files are created or modified."""
    reader = DocumentReader(path)
    paras = reader.all_paragraphs()
    headings = [p.text for p in paras if is_heading(p)]
    snap = reader.existing_footnotes()
    return {
        "path": str(path),
        "sha256": file_hash(path),
        "paragraph_count": len(paras),
        "existing_footnotes": len(snap.definitions),
        "headings": headings[:50],
    }


def find_reference_section(paragraphs, config):
    """Locate the reference section. Returns (start_index, method, confidence,
    note) or None (-> caller must abort, no guessing)."""
    return ReferenceParser(config).find_reference_section(paragraphs, config)


def extract_references(paragraphs, start_index):
    """Extract the number -> {url, text} map from the reference list."""
    return ReferenceParser().parse_references(paragraphs, start_index)


def find_citations(paragraphs, ref_section_start):
    """Detect citation markers in the body (index < ref_section_start)."""
    return CitationDetector().detect(paragraphs, ref_section_start)


def validate_citation(citation, ref_map):
    """Return (ok, reason) for a citation against the reference map."""
    for num in citation.numbers:
        if num not in ref_map:
            return False, f"Reference {num} not found"
        if not ref_map[num]["url"]:
            return False, f"Reference {num} has no URL"
    return True, ""


def verify_output(ctx: ValidationContext):
    """Run the full post-save validation. Returns {"ok", "failures", "results"}."""
    return validate_output(ctx)


def generate_report(path, rows, summary):
    """Write the CSV review report at ``path`` (the only file an agent may
    write through these tools)."""
    rep = Report()
    for row in rows:
        rep.add(*row)
    rep.set_summary(**summary)
    rep.write(path)
    return path


def run_pipeline(input_path, output_path=None, backend="xml",
                 keep_marker=False, overwrite=False,
                 input_dir="input", output_dir="output"):
    """One-call deterministic pipeline (the only way an agent may produce a
    modified document). Returns (exit_code, output_path, report_path).

    ``output_path`` is the file actually written - including the
    collision-safe ``..._with_footnotes_1.docx`` name - or None when no
    document was produced (e.g. aborted: reference section not found).
    ``report_path`` is None when no report was written (usage errors)."""
    from pathlib import Path
    from main import run  # local import: keep module import side-effect free

    config = Config()
    config.input_path = input_path
    config.output_path = output_path
    config.backend = backend
    config.replace_marker = not keep_marker
    config.overwrite_output = overwrite
    config.input_dir = Path(input_dir)
    config.output_dir = Path(output_dir)
    code = run(config)
    out = config.resolved_output_path
    if out is not None and not Path(out).is_file():
        out = None
    return code, out, config.report_path


# ---------------------------------------------------------------------------
# Tools for the AI agent (agent.py)
# ---------------------------------------------------------------------------
class ToolError(Exception):
    """A problem the agent should explain to the user (unknown file, ...)."""


_NAME_PARAM = {
    "type": "string",
    "description": ("Document file name as shown by list_documents. A unique "
                    "part of the name is enough (e.g. 'RALF')."),
}

# Plain-English meaning of the pipeline's statuses, returned with problems.
# add_footnotes(separator=...) -> text between footnotes placed side by side
SEPARATORS = {"space": " ", "comma": ",", "none": ""}

STATUS_HELP = {
    "ERROR": ("Not changed: the reference number is missing from the reference "
              "list, or that reference has no link. Fix the reference list."),
    "AMBIGUOUS": ("Left unchanged for safety (e.g. the [n] is a hyperlink, a "
                  "reference-manager field, or split by formatting). Add this "
                  "footnote by hand in Word if needed."),
    "SKIPPED": "Already has a footnote from an earlier run - nothing to do.",
}


def _short(text, limit=220):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text if len(text) <= limit else text[:limit - 3] + "..."


def _read_report(path):
    """Parse citation_review_report.csv -> (rows as dicts, summary dict)."""
    rows, summary = [], {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader, None) or []
        in_summary = False
        for rec in reader:
            if not rec:
                continue
            if rec == ["SUMMARY"]:
                in_summary = True
            elif in_summary:
                summary[rec[0]] = rec[1] if len(rec) > 1 else ""
            else:
                rows.append(dict(zip(header, rec)))
    return rows, summary


def _int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _number_list(value):
    """"3, 7" / "0" -> [3, 7] / []"""
    return [int(x) for x in re.findall(r"\d+", str(value or "")) if int(x) > 0]


class Toolbox:
    """Everything the AI agent can do - and nothing else.

    * reads only .docx files that exist in ``input_dir``;
    * writes only through the deterministic pipeline (a NEW file in
      ``output_dir``; previews use a temporary folder that is deleted);
    * returns plain JSON-serialisable dicts; problems come back as
      ``{"error": ...}`` so the model can explain them instead of crashing.

    The model never sees the body text of a document: only file names,
    section headings, citation markers, the reference list and results.
    """

    TOOL_SPECS = [
        {
            "name": "list_documents",
            "description": (
                "List the Word documents (.docx) in the input folder, and the "
                "files already created in the output folder. Call this first "
                "when the user doesn't give an exact file name."),
            "parameters": None,
        },
        {
            "name": "analyze_document",
            "description": (
                "Preview a document WITHOUT saving anything: finds the "
                "reference list and the numbered citations such as [1] or "
                "[2]-[4], and reports exactly which citations would get "
                "footnotes and which have problems (and why). Also returns the "
                "reference list with its links."),
            "parameters": {"type": "object",
                           "properties": {"name": _NAME_PARAM},
                           "required": ["name"]},
        },
        {
            "name": "add_footnotes",
            "description": (
                "Create a NEW copy of the document in the output folder in "
                "which every numbered citation gets a real Word footnote "
                "containing the source link from the reference list. The "
                "original file is never modified. Several references cited "
                "together ([1,2], [1-3], or [1], [2]) get their footnotes side "
                "by side. Returns the new file name, counts, the validation "
                "result and any citations that could not be processed."),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": _NAME_PARAM,
                    "keep_marker": {
                        "type": "boolean",
                        "description": (
                            "true = keep the [n] text and add the footnote "
                            "after it; false (default) = replace [n] with the "
                            "footnote number."),
                    },
                    "separator": {
                        "type": "string",
                        "enum": list(SEPARATORS),
                        "description": (
                            "What goes between footnotes placed side by side: "
                            "space (default, footnotes 1 2), comma (1,2) or "
                            "none (12)."),
                    },
                },
                "required": ["name"],
            },
        },
        {
            "name": "check_links",
            "description": (
                "Check the links in the document's reference list: does each "
                "one work and lead to the paper the reference names? Uses the "
                "official DOI registry (Crossref) where possible. Can take up "
                "to a minute. Verdicts: OK; CHECK (compare by eye); PROBLEM "
                "(broken, or a different paper); UNVERIFIED (the site blocks "
                "automatic checks - open it in a browser); NO_LINK."),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": _NAME_PARAM,
                    "numbers": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "Optional: only check these reference numbers.",
                    },
                },
                "required": ["name"],
            },
        },
    ]

    def __init__(self, input_dir="input", output_dir="output"):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)

    # ---- file access -------------------------------------------------------
    def documents(self):
        """.docx files in the input folder (Word's ~$ lock files excluded)."""
        if not self.input_dir.is_dir():
            return []
        return sorted((p for p in self.input_dir.iterdir()
                       if p.is_file() and p.suffix.lower() == ".docx"
                       and not p.name.startswith(("~$", "."))),
                      key=lambda p: p.name.lower())

    def resolve(self, name):
        """Map a name from the model to a file IN the input folder. Only the
        last path component is used and it is matched against the folder
        listing, so no other file can ever be reached."""
        docs = self.documents()
        if not docs:
            raise ToolError("There are no Word (.docx) files in the input folder "
                            f"({self.input_dir.resolve()}). Copy the document there first.")
        wanted = re.split(r"[\\/]", str(name or "").strip().strip("\"'"))[-1].strip()
        if not wanted:
            raise ToolError("No document name was given.")
        low = wanted.lower()
        stem = low[:-5] if low.endswith(".docx") else low
        for p in docs:
            if p.name.lower() in (low, low + ".docx"):
                return p
        matches = [p for p in docs if stem and stem in p.name.lower()]
        if len(matches) == 1:
            return matches[0]
        if matches:
            raise ToolError(f"'{wanted}' matches several documents: "
                            f"{[p.name for p in matches]}. Which one is meant?")
        raise ToolError(f"There is no document called '{wanted}' in the input "
                        f"folder. Available: {[p.name for p in docs]}")

    # ---- shared helpers ----------------------------------------------------
    @staticmethod
    def _references(path):
        """(section info or None, {n: {"text", "url"}}, parser warnings)"""
        cfg = Config()
        paras = DocumentReader(path).all_paragraphs()
        found = ReferenceParser(cfg).find_reference_section(paras, cfg)
        if found is None:
            return None, {}, []
        start, method, confidence, note = found
        ref_map, warnings = ReferenceParser().parse_references(paras, start)
        info = {"how_found": note, "method": method, "confidence": confidence}
        return info, ref_map, list(warnings)

    @staticmethod
    def _run(path, output_dir, keep_marker, separator="space"):
        """Run the real pipeline (main.run) quietly and collect its results."""
        from main import run  # local import: keep module import side-effect free
        cfg = Config()
        cfg.input_path = str(path)
        cfg.output_dir = Path(output_dir)
        cfg.replace_marker = not keep_marker
        cfg.footnote_separator = SEPARATORS[separator]
        log = io.StringIO()
        with contextlib.redirect_stdout(log):
            code = run(cfg)
        rows, summary = ([], {})
        if cfg.report_path and Path(cfg.report_path).is_file():
            rows, summary = _read_report(cfg.report_path)
        out = cfg.resolved_output_path
        if out is not None and not Path(out).is_file():
            out = None
        return {"code": code, "output": out, "report": cfg.report_path,
                "rows": rows, "summary": summary,
                "log": log.getvalue().splitlines()}

    @staticmethod
    def _problems(rows):
        problems = []
        for r in rows:
            status = r.get("Status", "")
            if status == "SUCCESS":
                continue
            problems.append({
                "citation": r.get("Citation Text", ""),
                "reference_number": r.get("Citation Number", ""),
                "where": r.get("Location", ""),
                "status": status,
                "why": r.get("Notes", ""),
                "meaning": STATUS_HELP.get(status, ""),
            })
        return problems[:60]

    @staticmethod
    def _no_reference_list(path):
        return {
            "document": path.name,
            "reference_list_found": False,
            "message": (
                "No reference list was found, so nothing can be done. The list "
                "needs a heading such as 'References' or 'Bibliography' (bold "
                "or a Heading style) directly above numbered entries like "
                "'1. Author. Title. https://...'."),
        }

    # ---- the tools ---------------------------------------------------------
    def list_documents(self):
        outputs = []
        if self.output_dir.is_dir():
            outputs = sorted(p.name for p in self.output_dir.iterdir()
                             if p.is_file() and not p.name.startswith(("~$", ".")))
        return {
            "input_folder": str(self.input_dir.resolve()),
            "documents": [p.name for p in self.documents()],
            "output_folder": str(self.output_dir.resolve()),
            "output_files": outputs,
        }

    def analyze_document(self, name):
        path = self.resolve(name)
        info, refs, warnings = self._references(path)
        if info is None:
            return self._no_reference_list(path)
        with tempfile.TemporaryDirectory() as tmp:     # preview: nothing kept
            res = self._run(path, tmp, keep_marker=False)
        s = res["summary"]
        return {
            "document": path.name,
            "preview_only": True,
            "reference_list_found": True,
            "reference_list": info,
            "references": [{"number": n, "text": _short(r["text"]),
                            "url": r["url"]} for n, r in sorted(refs.items())],
            "citations_found": _int(s.get("Citation_markers_detected")),
            "footnotes_that_would_be_added": _int(s.get("Footnotes_inserted")),
            "existing_footnotes_in_document": _int(s.get("Existing_footnotes_detected")),
            "references_without_link": _number_list(s.get("Missing_URLs")),
            "references_never_cited": _number_list(s.get("Uncited_references")),
            "problems": self._problems(res["rows"]),
            "reference_list_warnings": warnings,
        }

    def add_footnotes(self, name, keep_marker=False, separator="space"):
        path = self.resolve(name)
        if separator not in SEPARATORS:
            raise ToolError(f"separator must be one of {list(SEPARATORS)}")
        res = self._run(path, self.output_dir, keep_marker=bool(keep_marker),
                        separator=separator)
        s = res["summary"]
        if res["code"] == 3:
            return {"document": path.name, "done": False, "message": (
                "Stopped: the original file changed while it was being "
                "processed (is it open and being edited?). Close it and try again.")}
        if res["output"] is None:
            if s.get("Reference_section") == "NOT FOUND":
                return self._no_reference_list(path)
            reason = next((l for l in res["log"] if l.startswith("ERROR")), "")
            return {"document": path.name, "done": False,
                    "message": f"No new document was created. {reason}".strip()}
        failures = [l.split(":", 1)[1].strip() for l in res["log"]
                    if l.startswith("VALIDATION FAILURE:")]
        result = {
            "document": path.name,
            "done": res["code"] == 0,
            "output_file": res["output"].name,
            "output_folder": str(res["output"].parent.resolve()),
            "report_file": Path(res["report"]).name if res["report"] else None,
            "footnotes_added": _int(s.get("Footnotes_inserted")),
            "citations_found": _int(s.get("Citation_markers_detected")),
            "validation": "PASS" if res["code"] == 0 else "FAIL",
            "original_unchanged": s.get("Input_unchanged") == "True",
            "problems": self._problems(res["rows"]),
            "references_without_link": _number_list(s.get("Missing_URLs")),
        }
        if failures:
            result["validation_failures"] = failures
            result["warning"] = ("Validation failed - the new file should not be "
                                 "used. The original is unchanged.")
        return result

    def check_links(self, name, numbers=None):
        path = self.resolve(name)
        info, refs, _ = self._references(path)
        if info is None:
            return self._no_reference_list(path)
        if numbers is not None and not isinstance(numbers, (list, tuple)):
            numbers = [numbers]
        wanted = [int(n) for n in (numbers or [])]
        report = link_checker.check_references(refs, wanted or None)
        report["document"] = path.name
        if wanted:
            report["not_in_reference_list"] = sorted(set(wanted) - set(refs))
        return report

    # ---- dispatch ----------------------------------------------------------
    def call(self, name, args=None):
        """Run tool ``name`` with ``args`` from the model. Always returns a
        dict; errors become ``{"error": ...}``."""
        spec = next((t for t in self.TOOL_SPECS if t["name"] == name), None)
        if spec is None:
            return {"error": f"Unknown tool '{name}'."}
        allowed = set(((spec["parameters"] or {}).get("properties") or {}))
        kwargs = {k: v for k, v in dict(args or {}).items() if k in allowed}
        try:
            if "numbers" in kwargs and kwargs["numbers"] is not None:
                nums = kwargs["numbers"]
                nums = nums if isinstance(nums, (list, tuple)) else [nums]
                kwargs["numbers"] = [int(float(n)) for n in nums]  # JSON may send 3.0
            return getattr(self, name)(**kwargs)
        except ToolError as e:
            return {"error": str(e)}
        except TypeError as e:
            return {"error": f"Wrong arguments for {name}: {e}"}
        except Exception as e:  # one bad document must not end the chat
            return {"error": f"{name} failed: {type(e).__name__}: {e}"}
