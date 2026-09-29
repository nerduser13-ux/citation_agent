"""validator.py

Pre-save / post-save validation. The tool never fabricates data, and it must
prove that:
  * the input file was untouched (SHA-256),
  * the output is a valid, reopenable .docx with real OOXML footnotes,
  * every NEW footnote contains exactly the URL of its reference,
  * existing footnotes were preserved (definition content + body references),
  * no text was lost anywhere (per-paragraph raw-text fidelity),
  * tables, hyperlinks, bookmarks, fields, drawings, content controls,
    section breaks and page layout are preserved (structural counts),
  * citation markers were removed exactly where processing succeeded
    (and left alone everywhere else).
"""
import hashlib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

from docx.oxml.ns import qn

from document_reader import raw_text_of


def file_hash(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def original_unchanged(path, original_hash) -> bool:
    try:
        return file_hash(path) == original_hash
    except OSError:
        return False


# --------------------------------------------------------------------------
# OOXML helpers
# --------------------------------------------------------------------------
_STRUCTURAL_TAGS = (
    "w:p", "w:tbl", "w:hyperlink", "w:bookmarkStart", "w:bookmarkEnd",
    "w:fldChar", "w:drawing", "w:sdt", "w:object", "w:pgSz", "w:sectPr",
    "w:ins", "w:del", "w:pPr", "w:lastRenderedPageBreak",
)


def structural_counts(body_elem) -> dict:
    return {t: len(list(body_elem.iter(qn(t)))) for t in _STRUCTURAL_TAGS}


def direct_to_raw_map(p_elem):
    """Map direct-run coordinates to raw (all w:t) coordinates.

    Returns a list of (direct_start, direct_end, raw_start, raw_end), one entry
    per direct <w:r> child. Non-run children (hyperlinks, inline sdt, ...)
    contribute raw text only - exactly the difference between the two
    coordinate spaces."""
    spans = []
    direct_pos = 0
    raw_pos = 0
    for child in p_elem:
        if child.tag == qn("w:r"):
            d0, r0 = direct_pos, raw_pos
            for t in child.findall(qn("w:t")):
                direct_pos += len(t.text or "")
                raw_pos += len(t.text or "")
            spans.append((d0, direct_pos, r0, raw_pos))
        else:
            raw_pos += len(raw_text_of(child))
    return spans


def hyperlink_targets(doc) -> List[str]:
    """External target refs of every w:hyperlink in the body (rels resolved)."""
    part = doc.part
    out = []
    for hl in doc.element.body.iter(qn("w:hyperlink")):
        rid = hl.get(qn("r:id"))
        if rid and rid in part.rels:
            try:
                out.append(part.rels[rid].target_ref)
            except Exception:
                pass
    return out


# --------------------------------------------------------------------------
# Context + entry point
# --------------------------------------------------------------------------
@dataclass
class ValidationContext:
    in_path: str
    out_path: str
    in_hash: str
    ref_map: dict
    citations: list
    replace_marker: bool
    existing: object            # FootnoteSnapshot
    input_raw_texts: list       # raw text per paragraph, input-side
    new_footnotes: dict         # {fid: url} allocated by this run
    parser_warnings: list = field(default_factory=list)


def validate_output(ctx: ValidationContext) -> dict:
    """Run every post-save check. Returns {"ok": bool, "failures": [...],
    "results": {name: value}} - one results entry per reported metric."""
    res = {}
    failures = []

    def check(name, ok, detail=""):
        res[name] = bool(ok)
        if not ok:
            failures.append(f"{name}{': ' + detail if detail else ''}")

    # 1) output exists + valid zip ----------------------------------------
    check("output_exists", Path(ctx.out_path).is_file())
    check("valid_zip", zipfile.is_zipfile(ctx.out_path)
          and zipfile.ZipFile(ctx.out_path).testzip() is None)

    # 2) reopen with python-docx -------------------------------------------
    try:
        from docx import Document
        out_doc = Document(ctx.out_path)
        check("reopen_ok", True)
    except Exception as e:      # noqa: BLE001
        check("reopen_ok", False, str(e))
        return _result(res, failures)

    # 3) footnotes.xml present + real footnote definitions ------------------
    z = zipfile.ZipFile(ctx.out_path)
    ft = None
    for n in z.namelist():
        if n.endswith("footnotes.xml"):
            from lxml import etree
            ft = etree.fromstring(z.read(n))
    expect_regular = len(ctx.existing.definitions) + len(ctx.new_footnotes)
    if expect_regular > 0:
        check("footnotes_xml_present", ft is not None)
        if ft is None:
            return _result(res, failures)
    else:
        check("footnotes_xml_present", True)
    if ft is None:
        res["regular_footnote_count"] = 0
        check("regular_footnote_count", expect_regular == 0)
        return _result(res, failures)

    regular = [f for f in ft.findall(qn("w:footnote"))
               if f.get(qn("w:id")) not in ("-1", "0")]
    res["regular_footnote_count"] = len(regular)
    check("regular_footnote_count", len(regular) == expect_regular,
          f"expected {expect_regular}, found {len(regular)}")

    # 4) body footnote references ------------------------------------------
    body_refs = [r.get(qn("w:id")) for r in out_doc.element.body.iter(qn("w:footnoteReference"))]
    expect_refs = len(ctx.existing.body_reference_ids) + len(ctx.new_footnotes)
    res["body_reference_count"] = len(body_refs)
    check("body_reference_count", len(body_refs) == expect_refs,
          f"expected {expect_refs}, found {len(body_refs)}")

    # 5) every NEW footnote carries EXACTLY its reference URL ---------------
    defs = {f.get(qn("w:id")): f for f in regular}
    bad_new = []
    for fid, url in ctx.new_footnotes.items():
        f = defs.get(str(fid))
        if f is None:
            bad_new.append(f"id {fid} missing")
            continue
        text = "".join(t.text or "" for t in f.iter(qn("w:t")))
        if text != url:
            bad_new.append(f"id {fid}: {text!r} != {url!r}")
    check("new_footnote_urls_exact", not bad_new, "; ".join(bad_new))

    new_urls = set(ctx.new_footnotes.values())
    ref_urls = {v["url"] for v in ctx.ref_map.values() if v["url"]}
    check("all_new_urls_from_refs", new_urls.issubset(ref_urls),
          f"fabricated: {sorted(new_urls - ref_urls)}")

    # 6) existing footnotes preserved (content + body refs) -----------------
    lost = [fid for fid, text in ctx.existing.definitions.items()
            if fid not in defs or
            "".join(t.text or "" for t in defs[fid].iter(qn("w:t"))) != text]
    missing_refs = [fid for fid in ctx.existing.body_reference_ids
                    if fid not in set(body_refs)]
    check("existing_preserved", not lost and not missing_refs,
          f"lost defs: {lost}; missing body refs: {missing_refs}")

    # 7) per-paragraph text fidelity (incl. marker removal check) -----------
    from collections import defaultdict
    from document_reader import iter_paragraph_info
    from docx import Document as _Doc
    in_doc = _Doc(ctx.in_path)
    in_p_elems = [p._p for p, _w in iter_paragraph_info(in_doc)]
    out_paras = [p._p for p, _w in iter_paragraph_info(out_doc)]
    if len(in_p_elems) != len(out_paras):
        check("text_fidelity", False,
              f"paragraph count changed ({len(in_p_elems)} -> {len(out_paras)})")
        return _result(res, failures)
    cites_by_para = defaultdict(list)
    for c in ctx.citations:
        cites_by_para[c.para_index].append(c)
    misfit = []
    for idx, (in_p, out_p) in enumerate(zip(in_p_elems, out_paras)):
        expected = ctx.input_raw_texts[idx]
        for c in sorted(cites_by_para.get(idx, []),
                        key=lambda c: c.start, reverse=True):
            if c.status != "SUCCESS" or c.start < 0 or not ctx.replace_marker:
                continue
            for (d0, d1, r0, r1) in direct_to_raw_map(in_p):
                if c.start >= d0 and c.end <= d1:
                    rs, re = r0 + (c.start - d0), r0 + (c.end - d0)
                    expected = expected[:rs] + expected[re:]
                    break
        actual = raw_text_of(out_p)
        if actual != expected:
            misfit.append(f"para {idx}: {actual!r} != {expected!r}")
    check("text_fidelity", not misfit, "; ".join(misfit[:3]))

    # 8) structural preservation -------------------------------------------
    in_counts = structural_counts(in_doc.element.body)
    out_counts = structural_counts(out_doc.element.body)
    diff = {t: (in_counts[t], out_counts[t]) for t in _STRUCTURAL_TAGS
            if in_counts[t] != out_counts[t]}
    check("structural_preserved", not diff, str(diff))

    # 9) hyperlinks preserved ----------------------------------------------
    in_hl = hyperlink_targets(in_doc)
    out_hl = hyperlink_targets(out_doc)
    check("hyperlinks_preserved",
          set(in_hl).issubset(set(out_hl)) and in_counts["w:hyperlink"] == out_counts["w:hyperlink"],
          f"missing {sorted(set(in_hl) - set(out_hl))}")

    # 10) input file untouched ----------------------------------------------
    check("input_unchanged", original_unchanged(ctx.in_path, ctx.in_hash))

    return _result(res, failures)


def _result(res, failures):
    return {"ok": not failures, "failures": failures, "results": res}
