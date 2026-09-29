"""word_footnotes.py

Inserts REAL Microsoft Word footnotes while preserving the existing document as
much as technically possible.

Two backends:
  * XmlFootnoteBackend  - cross-platform OOXML manipulation (DEFAULT, validated).
  * process_document_com - Windows Microsoft Word COM automation (EXPERIMENTAL;
                           untested in this Linux environment).

Design principles (document safety first):
  * We never rebuild a whole paragraph. We only split the single <w:r> run that
    actually contains a citation marker, copying its formatting to the surrounding
    text pieces and inserting genuine <w:footnoteReference> runs in its place.
  * Hyperlinks, fields, drawings, bookmarks, tables, headings, lists and existing
    footnotes are left entirely untouched.
  * A marker that spans runs, or sits inside a run that contains a drawing/field,
    is NOT modified - it is reported as ambiguous instead of being destroyed.
"""
import copy
from collections import defaultdict
from lxml import etree

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.opc.constants import RELATIONSHIP_TYPE as RT

FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# --------------------------------------------------------------------------
# Styles + low-level OOXML helpers
# --------------------------------------------------------------------------
def ensure_footnote_styles(doc):
    """Make sure FootnoteReference (character) and FootnoteText (paragraph)
    styles exist in the document."""
    styles_root = doc.styles.element
    existing = {el.get(qn("w:styleId"))
                for el in styles_root.findall(qn("w:style"))}

    def add(style_id, style_type, name_val):
        if style_id in existing:
            return
        st = OxmlElement("w:style")
        st.set(qn("w:type"), style_type)
        st.set(qn("w:styleId"), style_id)
        nm = OxmlElement("w:name")
        nm.set(qn("w:val"), name_val)
        st.append(nm)
        styles_root.append(st)
        existing.add(style_id)

    add("FootnoteReference", "character", "footnote reference")
    add("FootnoteText", "paragraph", "footnote text")


def make_text_run(text, rpr):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def make_footnote_ref_run(fid):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rs = OxmlElement("w:rStyle")
    rs.set(qn("w:val"), "FootnoteReference")
    rPr.append(rs)
    r.append(rPr)
    fr = OxmlElement("w:footnoteReference")
    fr.set(qn("w:id"), str(fid))
    r.append(fr)
    return r


def add_footnote_def(root, fid, url):
    ft = OxmlElement("w:footnote")
    ft.set(qn("w:id"), str(fid))
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), "FootnoteText")
    pPr.append(ps)
    p.append(pPr)
    # footnote reference mark (Word renders the auto number here)
    r1 = OxmlElement("w:r")
    rPr1 = OxmlElement("w:rPr")
    rs1 = OxmlElement("w:rStyle")
    rs1.set(qn("w:val"), "FootnoteReference")
    rPr1.append(rs1)
    r1.append(rPr1)
    r1.append(OxmlElement("w:footnoteRef"))
    p.append(r1)
    # URL text
    r2 = OxmlElement("w:r")
    rPr2 = OxmlElement("w:rPr")
    rs2 = OxmlElement("w:rStyle")
    rs2.set(qn("w:val"), "FootnoteText")
    rPr2.append(rs2)
    r2.append(rPr2)
    t2 = OxmlElement("w:t")
    t2.set(qn("xml:space"), "preserve")
    t2.text = url
    r2.append(t2)
    p.append(r2)
    ft.append(p)
    root.append(ft)


# --------------------------------------------------------------------------
# OOXML backend (default, cross-platform)
# --------------------------------------------------------------------------
class XmlFootnoteBackend:
    def __init__(self, doc):
        self.doc = doc
        self.package = doc.part.package
        self.part, self.root = self._ensure_part()
        self.next_id = self._compute_next_id()
        self.inserted = 0

    def _ensure_part(self):
        for part in self.package.iter_parts():
            if str(part.partname).endswith("footnotes.xml"):
                return part, etree.fromstring(part.blob)
        sep = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0">'
            '<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            '</w:footnotes>')
        part = Part(PackURI("/word/footnotes.xml"), FOOTNOTES_CT,
                    sep.encode("utf-8"), self.package)
        self.doc.part.relate_to(part, RT.FOOTNOTES)
        return part, etree.fromstring(sep.encode("utf-8"))

    def _compute_next_id(self):
        ids = []
        for f in self.root.findall(qn("w:footnote")):
            try:
                ids.append(int(f.get(qn("w:id"))))
            except (TypeError, ValueError):
                pass
        return (max(ids) + 1) if ids else 1

    def add_footnote(self, url):
        fid = self.next_id
        self.next_id += 1
        add_footnote_def(self.root, fid, url)
        self.inserted += 1
        return fid

    def save(self):
        self.part._blob = etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _is_simple_run(run_elem):
    """A run is 'simple' (safe to split) if it contains only text, no drawing,
    OLE object or field character."""
    for tag in ("w:drawing", "w:object", "w:fldChar", "w:instrText"):
        if run_elem.find(qn(tag)) is not None:
            return False
    return True


def process_paragraph(p_elem, para_citations, backend, replace_marker, ref_map):
    """Surgically insert footnotes for the citations belonging to paragraph
    element p_elem. Only the run(s) containing markers are touched."""
    runs = p_elem.findall(qn("w:r"))
    run_info = []
    full = ""
    for r in runs:
        txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
        run_info.append({
            "elem": r,
            "text": txt,
            "start": len(full),
            "end": len(full) + len(txt),
            "rpr": r.find(qn("w:rPr")),
        })
        full += txt

    # 1) Decide placeability + local offsets.
    for c in para_citations:
        if c.status != "pending":
            continue
        host = None
        for ri in run_info:
            if c.start >= ri["start"] and c.end <= ri["end"]:
                host = ri
                break
        if host is None:
            if p_elem.find(".//" + qn("w:hyperlink")) is not None:
                c.status = "ERROR"
                c.note = "Citation inside hyperlink; skipped to preserve the link"
            else:
                c.status = "ERROR"
                c.note = ("Citation spans runs / complex structure; left unchanged "
                          "to preserve formatting")
            continue
        if not _is_simple_run(host["elem"]):
            c.status = "ERROR"
            c.note = "Run contains drawing/field; left unchanged"
            continue
        # Idempotency: if this exact marker is already flanked by a footnote
        # reference (i.e. a previous run of this tool already processed it), skip
        # it instead of inserting a duplicate. This is mode-independent and does
        # NOT block paragraphs that merely contain unrelated existing footnotes.
        prev = host["elem"].getprevious()
        nxt = host["elem"].getnext()
        for sib in (prev, nxt):
            if (sib is not None and sib.tag == qn("w:r")
                    and sib.find(qn("w:footnoteReference")) is not None):
                c.status = "SKIPPED"
                c.note = "Previously processed (footnote already present)"
                break
        if c.status != "pending":
            continue
        c.host = host["elem"]
        c.local_start = c.start - host["start"]
        c.local_end = c.end - host["start"]

    # 2) Validate numbers and allocate footnote ids.
    for c in para_citations:
        if c.status != "pending":
            continue
        c.fids = []
        ok = True
        for num in c.numbers:
            if num not in ref_map:
                ok = False
                c.note = f"Reference {num} not found"
            elif not ref_map[num]["url"]:
                ok = False
                c.note = f"Reference {num} has no URL"
        if not ok:
            c.status = "ERROR"
            c.fids = []
            continue
        for num in c.numbers:
            c.fids.append((num, backend.add_footnote(ref_map[num]["url"])))

    # 3) Replace each affected run with text + footnote-reference runs.
    groups = defaultdict(list)
    for c in para_citations:
        if getattr(c, "host", None) is not None and c.status == "pending":
            groups[c.host].append(c)

    for run_elem, cites in groups.items():
        cites.sort(key=lambda c: c.local_start)
        new_elems = _build_run_replacement(run_elem, cites, replace_marker)
        parent = p_elem
        idx = list(parent).index(run_elem)
        parent.remove(run_elem)
        for j, el in enumerate(new_elems):
            parent.insert(idx + j, el)
        for c in cites:
            if c.fids:
                c.status = "SUCCESS"


def _build_run_replacement(run_elem, cites, replace_marker):
    text = "".join(t.text or "" for t in run_elem.findall(qn("w:t")))
    rpr = run_elem.find(qn("w:rPr"))
    new_elems = []
    pos = 0
    for c in cites:
        ls, le = c.local_start, c.local_end
        before = text[pos:ls]
        if before:
            new_elems.append(make_text_run(before, rpr))
        if c.fids:
            for (_num, fid) in c.fids:
                if fid is not None:
                    new_elems.append(make_footnote_ref_run(fid))
            if not replace_marker:            # keep the [n] text
                marker = text[ls:le]
                if marker:
                    new_elems.append(make_text_run(marker, rpr))
        else:                                 # ERROR -> keep the [n] text
            marker = text[ls:le]
            if marker:
                new_elems.append(make_text_run(marker, rpr))
        pos = le
    after = text[pos:]
    if after:
        new_elems.append(make_text_run(after, rpr))
    return new_elems


# --------------------------------------------------------------------------
# Windows Microsoft Word COM backend (EXPERIMENTAL / untested on Linux)
# --------------------------------------------------------------------------
def process_document_com(path, citations, ref_map, replace_marker, out_path):
    """Use Word's own object model to insert footnotes.

    NOTE: requires Windows + Microsoft Word + pywin32. This path is provided for
    environments where Word automation is preferred; it is NOT executed in this
    Linux sandbox and has not been verified here. The OOXML backend above is the
    validated default.

    Known limitation: handling of repeated identical markers and of ranges is
    best-effort; prefer the XML backend for full fidelity.
    """
    import os
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(path))
    try:
        from collections import defaultdict
        groups = defaultdict(list)
        for c in citations:
            if c.status == "pending":
                groups[c.marker_text].append(c)

        for marker, clist in groups.items():
            rng = doc.Range(0, doc.Content.End)
            for c in clist:
                fnd = rng.Find
                fnd.ClearFormatting()
                fnd.Text = marker
                if fnd.Execute():
                    for num in c.numbers:
                        url = ref_map.get(num, {}).get("url")
                        if not url:
                            c.status = "ERROR"
                            c.note = f"Reference {num} has no URL"
                            continue
                        fn = fnd.Range.Footnotes.Add(fnd.Range)
                        fn.Range.Text = url
                        c.status = "SUCCESS"
                        c.note = "COM footnote"
                    if replace_marker and c.status == "SUCCESS":
                        fnd.Range.Delete()
                    rng = doc.Range(fnd.Range.End, doc.Content.End)
                else:
                    c.status = "ERROR"
                    c.note = "marker not found in document"
        doc.SaveAs(os.path.abspath(out_path))
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()
