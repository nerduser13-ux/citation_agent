"""word_footnotes.py

Inserts REAL Microsoft Word footnotes while preserving the existing document as
much as technically possible.

Two backends:
  * XmlFootnoteBackend  - cross-platform OOXML manipulation (DEFAULT, validated).
  * process_document_com - Windows Microsoft Word COM automation (EXPERIMENTAL;
                           requires Windows + Microsoft Word + pywin32; NOT
                           executed in the Linux build/test environment).

Surgical-editing principles (document safety first):
  * We never rebuild a whole paragraph. Only the single ``<w:r>`` run that
    actually contains a citation marker is replaced - and only when that run is
    "simple" (its children are exclusively ``w:rPr``/``w:t``). Every other
    paragraph child (bookmarks, hyperlinks, fields, drawings, proof errors,
    content controls) is left exactly where it is, in the same order.
  * Splitting happens at the text level: the marker's before/after pieces are
    new runs carrying a deep copy of the original run properties, so bold,
    italic, underline, fonts, sizes, colours and character styles survive.
  * The footnote-reference run keeps the original run's character formatting
    and merely gains the FootnoteReference character style (so Word renders the
    auto-number correctly while, e.g., bold is retained).
  * A marker that spans runs, or sits inside a hyperlink / a run containing a
    drawing / field / object / other non-text content, is NOT modified - it is
    reported AMBIGUOUS instead of risking corruption.
  * Idempotency: a marker immediately flanked by a footnote reference whose
    content equals one of the citation's own URLs is treated as previously
    processed (SKIPPED); unrelated pre-existing footnotes never cause a skip.
"""
import copy
import os
from collections import defaultdict

from lxml import etree

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part

FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")

FOOTNOTES_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:footnotes '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    "</w:footnotes>")


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------
def ensure_footnote_styles(doc):
    """Make sure the FootnoteReference (character) and FootnoteText (paragraph)
    styles exist. Reuses an existing style of the same name (any styleId);
    otherwise creates one with the standard styleId.

    Returns (reference_style_id, footnote_text_style_id)."""
    styles_root = doc.styles.element

    def find_by_name(name):
        for st in styles_root.findall(qn("w:style")):
            nm = st.find(qn("w:name"))
            if nm is not None and (nm.get(qn("w:val")) or "").lower() == name:
                sid = st.get(qn("w:styleId"))
                if sid:
                    return sid
        return None

    ref_id = find_by_name("footnote reference")
    text_id = find_by_name("footnote text")

    def add(style_id, style_type, name_val):
        st = OxmlElement("w:style")
        st.set(qn("w:type"), style_type)
        st.set(qn("w:styleId"), style_id)
        nm = OxmlElement("w:name")
        nm.set(qn("w:val"), name_val)
        st.append(nm)
        styles_root.append(st)

    if ref_id is None:
        ref_id = "FootnoteReference"
        add(ref_id, "character", "footnote reference")
    if text_id is None:
        text_id = "FootnoteText"
        add(text_id, "paragraph", "footnote text")
    return ref_id, text_id


# --------------------------------------------------------------------------
# Low-level OOXML helpers
# --------------------------------------------------------------------------
def _set_rstyle(rpr, style_id):
    """Return rPr (a copy, or a new element) with w:rStyle set to style_id.
    rStyle is always the FIRST child of rPr, per the OOXML schema."""
    if rpr is None:
        rpr = OxmlElement("w:rPr")
    existing = rpr.find(qn("w:rStyle"))
    if existing is not None:
        existing.set(qn("w:val"), style_id)
        rpr.remove(existing)
    else:
        existing = OxmlElement("w:rStyle")
        existing.set(qn("w:val"), style_id)
    rpr.insert(0, existing)
    return rpr


def make_text_run(text, rpr):
    """Plain text run; rPr is deep-copied so the original run is unaffected."""
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def make_ref_run(fid, rpr, ref_style_id):
    """A run carrying a genuine <w:footnoteReference>; keeps the source run's
    character formatting and adds the FootnoteReference character style."""
    r = OxmlElement("w:r")
    r.append(_set_rstyle(copy.deepcopy(rpr) if rpr is not None else None,
                         ref_style_id))
    fr = OxmlElement("w:footnoteReference")
    fr.set(qn("w:id"), str(fid))
    r.append(fr)
    return r


def add_footnote_def(root, fid, url, ref_style_id, text_style_id):
    """Append a genuine <w:footnote> definition containing exactly ``url``.

    The paragraph carries the FootnoteText paragraph style; the first run holds
    the automatic footnote number (w:footnoteRef) and the second run the URL
    text with no character style (it inherits the paragraph style - a
    paragraph style must never be used as a run style)."""
    ft = OxmlElement("w:footnote")
    ft.set(qn("w:id"), str(fid))
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), text_style_id)
    pPr.append(ps)
    p.append(pPr)
    # automatic footnote number (Word renders the sequence number here)
    r1 = OxmlElement("w:r")
    r1.append(_set_rstyle(None, ref_style_id))
    r1.append(OxmlElement("w:footnoteRef"))
    p.append(r1)
    # the exact URL text
    r2 = OxmlElement("w:r")
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
    def __init__(self, doc, style_ids=None):
        self.doc = doc
        self.package = doc.part.package
        self.style_ids = style_ids or ("FootnoteReference", "FootnoteText")
        self.part, self.root = self._ensure_part()
        self.next_id = self._compute_next_id(self.root)
        self.inserted = 0
        self.new_footnotes = {}   # {fid: url} footnotes added by THIS run
        self._footnote_text = self._build_text_map()

    def _ensure_part(self):
        for part in self.package.iter_parts():
            if str(part.partname).endswith("footnotes.xml"):
                return part, etree.fromstring(part.blob)
        part = Part(PackURI("/word/footnotes.xml"), FOOTNOTES_CT,
                    FOOTNOTES_TEMPLATE.encode("utf-8"), self.package)
        self.doc.part.relate_to(part, RT.FOOTNOTES)
        return part, etree.fromstring(FOOTNOTES_TEMPLATE.encode("utf-8"))

    @staticmethod
    def _compute_next_id(root):
        ids = []
        for f in root.findall(qn("w:footnote")):
            try:
                ids.append(int(f.get(qn("w:id"))))
            except (TypeError, ValueError):
                pass
        return (max(ids) + 1) if ids else 1

    def _build_text_map(self):
        m = {}
        for f in self.root.findall(qn("w:footnote")):
            fid = f.get(qn("w:id"))
            if fid in (None, "-1", "0"):
                continue
            m[fid] = "".join(t.text or "" for t in f.iter(qn("w:t")))
        return m

    def footnote_text(self, fid):
        return self._footnote_text.get(str(fid))

    def add_footnote(self, url):
        fid = self.next_id
        self.next_id += 1
        add_footnote_def(self.root, fid, url, *self.style_ids)
        self._footnote_text[str(fid)] = url
        self.new_footnotes[fid] = url
        self.inserted += 1
        return fid

    def save(self):
        self.part._blob = etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------
# Safety helpers
# --------------------------------------------------------------------------
def _run_is_simple(run_elem):
    """True iff the run contains ONLY rPr/t children - i.e. pure text.
    Any drawing, object, field character, tab, break, inline content control,
    or anything else makes the run unsafe to split."""
    return all(child.tag in (qn("w:rPr"), qn("w:t")) for child in run_elem)


def _is_ref_run(run_elem):
    return run_elem.tag == qn("w:r") and run_elem.find(qn("w:footnoteReference")) is not None


# --------------------------------------------------------------------------
# Paragraph processing (surgical)
# --------------------------------------------------------------------------
def process_paragraph(p_elem, para_citations, backend, replace_marker, ref_map):
    """Insert footnotes for the citations belonging to paragraph element
    ``p_elem``. ONLY the specific runs that contain safe citation markers are
    replaced; every other paragraph child is left untouched, in order."""
    runs = p_elem.findall(qn("w:r"))
    run_info = []
    full = ""
    for r in runs:
        txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
        run_info.append({"elem": r, "text": txt,
                         "start": len(full), "end": len(full) + len(txt),
                         "rpr": r.find(qn("w:rPr"))})
        full += txt

    # -- 1) map each pending citation to a safe host run -------------------
    for c in para_citations:
        if c.status != "pending":
            continue
        host = None
        for ri in run_info:
            if c.start >= ri["start"] and c.end <= ri["end"]:
                host = ri
                break
        if host is None:
            c.status = "AMBIGUOUS"
            c.note = ("Citation spans multiple runs / not fully inside a single "
                      "run; left unchanged to preserve formatting")
            continue
        if not _run_is_simple(host["elem"]):
            c.status = "AMBIGUOUS"
            c.note = ("Citation run contains non-text content (field/drawing/"
                      "break/...); left unchanged")
            continue
        # Idempotency: a footnote reference immediately before/after this
        # marker whose content equals one of this citation's own URLs means a
        # previous run of THIS tool already handled it. Unrelated existing
        # footnotes (different content) never trigger a skip.
        for sib in (host["elem"].getprevious(), host["elem"].getnext()):
            if sib is not None and _is_ref_run(sib):
                fid = sib.find(qn("w:footnoteReference")).get(qn("w:id"))
                text = backend.footnote_text(fid)
                if any(text == ref_map[n]["url"] for n in c.numbers
                       if n in ref_map and ref_map[n]["url"]):
                    c.status = "SKIPPED"
                    c.note = "Previously processed (matching footnote already present)"
                    break
        if c.status != "pending":
            continue
        c.host = host["elem"]
        c.local_start = c.start - host["start"]
        c.local_end = c.end - host["start"]

    # -- 2) validate numbers, allocate footnote ids ------------------------
    for c in para_citations:
        if c.status != "pending":
            continue
        c.fids = []
        ok = True
        for num in c.numbers:
            if num not in ref_map:
                ok = False
                c.note = f"Reference {num} not found in reference section"
            elif not ref_map[num]["url"]:
                ok = False
                c.note = f"Reference {num} has no URL in its entry"
        if not ok:
            c.status = "ERROR"
            continue
        c.fids = [(num, backend.add_footnote(ref_map[num]["url"]))
                  for num in c.numbers]

    # -- 3) splice replacements (grouped per host run, sorted by offset) ---
    groups = defaultdict(list)
    for c in para_citations:
        if c.host is not None and c.status == "pending" and c.fids:
            groups[c.host].append(c)

    for run_elem, cites in groups.items():
        cites.sort(key=lambda c: c.local_start)
        new_elems = _build_run_replacement(run_elem, cites, replace_marker,
                                           backend.style_ids[0])
        parent = run_elem.getparent()
        idx = list(parent).index(run_elem)
        parent.remove(run_elem)
        for j, el in enumerate(new_elems):
            parent.insert(idx + j, el)
        for c in cites:
            c.status = "SUCCESS"


def _build_run_replacement(run_elem, cites, replace_marker, ref_style_id):
    """Build the replacement element sequence for ONE host run.

    Layout:
      replace mode:  [before-text] [ref][ref]... [after-text]
      keep mode:     [before-text] [marker-text] [ref][ref]... [after-text]
    All text runs deep-copy the original run's rPr; ref runs additionally
    carry the FootnoteReference character style."""
    text = "".join(t.text or "" for t in run_elem.findall(qn("w:t")))
    rpr = run_elem.find(qn("w:rPr"))
    new_elems = []
    pos = 0
    for c in cites:
        ls, le = c.local_start, c.local_end
        before = text[pos:ls]
        if before:
            new_elems.append(make_text_run(before, rpr))
        if not replace_marker:
            marker = text[ls:le]
            if marker:
                new_elems.append(make_text_run(marker, rpr))
        for _num, fid in c.fids:
            new_elems.append(make_ref_run(fid, rpr, ref_style_id))
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

    Requires Windows + Microsoft Word + pywin32. Opens the input read-only,
    inserts footnotes for the resolvable pending citations (in document
    order, one Find per citation so repeated markers each get their own
    footnote) and SaveAs-es to ``out_path`` - the input file is never written.

    EXPERIMENTAL: not executed in the Linux build/test environment; the XML
    backend is the validated default. Citations Word cannot locate are marked
    AMBIGUOUS; missing references/URLs are marked ERROR (input untouched).
    """
    if os.name != "nt":
        raise RuntimeError(
            "--backend com requires Windows with Microsoft Word installed "
            "(pywin32). Use the default XML backend on other platforms.")
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    try:
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(path))
        try:
            last_pos = 0
            for c in sorted(
                    (c for c in citations if c.status == "pending"),
                    key=lambda c: (c.para_index, max(0, c.start))):
                rng = doc.Range(last_pos, doc.Content.End)
                fnd = rng.Find
                fnd.ClearFormatting()
                fnd.Text = c.marker_text
                fnd.Forward = True
                if not fnd.Execute():
                    c.status = "AMBIGUOUS"
                    c.note = "Marker not found by Word COM; left unchanged"
                    continue
                found = fnd.Range
                ok = True
                for num in c.numbers:
                    url = ref_map.get(num, {}).get("url")
                    if not url:
                        ok = False
                        c.note = (f"Reference {num} "
                                  + ("not found in reference section"
                                     if num not in ref_map else "has no URL in its entry"))
                        break
                    doc.Footnotes.Add(found, url)
                if ok:
                    if replace_marker:
                        found.Delete()
                    last_pos = found.End + 1
                    c.status = "SUCCESS"
                    c.note = "COM footnote inserted"
                else:
                    c.status = "ERROR"
            doc.SaveAs(os.path.abspath(out_path))
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
