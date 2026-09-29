"""Create a small but realistic .docx to exercise the agent.

The sample intentionally includes:
  * a heading-based reference section ("References", not "5. References")
  * inline citations, multiple citations, repeated numbers and a range ([4]-[6])
  * a citation inside a bold run (formatting preservation)
  * a real hyperlink (must be preserved, not touched)
  * a citation inside a table cell
  * a PRE-EXISTING manual footnote in the introduction (must be preserved)
  * one reference with NO url (-> reported as missing URL)
  * one reference that is never cited (-> reported as uncited)

Usage:
    python tests/create_sample.py              # writes input/sample.docx
    python tests/create_sample.py OUT.docx     # writes OUT.docx instead
"""
import sys
from pathlib import Path

from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part

OUT = Path(__file__).resolve().parent.parent / "input" / "sample.docx"

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


def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_preexisting_footnote(doc, paragraph, text):
    """Add a genuine manual footnote (id=1) to ``paragraph`` so the sample
    already contains footnote infrastructure the tool must preserve."""
    part = Part(PackURI("/word/footnotes.xml"), FOOTNOTES_CT,
                FOOTNOTES_TEMPLATE.encode("utf-8"), doc.part.package)
    doc.part.relate_to(part, RT.FOOTNOTES)
    root = etree.fromstring(FOOTNOTES_TEMPLATE.encode("utf-8"))
    ft = OxmlElement("w:footnote")
    ft.set(qn("w:id"), "1")
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:footnoteRef"))
    p.append(r)
    r2 = OxmlElement("w:r")
    t2 = OxmlElement("w:t")
    t2.text = text
    r2.append(t2)
    p.append(r2)
    ft.append(p)
    root.append(ft)
    part._blob = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # body reference run
    run = OxmlElement("w:r")
    fr = OxmlElement("w:footnoteReference")
    fr.set(qn("w:id"), "1")
    run.append(fr)
    paragraph._p.append(run)


def main(out=OUT):
    out = Path(out)
    doc = Document()

    doc.add_paragraph("Sample Academic Paper").runs[0].bold = True

    doc.add_heading("1. Introduction", level=1)
    p = doc.add_paragraph(
        "Climate research has grown rapidly [1]. Related work is also relevant "
        "[2] and [3]. See this very important [2] point restated.")
    add_preexisting_footnote(
        doc, p, "This footnote already existed in the source document.")

    doc.add_heading("2. Background", level=1)
    p = doc.add_paragraph()
    p.add_run("Measurement studies show clear effects [4]-[6]. ")
    r = p.add_run("A bold claim with a citation [7] inside it.")
    r.bold = True
    p.add_run(" Additional note [8].")
    # a hyperlink that must survive untouched
    add_hyperlink(p, "https://example.com/external", " (external resource)")
    p.add_run(" More text after the link [1].")

    doc.add_heading("3. Methods", level=1)
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    cell.text = "The method builds on prior art [9]."

    # ---- Reference section (heading-based; NOT "5. References") -------------
    doc.add_heading("References", level=1)
    refs = [
        "1. Author A. Title one. Journal. 2020. https://example.com/one",
        "2. Author B. Title two. Journal. 2021. https://example.com/two",
        "3. Author C. Title three. https://doi.org/10.1000/three",
        "4. Author D. Title four. https://example.com/four",
        "5. Author E. Title five. https://example.com/five",
        "6. Author F. Title six. https://example.com/six",
        "7. Author G. Title seven. https://example.com/seven",
        "8. Author H. Title eight. https://example.com/eight",
        "9. Author I. Title nine. https://example.com/nine",
        "10. Author J. Title ten with no URL. Journal. 2022.",   # NO URL
        "11. Author K. Title eleven uncited. https://example.com/eleven",  # uncited
    ]
    for rtext in refs:
        doc.add_paragraph(rtext)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print("Wrote sample document:", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else OUT)
