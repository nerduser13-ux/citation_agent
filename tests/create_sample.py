"""Create a small but realistic .docx to exercise the agent.

The sample intentionally includes:
  * a heading-based reference section ("References", not "5. References")
  * inline citations, multiple citations and a range ([4]-[6])
  * a citation inside a bold run (formatting preservation)
  * a real hyperlink (must be preserved, not touched)
  * a citation inside a table cell
  * one reference with NO url (-> reported as missing URL)
  * one reference that is never cited (-> reported as uncited)
"""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

OUT = Path(__file__).resolve().parent.parent / "input" / "sample.docx"


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


def main():
    doc = Document()

    doc.add_paragraph("Sample Academic Paper").runs[0].bold = True

    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph(
        "Climate research has grown rapidly [1]. Related work is also relevant "
        "[2] and [3]. See this very important [2] point restated.")

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

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT))
    print("Wrote sample document:", OUT)


if __name__ == "__main__":
    main()
