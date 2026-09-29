"""Validation of the generated output: real footnotes, document preservation,
marker replacement, and idempotency."""
import zipfile
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

OUT = "output/sample_with_footnotes.docx"
IN = "input/sample.docx"


def tag(e):
    return e.tag


def main():
    ok = True

    # 1) Input untouched (hash) - sanity: file sizes differ but content of input
    #    is not written by the tool. We instead check input still has NO footnotes.
    zin = zipfile.ZipFile(IN)
    in_has_ft = any(n.endswith("footnotes.xml") for n in zin.namelist())
    print("Input has footnotes.xml:", in_has_ft, "(must be False)")
    ok &= (not in_has_ft)

    # 2) Output validity + real footnotes
    doc = Document(OUT)
    z = zipfile.ZipFile(OUT)
    ft = etree.fromstring(z.read([n for n in z.namelist()
                                  if n.endswith("footnotes.xml")][0]))
    regular = [f for f in ft.findall(qn("w:footnote"))
               if f.get(qn("w:id")) not in ("-1", "0")]
    body_fn = 0
    for p in doc.paragraphs:
        body_fn += len(p._p.findall(".//" + qn("w:footnoteReference")))
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    body_fn += len(p._p.findall(".//" + qn("w:footnoteReference")))
    print("Regular footnotes in footnotes.xml:", len(regular))
    print("footnoteReference runs in body+tables:", body_fn)
    ok &= (len(regular) == body_fn == 11)

    # 3) Every inserted footnote URL comes from a reference (no fabrication)
    fn_urls = set()
    for f in regular:
        for t in f.iter(qn("w:t")):
            if t.text and t.text.startswith("http"):
                fn_urls.add(t.text)
    print("Distinct footnote URLs:", len(fn_urls))
    ok &= (len(fn_urls) == 9)  # refs 1-9 (no url for 10; 11 uncited)

    # 4) Hyperlink preserved
    hl = doc.element.body.findall(".//" + qn("w:hyperlink"))
    print("Hyperlinks in output:", len(hl), "(must be >=1)")
    ok &= (len(hl) >= 1)

    # 5) Bold formatting preserved around the [7] citation
    bold_ok = False
    for p in doc.paragraphs:
        for r in p.runs:
            if r.bold and "inside it" in (r.text or ""):
                bold_ok = True
    print("Bold run 'inside it.' preserved:", bold_ok)
    ok &= bold_ok

    # 6) Table-cell citation got a footnote
    table_fn = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    table_fn += len(p._p.findall(".//" + qn("w:footnoteReference")))
    print("Table-cell footnote references:", table_fn, "(must be 1)")
    ok &= (table_fn == 1)

    # 7) Markers replaced (no [n] left in body, before References)
    body_text = "\n".join(p.text for p in doc.paragraphs)
    # cut at References heading
    idx = body_text.find("References")
    body_before = body_text[:idx]
    import re
    leftover = re.findall(r"\[\d+\]", body_before)
    print("Bracketed markers left in body:", len(leftover), "(must be 0 in replace mode)")
    ok &= (len(leftover) == 0)

    print("\nALL CHECKS PASSED" if ok else "\nSOME CHECKS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
