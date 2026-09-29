"""validator.py

Pre-save / post-save validation. The tool never fabricates data, and it must
prove the input was untouched and the output is a valid, reopenable .docx whose
footnote URLs all come from the reference list.
"""
import hashlib
import zipfile
from lxml import etree

from docx import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def original_unchanged(path, original_hash):
    return file_hash(path) == original_hash


def validate_output(out_path, ref_map, citations, existing_footnote_count):
    res = {
        "reopen_ok": False,
        "valid_docx": False,
        "footnote_count": 0,
        "all_footnote_urls_in_refs": True,
        "existing_preserved": False,
        "error": "",
    }
    try:
        Document(out_path)          # reopen with python-docx
        res["reopen_ok"] = True
    except Exception as e:           # noqa: BLE001
        res["error"] = f"python-docx reopen failed: {e}"
        return res

    z = zipfile.ZipFile(out_path)
    ft = None
    for n in z.namelist():
        if n.endswith("footnotes.xml"):
            ft = etree.fromstring(z.read(n))
    if ft is None:
        res["valid_docx"] = True
        return res

    regular = [f for f in ft.findall(qn("w:footnote"))
               if f.get(qn("w:id")) not in ("-1", "0")]
    res["footnote_count"] = len(regular)
    res["valid_docx"] = True

    fn_urls = set()
    for f in regular:
        for t in f.iter(qn("w:t")):
            if t.text and t.text.startswith("http"):
                fn_urls.add(t.text)
    ref_urls = {v["url"] for v in ref_map.values() if v["url"]}
    res["all_footnote_urls_in_refs"] = fn_urls.issubset(ref_urls)
    res["existing_preserved"] = res["footnote_count"] >= existing_footnote_count
    return res


def qn(tag):
    return "{" + W + "}" + tag
