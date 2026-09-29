# Word Citation Footnote Agent

A document-preserving tool that inserts **real Microsoft Word footnotes** containing
the explicit source URL for every numbered citation marker (`[1]`, `[1], [2]`,
`[1]–[3]`, `[1-3]`) found in an existing `.docx` file.

* The **input file is never modified.**
* The reference section is detected **generically** (no hard-coded "5. References").
* Only URLs **explicitly present** in the reference list are used. Nothing is
  invented (no URLs, DOIs, authors, or titles).
* Existing Word structures are preserved: headings, bold/italic, fonts, colors,
  hyperlinks, tables, fields, drawings, lists, bookmarks, headers/footers and
  **existing footnotes**.
* Re-running on an already-processed file will **not** create duplicate footnotes.

---

## Folder structure

```
citation_agent/
├── main.py                # CLI + workflow orchestration
├── config.py              # all paths / behaviour switches / thresholds
├── document_reader.py     # read docx, enumerate body + table paragraphs, heading detection
├── reference_parser.py    # locate reference section, extract number -> URL map
├── citation_detector.py   # find [n] markers in the body, expand ranges
├── word_footnotes.py      # surgical OOXML insertion + Windows COM backend
├── validator.py           # input-unchanged + output-validity checks
├── report.py              # citation_review_report.csv writer
├── agent_tools.py         # narrow tool wrappers for a future AI agent
├── requirements.txt
├── input/                 # drop your .docx here (or use --input)
├── output/                # generated files land here
└── tests/
    ├── create_sample.py   # builds a realistic sample .docx
    └── validate_sample.py # verifies the output
```

## Requirements

```
python-docx>=1.1.0
lxml>=4.9.0
# Windows-only optional COM backend:
# pywin32>=306
```

Install (cross-platform core):

```bash
cd citation_agent
python -m pip install -r requirements.txt
```

## How to run

```bash
# 1) Use the first .docx found in ./input
python main.py

# 2) Explicit input
python main.py --input path/to/thesis.docx

# 3) Explicit input + output
python main.py --input thesis.docx --output thesis_with_footnotes.docx

# Options
--backend xml|com     # xml = OOXML (default, cross-platform); com = Windows Word COM (experimental)
--keep-marker        # keep the [n] text and insert the footnote after it (default: replace [n])
--overwrite          # overwrite the output if it already exists
--input-dir input    # default search directory
--output-dir output  # default output directory
```

Example session:

```
Word Citation Footnote Agent
Input : input/sample.docx
Output: output/sample_with_footnotes.docx
Analyzing document...
Reference section : FOUND (heading, high)
References detected: 11
URLs detected     : 10
Citation markers  : 9
Processing...

Footnotes inserted : 11
Unresolved citations: 0
Missing URLs       : 1

Output : output/sample_with_footnotes.docx
Report : output/citation_review_report.csv
```

The CSV report contains one row per citation occurrence plus a SUMMARY block:
References detected, URLs detected, Citation markers detected, Footnotes inserted,
Existing footnotes detected / preserved, Unresolved citations, Missing URLs,
Ambiguous citations, Uncited references, Previously processed, Output valid,
Reopened with python-docx, All footnote URLs from refs, Input unchanged.

## Windows / Microsoft Word COM backend

Word's own automation can be used instead of the OOXML backend:

```bash
pip install pywin32
python main.py --input thesis.docx --backend com
```

This launches Microsoft Word, inserts genuine footnotes via the Word object
model, and `SaveAs`-es to the output path (the input stays untouched). The COM
backend is **experimental and was not executed in the Linux build/test
environment**; the OOXML (`xml`) backend is the validated default and is what the
test suite exercises.

## Minimal test scenario

```bash
python tests/create_sample.py     # writes input/sample.docx
python main.py                    # produces output/sample_with_footnotes.docx + report
python tests/validate_sample.py   # checks footnotes, preservation, marker removal
```

The sample deliberately includes: a heading-based (not "5. References")
reference section, inline + multiple + range citations, a citation inside a bold
run, a real hyperlink (must survive), a citation inside a table cell, one
reference with **no URL** (reported), and one **uncited** reference (reported).

### Expected output (sample)

* 11 references detected, 10 with URLs.
* 9 citation markers -> 11 footnotes inserted (range `[4]-[6]` expands to 3;
  duplicates handled).
* 0 unresolved citations, 1 missing URL (reference 10), 2 uncited references
  (10 has no URL, 11 never cited).
* The original `input/sample.docx` is untouched (no `footnotes.xml`).
* Output reopens cleanly; the hyperlink, bold run, and table cell are preserved;
  `[n]` markers are replaced by footnote references.

## Known limitations / design choices

* **No fabrication.** If a reference has no URL, or a citation points to a missing
  reference, the tool reports it and leaves the document unchanged there.
* **Cross-run / hyperlink / field / drawing markers are not modified.** If a
  marker is split across runs, sits inside a hyperlink, or inside a run that
  contains a drawing or field, it is reported as ambiguous and skipped rather
  than risk corrupting the document.
* **Reference-section confidence.** If no reference heading/block can be
  confidently identified, the tool stops safely (input unchanged) instead of
  guessing.
* **COM backend** is experimental and untested on non-Windows; the OOXML backend
  is the validated default. Range handling and repeated-marker handling in the
  COM backend are best-effort.
* **Future Mendeley stage** is intentionally out of scope; the tool stops at
  `citation marker -> reference URL -> real Word footnote`. `agent_tools.py`
  exposes narrow, explicit functions (`inspect_document`, `find_reference_section`,
  `extract_references`, `find_citations`, `validate_citation`, `verify_output`,
  `generate_report`) so an AI agent can drive the same pipeline without
  unrestricted computer access.

## AI-agent design

Deterministic work (reading the file, extracting URLs, matching numbers,
validation, Word automation) is done in plain Python. The agent layer only calls
the explicit tools above. No LLM is used for the deterministic matching.
