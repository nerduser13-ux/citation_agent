# Word Citation Footnote Agent

A document-preserving tool that inserts **real Microsoft Word footnotes**
containing the explicit source URL for every numbered citation marker
(`[1]`, `[1], [2]`, `[1]–[3]`, `[4]-[6]`, …) found in an existing `.docx`
file. The References/Bibliography section of the document is used as the only
authoritative number → source-URL mapping.

* The **input file is never modified** — its SHA-256 hash is verified before
  and after processing, and the result is written to a new output path.
* The reference section is detected **generically** (no hard-coded
  "5. References"); if it cannot be identified with confidence the tool
  **aborts without producing output**.
* Only URLs **explicitly present** in a reference entry are used. Nothing is
  invented — no URLs, DOIs, authors, titles, publications, or numbers.
  Bare "DOI: 10.1000/abc" text is **not** turned into a URL.
* Editing is **surgical**: only the exact `<w:r>` run containing a safe
  citation marker is replaced. Paragraph/character formatting, bold, italic,
  underline, fonts, sizes, colors, styles, headings, lists, tables,
  hyperlinks, bookmarks, fields, drawings, images, headers, footers, page
  layout, section breaks, content controls, existing footnotes and all other
  unrelated OOXML are preserved.
* Markers that cannot be safely modified (split across runs, inside a
  hyperlink, inside a field/drawing/complex run, reversed ranges) are
  reported as `AMBIGUOUS` and left **completely untouched**.
* Re-running on an already-processed file does **not** create duplicate
  footnotes (idempotency guard), while unrelated pre-existing footnotes never
  block new citations and are preserved verbatim.

---

## Folder structure

```
citation_agent/
├── main.py                # CLI + workflow orchestration
├── config.py              # all paths / behaviour switches / detection thresholds
├── document_reader.py     # read-only docx access; single shared paragraph ordering
│                          #   (body + tables incl. nested + content controls),
│                          #   heading detection, raw-text helpers, footnote snapshot
├── reference_parser.py    # reference-section location + number -> URL extraction
├── citation_detector.py   # [n] / [a]-[b] detection in the body only, range expansion
├── word_footnotes.py      # surgical OOXML footnote insertion + experimental COM backend
├── validator.py           # input-unchanged + full post-save output validation
├── report.py              # citation_review_report.csv writer
├── agent_tools.py         # narrow tool wrappers for a future AI agent
├── requirements.txt
├── input/                 # drop your .docx here (or use --input)
├── output/                # generated files land here
└── tests/
    ├── create_sample.py   # builds a realistic sample .docx (incl. a manual footnote)
    ├── validate_sample.py # deep checks of the sample output (OOXML level)
    └── test_pipeline.py   # pytest end-to-end suite (39 tests)
```

## Requirements

Python 3.9+ and:

```
python-docx>=1.1.0
lxml>=4.9.0
# optional, Windows-only COM backend:
# pywin32>=306
# optional, for the test suite:
# pytest
```

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/pip install pytest
```

## How to run

```bash
# 1) Use the first .docx found in ./input
python main.py

# 2) Explicit input
python main.py --input path/to/thesis.docx

# 3) Explicit input + output
python main.py --input thesis.docx --output thesis_with_footnotes.docx
```

Options:

| Flag | Meaning |
| --- | --- |
| `--backend xml\|com` | `xml` = OOXML manipulation (default, cross-platform, validated). `com` = Windows Microsoft Word COM automation (experimental, Windows-only). |
| `--keep-marker` | Keep the `[n]` text and insert the footnote reference **after** it. Default: replace `[n]` with the footnote reference. |
| `--overwrite` | Allow overwriting an existing output file (default: create `<name>_with_footnotes_1.docx`, `_2`, …). |
| `--input-dir` | Directory scanned for the default input (default `input`). |
| `--output-dir` | Output directory (default `output`). |

Exit codes: `0` success (output produced **and validated**); `1` aborted
(reference section not found) or validation failed; `2` usage error /
unsupported backend; `3` input file changed during processing (fatal).

Example session (sample document):

```
Word Citation Footnote Agent
Input : input/sample.docx
Output: output/sample_with_footnotes.docx
Analyzing document...
Reference section : FOUND (heading, high) - Reference heading detected: "References"
References detected: 11
URLs detected     : 10
Citation markers  : 9
Processing...

Footnotes inserted : 11
Unresolved citations: 0
Missing URLs       : 1
Ambiguous citations: 0
Validation         : PASS
```

## Citation syntax supported

* `[1]` — single citation
* `[1] and [2]`, `[1], [2], and [3]` — multiple distinct citations
* `[2] … [2]` — repeated citation numbers (each occurrence gets a footnote)
* `[1]-[3]`, `[1]–[3]` (en dash), `[1]—[3]` (em dash) — ranges, expanded
  (`[4]-[6]` → references 4, 5, 6 → three footnotes)

Detection operates on the actual OOXML structure:

* Only the **direct runs** of a paragraph are used as the coordinate space,
  so a marker next to a hyperlink is never mis-attributed to the link.
* Detection **stops at the start of the reference section** — the reference
  list itself is never scanned for body citations.
* Arbitrary numbers in prose are never citations (brackets are required).

## Reference-section detection

The heading is matched generically, most-conservative-first:

1. **Heading (high confidence)** — a styled heading (Heading 1–9 / Title /
   outline level, or an all-bold short line) whose text, after stripping
   leading numbering (`5.`, `1.2.`, `A.`) and trailing punctuation, is
   *exactly* a keyword: *references, reference, bibliography, works cited,
   reference list, literature cited, sources*. The **last** such heading wins
   (reference sections sit at the end of the document).
2. **Heading + block (medium)** — a heading that merely *contains* a keyword
   is accepted only if it is directly followed by a qualifying numbered block
   (≥ 3 numbered entries, ≥ 1 URL in entry lines).
3. **Block detection (medium)** — with no heading at all, a numbered block
   near the end of the document (starting at/after 40% of the paragraph list)
   with ≥ 3 numbered entries and ≥ 1 URL.
4. **Nothing** → the tool **aborts**: no output file, input unchanged, and
   the CSV report says `ABORTED - reference section could not be
   confidently identified`. It never guesses.

A short body sentence that merely contains a keyword (e.g. "Our data sources
were collected …") is **not** treated as a heading.

### Reference parsing rules

* Entry lines look like `1. …`, `1) …`, `1 …` or `[1] …`.
* A line starting with a year (`2020. …`) is **not** a numbered entry; it ends
  the current entry (its URLs must not be attributed to a reference number).
* Wrapped continuation lines are appended to the current entry.
* Duplicate numbers: the first occurrence wins (warning recorded).
* Only explicit `http(s)://` URLs are extracted. Trailing sentence
  punctuation (`. , ; : ! ?`) is stripped from the end of a URL.
* **Multiple URLs in one entry** — deterministic selection rule: the **last**
  URL that is not a DOI resolver (`doi.org`); if every URL is a DOI resolver,
  the **last** URL. URLs are used exactly as written.

## Safety rules (what is never modified)

A citation marker is modified only when **all** of the following hold:

1. it lies entirely inside a **single direct run** of the paragraph,
2. that run contains **only text** (its children are exclusively
   `w:rPr`/`w:t` — no fields, drawings, OLE objects, tabs, breaks, or inline
   content controls),
3. its number exists in the reference section **and** that reference has an
   explicit URL.

Otherwise the marker is left byte-for-byte untouched and reported:

| Status | Meaning |
| --- | --- |
| `SUCCESS` | footnote(s) inserted |
| `ERROR` | reference number missing, or reference has no URL (unresolved) |
| `AMBIGUOUS` | split across runs / inside hyperlink / complex run / reversed range |
| `SKIPPED` | already processed by a previous run of this tool |

Specifically, markers inside `<w:hyperlink>`, fields, or runs containing
drawings/objects/breaks are **not** modified; unrelated hyperlinks, fields,
bookmarks and drawings are always left exactly where they are.

## Real Word footnotes (XML backend)

The output contains genuine OOXML footnotes:

* `word/footnotes.xml` with real `<w:footnote>` definitions (including the
  required `separator` / `continuationSeparator` entries) and a
  `footnotes.xml` relationship on the main document part,
* corresponding `<w:footnoteReference>` elements in the body,
* the FootnoteText paragraph style and FootnoteReference character style
  (reused if the document already defines them, created otherwise),
* the displayed numbering is Word's own footnote mechanism (automatic
  `<w:footnoteRef>`), never typed text.

Each generated footnote contains **exactly** the URL from the corresponding
reference entry, unaltered. For a range like `[4]-[6]` three footnotes are
created (references 4, 5, 6).

The insertion is surgical at the XML level: only the affected run is replaced
by `[before-text] [footnote-ref…] [after-text]` (or
`[before-text] [marker-text] [footnote-ref…] [after-text]` with
`--keep-marker`). The reference runs deep-copy the original run's character
formatting (so e.g. bold is kept) and gain the FootnoteReference character
style, inserted as the schema-correct first child of `w:rPr`.

## Idempotency

* **Replace mode** (default): markers disappear after the first run, so a
  second run finds nothing to do and preserves everything.
* **Keep-marker mode**: on the second run the marker is still present; it is
  recognized as previously processed when a footnote reference immediately
  adjacent to it carries **one of the citation's own URLs** — and is skipped
  without creating a duplicate. A pre-existing footnote with different
  content never triggers a skip.

## Input protection & output naming

* SHA-256 of the input is computed before processing and re-verified after;
  any change aborts with a loud failure (exit 3).
* Processing happens on a separate in-memory `Document` instance; the only
  file written is the output.
* Default output: `output/<input-name>_with_footnotes.docx`. If that file
  exists, `<input-name>_with_footnotes_1.docx`, `_2`, … is used unless
  `--overwrite` is given. An output path equal to the input path is rejected.

## Validation

After saving, the tool automatically verifies (and reports `PASS`/`FAIL`):

* output exists, is a valid ZIP, reopens with python-docx;
* `word/footnotes.xml` exists and the **count** of regular footnotes and body
  footnote references is exactly right (existing + newly inserted);
* every **new** footnote contains exactly the URL of its reference (no
  fabricated URLs);
* **existing footnotes are preserved** — every pre-existing definition still
  exists with identical text, and every pre-existing body reference is intact;
* **per-paragraph raw-text fidelity** — the output text equals the input text
  minus exactly the removed markers (this also proves no text was lost from
  tables, hyperlinks, or any other structure);
* structural preservation — paragraph, table, hyperlink, bookmark, field,
  drawing, content-control, page-size and section-break counts unchanged, and
  every input hyperlink target is still present;
* the input file's hash is unchanged.

If any check fails the run exits non-zero and the failures are printed.

## CSV review report

`citation_review_report.csv` (next to the output) contains one row per
citation occurrence:

`Citation Number, Location, Citation Text, Reference Text, URL, Action,
Status, Notes`

followed by a `SUMMARY` block: Input, Output, Reference section
(+ method + confidence), References detected, URLs detected, Citation markers
detected, Footnotes inserted, Existing footnotes detected / preserved,
Unresolved citations, Missing URLs, Ambiguous citations, Uncited references,
Previously processed, Output valid, Reopened with python-docx, All footnote
URLs from refs, Input unchanged, Status.

## Windows / Microsoft Word COM backend

```bash
pip install pywin32
python main.py --input thesis.docx --backend com
```

Opens the document in Microsoft Word, inserts genuine footnotes via the Word
object model (one Find per citation, in document order), deletes the marker
text when replacing, and `SaveAs`-es to the output path (input untouched).
The **experimental** backend is Windows-only and was **not executed** in the
Linux build/test environment (it fails loudly on other platforms); the XML
backend is the validated default and is what the test suite exercises.

## Tests

```bash
python -m pytest tests/ -v                 # 39 end-to-end tests (OOXML-level)
python tests/create_sample.py              # (re)generate input/sample.docx
python main.py                             # process the sample
python tests/validate_sample.py            # deep checks of the sample output
```

The pytest suite covers: simple `[1]`; multiple citations; repeated numbers;
`[1]-[3]`, `[1]–[3]` and `[1]—[3]` ranges; reversed range (ambiguous);
citation in bold text; citation in a table (incl. nested tables); hyperlink
preservation; citation **inside** a hyperlink (untouched); citation **split
across runs** (untouched); fields/bookmarks preservation; missing reference
URL; missing reference number; uncited reference; reference section not
scanned for citations; year-starting line not an entry; DOI-preference URL
selection; bare DOI text never invented; pre-existing footnotes preserved;
unrelated existing footnote not blocking; idempotent second run (both
modes); input hash unchanged; non-standard headings ("Bibliography",
"5. References", bold unstyled "Works Cited"); block fallback without
heading; keyword-sentence false-positive avoidance; abort when no reference
section; keep-marker mode; output-collision safe naming; `--overwrite`;
output==input rejection; COM backend rejection on Linux; sample end-to-end.

The sample document deliberately includes: a heading-based reference section,
inline/multiple/repeated/range citations, a citation inside a bold run, a
real hyperlink, a citation inside a table cell, a **pre-existing manual
footnote**, one reference with **no URL** (reported unresolved), and one
**uncited** reference (reported).

## Known limitations / design choices

* **Citations inside inline content controls, text boxes, headers, footers
  or footnotes themselves** are out of scope: they are never modified (their
  surrounding structure is preserved) but not reported either, because they
  are not part of the body-paragraph coordinate space.
* **Keep-marker idempotency** is position-sensitive: it matches a footnote
  reference *immediately* adjacent to the marker with matching content. If a
  user inserts extra text between marker and footnote after processing, a
  re-run may add another footnote (conservative, documented behavior).
* **`[n]` inside a reference entry that has no URL** and other
  unresolvable markers are left in place by design (never invented).
* **COM backend** is experimental, Windows-only, and untested here; prefer
  the XML backend.
* The tool operates on the **first document part only** (`word/document.xml`);
  sub-documents (e.g. attached OLE Word objects) are preserved but not
  processed.
* **Future Mendeley stage** is intentionally out of scope; the pipeline stops
  at `citation marker → reference URL → real Word footnote`. `agent_tools.py`
  exposes narrow, explicit functions (`inspect_document`,
  `find_reference_section`, `extract_references`, `find_citations`,
  `validate_citation`, `verify_output`, `generate_report`, `run_pipeline`)
  so an AI agent can drive the same deterministic pipeline without
  unrestricted filesystem or shell access.

## AI-agent design

All deterministic work (reading the file, extracting URLs, matching numbers,
insertion, validation) is plain Python and fully reproducible. The agent
layer (`agent_tools.py`) only calls the explicit tools above and can never
manipulate arbitrary files.
