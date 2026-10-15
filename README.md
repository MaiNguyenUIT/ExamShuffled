# Exam Shuffler

**[Hướng dẫn sử dụng bằng tiếng Việt](README.vi.md)** · The application interface, warnings, validation results and exported answer-key text now use Vietnamese. Original exam content is preserved.

A local Streamlit application for randomizing Vietnamese Word exams. Python 3.11+ on Windows or Linux; no Microsoft Word, LibreOffice, database, Docker or external API is required. Documents stay in the local Python process. Installing dependencies requires internet access; processing does not.

## Install and start

Windows PowerShell, from this project directory:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Python 3.11 or a newer supported version can replace `-3.12`. Activation is optional. To activate:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
python -m streamlit run app.py
```

Ubuntu:

```bash
sudo apt update
sudo apt install python3 python3-venv python3-pip
# Check python3 --version: Python 3.11+ is required (Ubuntu 24.04 includes 3.12).
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Open **http://127.0.0.1:8501**. Stop with Ctrl+C. The default config binds to localhost and disables Streamlit usage statistics. Run commands from this directory so `.streamlit/config.toml` is applied.

## Workflow

1. Upload a `.docx` exam (up to 30 MiB).
2. Review sections, questions, options, table contents, scores and warnings. Select the review checkbox. Blocking errors require correcting the source; warnings require review.
3. Set version count (default 4), first code (101), integer seed (42) and shuffle settings. Questions stay in their section. Essay movement requires both question shuffling and essay shuffling. MC options default to shuffled; true/false statements default to fixed.
4. Optionally enter answers using **original** question numbers and labels. For a true/false question, fill all statements or leave the entire question blank. Partial keys across questions are supported. Colored-text hints are never verified automatically; review and confirm your entered key.
5. Generate and download individual DOCX files, the optional `Answer_Key.docx`, or the ZIP. Every exam is validated and reopened before downloads are published. Changing settings or answers hides stale downloads until you generate again.

An exam-code paragraph is inserted above the first section, below school information. Source files are never overwritten. Generation uses in-memory copies.

## Included reference run

The supplied `docxTest/SỬ-11_CK2.docx` was parsed as 12 multiple-choice, 4 true/false and 2 essay questions. Four sample exams are saved in `output/reference/`, alongside `Exam_Shuffler.zip` and `Validation_Report.json`. This sample run enables all question/choice shuffles with seed 42. The four question orders differ; all files reopen and pass structural/package/content validation. The original upload is unchanged. No sample answer key is included because no verified answers were supplied. Visual pagination has not been rendered or compared.

## Input conventions

- Body-paragraph anchors: `Câu 1:`, `Câu 1.`, `Câu 1. (2,0 điểm).`, including labels split across styled runs.
- Section headings begin with `PHẦN I.`, `PHẦN 1.`, `PART I`, and similar Roman/numeric headings. Types are inferred from heading wording and options. Section/question counts are not fixed.
- Four explicit MC options `A.`–`D.` (also `A)`–`D)`), or four true/false statements `a)`–`d)`, in original label order. Choices can occupy separate body paragraphs with continuation paragraphs, or all four can share one simple paragraph separated by whitespace/tabs.
- Passages and tables between a question anchor and its first option belong to that question. Whole-question movement retains its complete ordered paragraph/table group. Parenthesized essay scores travel with the question OOXML.
- `HẾT`, `END`, and dashed variants mark the body ending. Without a marker, trailing text is conservatively attached to the last question and a warning appears.
- Missing section headings create an implicit section with a warning. Add headings if different groups must remain separate.

## Preservation and validation

The exporter deep-copies complete OOXML question blocks and edits only question-number/choice-label spans across existing runs. It never rebuilds exam paragraphs or tables from plain text. Simple inline choices retain run styles, paragraph properties and whitespace separators; runs may be split at choice boundaries.

Only `word/document.xml` is rewritten in the original ZIP package. All other parts, including styles, numbering, images, embedded objects, charts, headers, footers, relationships, fonts and custom XML, retain their original payload bytes. Content stays in the same document part, so existing relationship IDs remain valid. No relationships are copied between unrelated documents. The answer key is a new python-docx document.

Validation checks question identity/multiplicity, section membership, numbers, option permutations, round-trip answer mapping, reopened document structure, question passages/tables/styles, scores, section properties, unchanged package parts and document relationship references. Source choices are compared with their mapped destinations. ZIP timestamps are fixed for reproducible archives; source DOCX entry metadata is retained.

Question orders are sampled without replacement from the enabled within-section permutation space, without enumerating it or using unbounded retries. Repetition occurs only after that space is exhausted. Option orders are reproducible but not guaranteed unique across versions. Disabling question shuffling or limiting movable questions can make unique question orders impossible; the report states this.

## Known DOCX limitations

Structural checks do **not** prove identical visual rendering. No Word rendering engine is used. Shuffling and the exam-code paragraph can change page breaks, pagination, widow/orphan placement and table layout. Fonts must be installed on the viewing computer. Inspect downloaded documents in a compatible viewer before printing.

- Questions inside tables/text boxes, tracked revisions, alternate-content imports, digitally signed files, automatic numbered question content and section breaks inside question groups are blocked with explanations.
- Four-option/four-statement questions are supported. Other counts, duplicate/missing labels and table-based choices are blocking errors rather than silently misparsed questions.
- Mixed multi-paragraph inline grids, complex inline fields/hyperlinks/drawings, or tables after the first answer label must keep option order. Disable the relevant shuffle. Whole questions still retain these objects when moved.
- Bookmarks/comments crossing question boundaries are blocked. Fields and cross-references are retained but not recalculated. PAGE fields in headers/footers survive; cached values may need refresh. Semantic dependencies and phrases like “all of the above” require manual review before shuffling.
- Reading passages shared across separately numbered questions are not a supported group structure. Place the dependent group under one anchor, or disable question shuffling after reviewing the parse.
- Trailing/blank paragraphs remain attached to their question. Unmarked body instructions cannot reliably be separated from essay response areas; use an ending marker.
- Color detection inspects direct run properties. Inherited colors, highlights and underlines are not reliable answer evidence. Inline paragraph colors can flag several choices. Colors are preserved in exported exams; remove answer-revealing formatting from student templates yourself.
- Existing static exam codes in headers/footers or school tables are retained. Remove old codes from templates before uploading to avoid a conflicting static code beside the generated code.
- Intended for trusted local documents, not a hardened untrusted-document service. Package size limits and disabled external XML entities reduce malformed-input risks. Batches are limited to 100 versions and an estimated 256 MiB input-size × version-count budget.

## Tests

Windows:

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

Ubuntu:

```bash
.venv/bin/python -m pytest -q
```

Tests cover parsing, split-run labels, seeded permutations, permutation exhaustion, disabled shuffles, MC/TF answer mapping, confirmation, inline styles/spacing, tables/passages/scores, package parts, media, hyperlinks, equations, headers/footers, ZIP exports and corruption rejection. Streamlit tests exercise UI rendering and generated-download state.

Place `SỬ-11_CK2.docx` in `docxTest/` for reference integration tests. These verify **12 MC, 4 true/false, 2 essay**, the Part II Question 2 table, four different question orders, scores, successful DOCX reopening and source immutability. Reference tests skip if the file is absent. Test answer data is synthetic, not a verified History answer key.

## Architecture

```text
app.py                      Streamlit workflow and session state
core/models.py              Typed data classes
core/question_detector.py   Vietnamese anchors and labels
core/document_blocks.py     Safe package and OOXML operations
core/parser.py              Sections, questions, choices and warnings
core/shuffler.py             Reproducible generation plans
core/answer_key.py           Confirmation and answer remapping
core/exporter.py             DOCX, answer keys and ZIP bundles
core/validator.py            Plan/package/content validation
tests/                      Unit and integration tests
```

Development phases: parser/preview tests, whole-question export, options/answer mapping, then full validation/batch export/UI coverage. Dependencies have bounded compatible ranges; pin your validated environment for controlled deployments.
