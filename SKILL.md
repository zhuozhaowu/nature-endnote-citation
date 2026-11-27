---
name: nature-endnote-citation
description: Use when converting scientific manuscripts, Nature-style drafts, Markdown manuscripts, or DOCX files into Word documents that must contain real EndNote/CWYW citation fields rather than plain temporary citation text.
---

# Nature EndNote Citation

## Overview

Create Word manuscripts whose citations are true EndNote fields. Plain `{Author, Year}` text, bracketed numbers, hyperlinks, or copied bibliography text do not satisfy this skill.

Use this together with the document/DOCX skill when creating or editing `.docx` files.

## Core Contract

Every inserted citation must be a Word field with this shape in `word/document.xml`:

```xml
<w:fldChar w:fldCharType="begin"/>
<w:instrText xml:space="preserve"> ADDIN EN.CITE ...</w:instrText>
<w:fldChar w:fldCharType="separate"/>
<w:t>[visible citation]</w:t>
<w:fldChar w:fldCharType="end"/>
```

The bibliography location should contain an EndNote bibliography field:

```xml
<w:instrText xml:space="preserve"> ADDIN EN.REFLIST </w:instrText>
```

The `ADDIN EN.CITE` instruction should include `<EndNote><Cite>...<record>...</record></Cite></EndNote>` so the DOCX carries a traveling-library record. Use the visible citation style requested by the user or target journal, commonly superscript `[1]`, `[4,6]`, or `[23-28]`.

## Workflow

1. Read the manuscript source and reference source.
   - For Markdown manuscripts, identify citation tokens such as `[4]`, `[4,6]`, and `[23-28]`.
   - For EndNote exports, parse `.enw` records into numbered metadata: authors, title, journal, year, volume, issue, pages, DOI.
   - If the user provides a sample DOCX, unzip it and inspect `word/document.xml` for `ADDIN EN.CITE` and `ADDIN EN.REFLIST` before generating output.

2. Generate the Word body normally.
   - Preserve manuscript headings, captions, Methods placeholders, figure placeholders, chemical notation, and sample labels.
   - If figures are not supplied or the user wants to insert them manually, insert clear figure image placeholders before captions.

3. Replace each source citation token with an EndNote field, not text.
   - Split paragraph text around citation tokens.
   - Append field runs in order: `begin`, `instrText`, `separate`, visible citation run, `end`.
   - Put one `ADDIN EN.CITE` field per citation token. Grouped/ranged source tokens can contain multiple `<Cite>` records inside one field while preserving the original visible token.

4. Insert the bibliography field.
   - Add a `References` heading if the manuscript has one.
   - Under it, insert a single `ADDIN EN.REFLIST` field. Do not paste a manual numbered bibliography as final text unless the user explicitly asks for a static bibliography.

5. Verify by inspecting DOCX internals.
   - Count `ADDIN EN.CITE`.
   - Count `ADDIN EN.REFLIST`.
   - Confirm `<record>` payloads are present.
   - Confirm the manual reference list was not duplicated.
   - If LibreOffice/Word rendering is available, also render or open the file to check visible layout.

## Minimal Python Pattern

Use `python-docx` plus OOXML runs. The important part is creating actual field runs:

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def append_field_run(paragraph, fld_type):
    run = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), fld_type)
    run._r.append(fld)

def append_instr_run(paragraph, instruction):
    run = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    run._r.append(instr)

def add_endnote_cite(paragraph, endnote_xml, visible_text):
    append_field_run(paragraph, "begin")
    append_instr_run(paragraph, f" ADDIN EN.CITE {endnote_xml}")
    append_field_run(paragraph, "separate")
    paragraph.add_run(visible_text)
    append_field_run(paragraph, "end")

def add_endnote_reflist(paragraph):
    append_field_run(paragraph, "begin")
    append_instr_run(paragraph, " ADDIN EN.REFLIST ")
    append_field_run(paragraph, "separate")
    append_field_run(paragraph, "end")
```

The `endnote_xml` argument should be a valid escaped EndNote payload, for example:

```xml
<EndNote>
  <Cite>
    <Author>Jin</Author>
    <Year>2021</Year>
    <RecNum>9004</RecNum>
    <DisplayText><style face="superscript">[4,6]</style></DisplayText>
    <record>...</record>
  </Cite>
</EndNote>
```

## Validation Script

Use `scripts/validate_endnote_fields.py` on the final DOCX. Example:

```bash
python scripts/validate_endnote_fields.py manuscript.docx --min-cites 1 --require-reflist
```

Passing this script is structural evidence that the DOCX contains EndNote fields. It does not prove that the user's local EndNote library has already formatted the bibliography; Word with EndNote installed must update the fields for final bibliography formatting.

## Common Mistakes

| Mistake | Correct action |
|---|---|
| Writing `{Author, Year}` text | Create `ADDIN EN.CITE` Word fields instead |
| Keeping only visible `[1]` text | Add field begin/instr/separate/end around the visible citation |
| Pasting a static reference list | Insert `ADDIN EN.REFLIST` and let EndNote update it |
| Assuming `.enw` import equals DOCX fields | Embed citation fields in the DOCX and copy `.enw` only as support |
| Verifying with `python-docx` paragraph text only | Inspect zipped `word/document.xml` for `ADDIN EN.CITE` |

## Completion Criteria

Report the final DOCX path and these counts:

- `ADDIN EN.CITE`
- `ADDIN EN.REFLIST`
- embedded `<record>` payloads
- visible citation tokens
- figure placeholders, if relevant

If visual DOCX rendering is unavailable, say that structural EndNote-field validation passed but visual render QA could not be completed.
