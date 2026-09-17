# EngGeo Manuscript Format

A reusable Codex and Claude Code skill for creating, revising, and auditing English-language engineering-geology and SCI manuscript DOCX files. The repository follows the latest `liuhuaxin-mines/SCI_manuscript_format` rules as its baseline and adds the owner's reviewed amendments for line numbering, citation preservation, reference order, figure-caption alignment, heading typography, and body indentation.

This is a general draft convention. Current target-journal instructions override a conflicting rule only when they are supplied or verified.

## Included files

- `SKILL.md` — skill entry point and workflow
- `references/format-spec.md` — authoritative reviewed specification
- `assets/enggeo_manuscript_template.docx` — reusable manuscript template
- `scripts/build_template.py` — deterministic template generator
- `scripts/audit_template.py` — structural template audit
- `agents/openai.yaml` — Codex display metadata

## Current core conventions

- A4 portrait with 2.54 cm margins and a centered dynamic footer page number
- continuous Word line numbering in every section; `Line Number` uses Times New Roman 11 pt
- `Normal` body text: Times New Roman 12 pt, justified, 1.5 spacing, 6 pt before and after, 0.85 cm first-line indent
- zero-indent exceptions for title-page text, abstract prose, Keywords, captions, equations, tables, and references
- Times New Roman for the manuscript title and every heading, with Word theme-font overrides removed
- exact English manuscript-role styles: `Normal`, `Figure`, `Figure Caption`, `Table Caption`, `Equation`, `Table`, `Reference`, and `Heading 1`–`Heading 3`
- 12 pt italic justified figure captions below figures
- 12 pt italic left-aligned table captions above three-line tables
- native editable Word OMML equations
- preservation of the source author-year or numeric citation system, source citation typography, and exact reference order

## Regenerate and audit the template

Install the packages listed in `requirements.txt`, then run:

```powershell
python scripts/build_template.py
python scripts/audit_template.py assets/enggeo_manuscript_template.docx
```

## Install for Codex

```powershell
git clone https://github.com/gushancongan-arch/enggeo-manuscript-format.git "$HOME\.agents\skills\enggeo-manuscript-format"
```

Invoke with `$enggeo-manuscript-format`.

## Install for Claude Code

```powershell
git clone https://github.com/gushancongan-arch/enggeo-manuscript-format.git "$HOME\.claude\skills\enggeo-manuscript-format"
```

Invoke with `/enggeo-manuscript-format`.

The repository is public and can be cloned without GitHub authentication.

## Microsoft Word preference

Use locally installed Microsoft Word for document editing, pagination and final rendering. LibreOffice/soffice is not an automatic fallback. Structural Python/OpenXML helpers remain available, followed by Word-based visual QA. Export a PDF for inspection with `scripts/export_word_pdf.ps1 -InputDocx <absolute-docx-path> -OutputPdf <absolute-pdf-path>`. The export opens the source read-only and does not save changes or refresh external links/citation-manager fields.

## v1.0 front-matter pagination

Page 1 contains the title, authors, affiliations and corresponding-author details. Abstract starts on page 2, followed by Keywords. Main text begins on a new page (normally page 3). Insert missing boundary breaks when revising existing documents. Check the actual boundary locations and Microsoft Word pagination; counting total breaks is insufficient.
