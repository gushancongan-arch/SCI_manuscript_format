# SCI Manuscript Format

A reusable Codex and Claude Code skill for creating, revising, and auditing English-language SCI manuscript DOCX files. The repository follows the latest `liuhuaxin-mines/SCI_manuscript_format` rules as its baseline and adds the owner's reviewed amendments for line numbering, citation preservation, reference order, and figure-caption alignment.

This is a general draft convention. Current target-journal instructions override a conflicting rule only when they are supplied or verified.

## Included files

- `SKILL.md` — skill entry point and workflow
- `references/format-spec.md` — authoritative reviewed specification
- `assets/SCI_manuscript_template.docx` — reusable manuscript template
- `scripts/build_template.py` — deterministic template generator
- `scripts/audit_template.py` — structural template audit
- `agents/openai.yaml` — Codex display metadata

## Current core conventions

- A4 portrait with 2.54 cm margins and a centered dynamic footer page number
- continuous Word line numbering in every section; `Line Number` uses Times New Roman 11 pt
- `Normal` body text: Times New Roman 12 pt, justified, 1.5 spacing, 6 pt before and after, no indent
- exact English manuscript-role styles: `Normal`, `Figure`, `Figure Caption`, `Table Caption`, `Equation`, `Table`, `Reference`, and `Heading 1`–`Heading 3`
- 12 pt italic justified figure captions below figures
- 12 pt italic left-aligned table captions above three-line tables
- native editable Word OMML equations
- preservation of the source author-year or numeric citation system, source citation typography, and exact reference order

## Regenerate and audit the template

Install the packages listed in `requirements.txt`, then run:

```powershell
python scripts/build_template.py
python scripts/audit_template.py assets/SCI_manuscript_template.docx
```

## Install for Codex

```powershell
git clone https://github.com/gushancongan-arch/SCI_manuscript_format.git "$HOME\.agents\skills\sci-manuscript-format"
```

Invoke with `$sci-manuscript-format`.

## Install for Claude Code

```powershell
git clone https://github.com/gushancongan-arch/SCI_manuscript_format.git "$HOME\.claude\skills\SCI_manuscript_format"
```

Invoke with `/SCI_manuscript_format`.

The repository is public and can be cloned without GitHub authentication.
