# SCI Manuscript Format

A reusable Codex and Claude Code skill for creating, revising, and auditing English-language SCI manuscript DOCX files. It combines a user-reviewed formatting convention with a deterministic DOCX template, native Word equations, three-line tables, citation-system preservation, and structural/visual QA.

This repository records a draft convention. Current target-journal instructions override conflicting draft rules only when supplied or verified.

## Included files

- `SKILL.md` - skill entry point and workflow
- `references/format-spec.md` - authoritative reviewed specification
- `assets/SCI_manuscript_template.docx` - reusable manuscript template
- `scripts/build_template.py` - deterministic template generator
- `scripts/audit_template.py` - structural template audit
- `agents/openai.yaml` - Codex display metadata

## Confirmed user conventions

- Times New Roman 14 pt body text with a 28 pt first-line indent
- semantic body, abstract, statement, equation, caption, table, and reference styles
- continuous 11 pt line numbering and a dynamic page-number field
- three-line tables with fixed geometry and zero cell-paragraph indents
- 12 pt italic justified figure captions below figures
- 12 pt italic table captions above tables
- native editable Word OMML equations
- preservation of the source citation system and reference order

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
