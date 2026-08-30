# SCI_manuscript_format

A reusable Word-formatting skill for SCI manuscript templates and manuscript-format audits. It is designed for both Codex and Claude Code.

This repository records a user-reviewed manuscript convention. It does not claim to replace the current author guidelines of any specific journal.

## Included files

- `SKILL.md` — skill entry point and workflow
- `references/format-spec.md` — reviewed formatting specification
- `agents/openai.yaml` — Codex display metadata

## Install for Codex

Clone the repository into the shared skill directory:

```powershell
git clone https://github.com/gushancongan-arch/SCI_manuscript_format.git "$HOME\.agents\skills\sci-manuscript-format"
```

Invoke it with:

```text
$sci-manuscript-format
```

## Install for Claude Code

Clone the repository into the Claude Code skill directory:

```powershell
git clone https://github.com/gushancongan-arch/SCI_manuscript_format.git "$HOME\.claude\skills\SCI_manuscript_format"
```

Invoke it with:

```text
/SCI_manuscript_format
```

The repository is private by default, so cloning requires access through the repository owner account or an invited collaborator.
