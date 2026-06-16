# opencode-skill-pandoc-docs

An [OpenCode](https://opencode.ai) skill that creates professional documents using [pandoc](https://pandoc.org) and uploads them to Google Drive via the [gws CLI](https://github.com/googleworkspace/cli).

## What it does

When you ask your AI assistant to create a document, RFC, report, or technical analysis, this skill teaches it to:

1. Write content as markdown with YAML frontmatter
2. Convert to the right format (DOCX, PDF, PPTX, HTML) using pandoc
3. Apply consistent styling via a bundled reference template
4. Upload to Google Drive using the `gws` CLI
5. Return a shareable Google Docs link

## Prerequisites

- [OpenCode](https://opencode.ai) installed and configured
- [pandoc](https://pandoc.org/installing.html) 3.0+ on `$PATH`
- [gws CLI](https://github.com/googleworkspace/cli) on `$PATH` (for Google Drive uploads)

## Installation

Clone this repo into your OpenCode skills directory:

```bash
git clone https://github.com/nikhilp-esper/opencode-skill-pandoc-docs.git \
  ~/.opencode/skill/pandoc-docs
```

Restart OpenCode. The skill will appear in the available skills list.

## Usage

Just ask naturally:

- *"Create an RFC about device identity resolution and upload it to Drive"*
- *"Write a technical analysis of the serial number access methods and export as DOCX"*
- *"Make a sprint report for Sprint 188"*

The agent will load the skill automatically when the task matches document creation, or you can reference it explicitly.

## Supported output formats

| Format | Extension | Best for |
|--------|-----------|----------|
| DOCX | `.docx` | Google Docs (auto-converts on upload) |
| PDF | `.pdf` | Printable/archival (requires LaTeX) |
| PPTX | `.pptx` | Presentations |
| HTML | `.html` | Web/email |

## Templates

Five professional, lightly-branded reference templates ship with the skill. The
agent picks the right one based on document type (see `SKILL.md` for the
selection guide).

| Template | Use case | Recommended flags |
|----------|----------|-------------------|
| `reference/esper-brief.docx` | Meeting prep, customer briefs (dense, no cover) | — |
| `reference/esper-technical.docx` | RFCs, analyses, architecture docs | `--toc --number-sections` |
| `reference/esper-executive.docx` | One-pagers, executive summaries | — |
| `reference/esper-report.docx` | QBRs, formal multi-page reports | `--toc` |
| `reference/esper-proposal.docx` | RFP responses, proposals | `--toc --number-sections` |

All templates share consistent styling:

- Headings: Calibri Light, navy (`#1A237E`), not bolded
- Body: Calibri 11pt, dark gray (`#212121`)
- Tables: navy header row with white text, alternating `#F5F8FC` rows, thin gray borders
- No bookmarks, no embedded logos, no fragile assets

## Bundled files

| File | Purpose |
|------|---------|
| `SKILL.md` | Skill instructions (loaded by OpenCode) |
| `reference/esper-brief.docx` | Brief / meeting-prep template |
| `reference/esper-technical.docx` | Technical document template |
| `reference/esper-executive.docx` | Executive one-pager template |
| `reference/esper-report.docx` | Formal report / QBR template |
| `reference/esper-proposal.docx` | Proposal / RFP response template |
| `reference/esper-reference.docx` | Legacy generic template (deprecated) |
| `reference/github.css` | GitHub-style CSS for HTML output |
| `filters/strip-bookmarks.lua` | Strips pandoc heading-anchor bookmarks from DOCX |
| `scripts/generate_templates.py` | Regenerates all templates from code |

## Regenerating / customizing the templates

The templates are generated programmatically with [python-docx](https://python-docx.readthedocs.io)
so styling changes are version-controlled and reproducible.

```bash
pip install python-docx

# Regenerate into the repo's reference/ directory (default when run from the repo)
python3 scripts/generate_templates.py

# Or write to a custom location
PANDOC_TEMPLATE_DIR=~/my-templates python3 scripts/generate_templates.py
```

To change brand colors, fonts, or spacing, edit the constants and
`configure_base_styles()` near the top of `scripts/generate_templates.py`, then
re-run. Pandoc uses the *styles* from the reference doc, not its content.

## License

MIT
