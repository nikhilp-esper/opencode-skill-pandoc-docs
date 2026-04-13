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

## Bundled files

| File | Purpose |
|------|---------|
| `SKILL.md` | Skill instructions (loaded by OpenCode) |
| `reference/esper-reference.docx` | DOCX reference template for consistent styling |
| `reference/github.css` | GitHub-style CSS for HTML output |

## Customizing the template

To use your own DOCX styling:

1. Open `reference/esper-reference.docx` in Word or Google Docs
2. Modify the heading styles, fonts, margins, and colors
3. Save back as `.docx`

Pandoc uses the styles from the reference doc, not the content.

## License

MIT
