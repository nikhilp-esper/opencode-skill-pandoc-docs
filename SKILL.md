---
name: pandoc-docs
description: "Create formatted documents with pandoc and upload to Google Drive."
license: MIT
compatibility: opencode
metadata:
  audience: agents
  domain: documents
  requires:
    bins: ["pandoc", "gws"]
---

# pandoc-docs

Create professional Esper-branded documents from markdown using pandoc, then upload to Google Drive.

## When to use this skill

- User asks you to create a document, RFC, report, brief, proposal, or any formatted deliverable
- User asks to export content to DOCX, PDF, or PPTX
- User asks to upload a document to Google Drive

## Template Selection (IMPORTANT)

Choose the right template based on document purpose:

| Template | Use Case | Flags |
|----------|----------|-------|
| `esper-brief.docx` | Meeting prep briefs, customer summaries, 1-2 page dense docs | None |
| `esper-technical.docx` | Technical analyses, RFCs, architecture docs | `--number-sections` |
| `esper-executive.docx` | One-pagers, exec summaries, decision docs | None |
| `esper-report.docx` | QBRs, formal reports, multi-page deliverables | None |
| `esper-proposal.docx` | RFP responses, proposals, formal submissions | `--number-sections` |

**Decision guide:**
- Short (1-2 pages), dense info, internal use → `esper-brief`
- Technical depth, code blocks, engineering audience → `esper-technical`
- Single page, metrics-heavy, exec/leadership audience → `esper-executive`
- Multi-page, customer-facing, formal tone → `esper-report`
- Responding to a customer request/RFP, requirements matrix → `esper-proposal`

> [!IMPORTANT]
> Never use `--toc`. Tables of contents are not wanted in any Esper document.

## Workflow

### 1. Write content as markdown

Write document content to a `.md` file in `/tmp/`. Use standard markdown with pandoc extensions:

```markdown
---
title: "Document Title"
subtitle: "Context or author"
date: "June 2026"
---

# First Section

Body text with **bold** and *italic*.

| Column A | Column B |
|----------|----------|
| data     | data     |
```

The YAML frontmatter sets document metadata (title page, headers).

### 2. Convert with pandoc

**Standard command (all templates):**

```bash
pandoc /tmp/doc.md -o /tmp/doc.docx \
  --reference-doc=SKILL_DIR/reference/TEMPLATE_NAME.docx \
  --from=markdown+smart+pipe_tables+yaml_metadata_block \
  --lua-filter=SKILL_DIR/filters/strip-bookmarks.lua
```

**With numbered sections (technical, proposal):**

```bash
pandoc /tmp/doc.md -o /tmp/doc.docx \
  --reference-doc=SKILL_DIR/reference/esper-technical.docx \
  --from=markdown+smart+pipe_tables+yaml_metadata_block \
  --lua-filter=SKILL_DIR/filters/strip-bookmarks.lua \
  --number-sections
```

Replace `SKILL_DIR` with the resolved path to this skill's directory.

**Always include both `--reference-doc` and `--lua-filter`. Never use `--toc`.**

### 3. Upload to Google Drive

#### First upload (creates a new file)

```bash
gws drive +upload /tmp/doc.docx --name "Document Title.docx"
gws drive +upload /tmp/doc.docx --parent FOLDER_ID
```

Save the returned file `id` for future updates.

#### Update an existing file (same link, new content)

```bash
gws drive files update \
  --params '{"fileId": "EXISTING_FILE_ID"}' \
  --upload /tmp/doc.docx
```

> [!CAUTION]
> Confirm with the user before uploading or updating. Show the filename and destination.

### 4. Get the Google Docs link

After upload, build the link from the file ID:
- Google Docs: `https://docs.google.com/document/d/FILE_ID/edit`
- Drive file: `https://drive.google.com/file/d/FILE_ID/view`

DOCX files auto-convert to Google Docs format on Drive.

## Template Details

### esper-brief.docx — Customer Brief

Best for: meeting prep, customer summaries, handoff docs.

- No cover page — starts immediately with title
- Tight line spacing (1.15) for density
- "At a Glance" summary tables work great
- Navy header tables with alternating row shading
- Footer: "Confidential - Esper Inc."
- Margins: 0.85" (slightly tighter for more content)

**Example frontmatter:**
```yaml
---
title: "Meeting Brief: Customer Name"
subtitle: "Weekly Sync — June 15, 2026"
---
```

### esper-technical.docx — Technical Document

Best for: RFCs, architecture docs, analyses, engineering deliverables.

- Heading 1 has a navy rule (line) below it for visual separation
- Code blocks get light gray background
- Designed for multi-page documents
- Page numbers in footer
- Standard 1.0" margins

**Example frontmatter:**
```yaml
---
title: "RFC: Feature Name"
subtitle: "Esper Engineering"
author: "Author Name"
date: "June 2026"
---
```

**Flags:** `--number-sections`

### esper-executive.docx — Executive Summary

Best for: one-pagers, decision docs, exec briefs.

- Large title (24pt) for impact
- Metrics work well in a borderless table (bold values, gray labels)
- Call-to-action sections get a light blue background
- Tight spacing for single-page density
- No footer (cleaner for exec audience)

**Example frontmatter:**
```yaml
---
title: "Executive Summary"
subtitle: "Acme Corp — Expansion Opportunity"
---
```

### esper-report.docx — QBR / Formal Report

Best for: QBRs, customer reports, multi-page formal deliverables.

- Cover page section (centered title, subtitle, date)
- Use `\newpage` after frontmatter for cover page effect
- Header: "Esper" text, Footer: "Confidential"
- Heading 1 with navy rule below
- Designed for 5-15 page documents

**Example frontmatter:**
```yaml
---
title: "Quarterly Business Review"
subtitle: "Prepared for Customer Name"
date: "Q2 2026"
---
```

**Flags:** None

**Cover page pattern:**
```markdown
---
title: ""
---

\begin{center}
\vspace*{2cm}
{\Huge Quarterly Business Review}

\vspace{0.5cm}
{\large Prepared for Acme Corp}

\vspace{1cm}
{\normalsize Q2 2026}
\end{center}

\newpage

# Executive Summary
...
```

Note: The LaTeX commands above only work with PDF output. For DOCX, just use a title in frontmatter — pandoc generates the title block automatically.

### esper-proposal.docx — Proposal / RFP Response

Best for: RFP responses, formal proposals, contract submissions.

- Cover page with title, subtitle, submission details
- Numbered sections (1. 1.1. 1.1.1.) for RFP cross-referencing
- Requirements matrix tables (navy header, status column)
- Implementation timeline as numbered list
- Footer: "Confidential - Esper Inc."

**Example frontmatter:**
```yaml
---
title: "Enterprise MDM Solution Proposal"
subtitle: "In Response to Customer RFP-2026-042"
date: "June 2026"
---
```

**Flags:** `--number-sections`

## Styling Reference (All Templates)

All templates share consistent brand styling:

- **Headings**: Calibri Light, navy (#1A237E), not bolded
- **Body**: Calibri 11pt, dark gray (#212121)
- **Tables**: Navy header row with white text, alternating #F5F8FC rows, thin gray borders
- **Lists**: Calibri 11pt, tight spacing
- **Quotes**: Gray italic, left-indented
- **Code**: Consolas 10pt, light gray background
- **Links**: Blue (#0288D1) — set by Word/Docs rendering

## Pandoc Markdown Tips

### Tables

```markdown
| Feature | Status | Notes |
|---------|--------|-------|
| Auth    | Done   | OAuth2 |
| API     | WIP    | v2 endpoints |
```

### Code blocks

````markdown
```kotlin
fun hello() = "world"
```
````

### Page breaks (DOCX only)

```markdown
\newpage
```

### Images

```markdown
![Caption text](path/to/image.png){ width=80% }
```

### Definition lists

```markdown
Term
:   Definition text here.
```

## PDF Output

If PDF is specifically requested:

```bash
pandoc /tmp/doc.md -o /tmp/doc.pdf \
  --from=markdown+smart+pipe_tables+yaml_metadata_block \
  --pdf-engine=pdflatex \
  -V geometry:margin=1in \
  -V fontsize=11pt
```

If `pdflatex` is unavailable, generate DOCX and let the user export to PDF from Google Docs.

## PPTX Output

For basic presentations (Marp or Gamma API are better for polished decks):

```bash
pandoc /tmp/slides.md -o /tmp/slides.pptx \
  --from=markdown+smart+pipe_tables+yaml_metadata_block \
  --slide-level=2
```

Each `##` heading starts a new slide.

## Bundled Files

| File | Purpose |
|------|---------|
| `reference/esper-brief.docx` | Dense brief/meeting prep template |
| `reference/esper-technical.docx` | Technical doc with ruled headings and code styling |
| `reference/esper-executive.docx` | High-impact one-pager template |
| `reference/esper-report.docx` | Formal multi-page report with cover section |
| `reference/esper-proposal.docx` | RFP/proposal with numbered sections |
| `reference/esper-reference.docx` | Legacy generic template (deprecated — use specific templates above) |
| `filters/strip-bookmarks.lua` | Removes pandoc heading anchor bookmarks from DOCX |
| `reference/github.css` | CSS for HTML output |

## Troubleshooting

- **Tables have no styling**: Always use `--reference-doc` — pandoc default has no table formatting
- **Bookmark icons next to headings**: Always include `--lua-filter=SKILL_DIR/filters/strip-bookmarks.lua`
- **No title page**: Ensure YAML frontmatter is at the very top with `---` delimiters
- **Upload fails**: Run `gws auth login` to refresh OAuth
- **Missing pdflatex**: Use DOCX output instead, let user export to PDF from Google Docs
