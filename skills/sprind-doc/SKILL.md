---
name: sprind-doc
description: >
  Format documents with SPRIND (Bundesagentur für Sprunginnovationen / Federal Agency for Disruptive Innovation)
  branding. Creates .docx files with the SPRIND letterhead logo, GT America fonts, version/date footer, and
  page numbers. Use when the user wants to create, format, or convert any document into SPRIND branded format.
  Triggers: mentions of SPRIND, SPRIN-D, Sprunginnovationen, or requests for SPRIND document formatting/letterhead.
---

# SPRIND Document Formatter

Create professional SPRIND-branded .docx documents with proper letterhead, typography, and page structure.

## What This Skill Produces

- A4 .docx built from the official SPRIND Word template (preserves headers, background images, styles)
- GT America Extended (headings) and GT America Light (body) typography at correct sizes
- Right-aligned footer with version, date, and page numbers
- Footer language auto-detected from content (English or German format)
- Tables preserved when converting from existing .docx files
- Large-font Normal paragraphs auto-detected as headings

## Workflow

### 1. Gather Content

Accept content from the user in any of these forms:
- **Markdown text** provided directly in the conversation
- **Plain text** provided directly
- **Path to an existing file** (.md, .txt, or .docx to reformat)

Also collect:
- **Version number** (e.g., "1.0", "2.1") — optional, omit if the user doesn't specify one
- **Date** — defaults to today if not specified; accepts DD.MM.YYYY, YYYY-MM-DD, or natural language
- **Language** — auto-detected from content; user can override with "en" or "de"
- **Output filename** — ask if not obvious from context

### 2. Prepare the Input

If the user provides content inline (not a file path), save it to a temporary .md file first:

```python
# Save inline content to temp file
import tempfile
with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as f:
    f.write(content)
    temp_path = f.name
```

### 3. Run the Formatter

Execute the formatting script:

```bash
python3 "{SKILL_DIR}/scripts/sprind_format.py" \
  --input "INPUT_PATH" \
  --output "OUTPUT_PATH.docx" \
  --version "VERSION" \
  --date "DATE"
```

**Arguments:**
| Flag | Required | Description |
|------|----------|-------------|
| `--input`, `-i` | Yes* | Input file (.md, .txt, .docx). *Can also pipe via stdin. |
| `--output`, `-o` | Yes | Output .docx path |
| `--version`, `-v` | No | Version string (e.g., "1.0") |
| `--date`, `-d` | No | Date (default: today) |
| `--language`, `-l` | No | "en" or "de" (auto-detected) |
| `--title`, `-t` | No | Document title (prepended as H1 if content has no heading) |

### 4. Return Result

Tell the user the output path. Remind them:
- The .docx uses GT America fonts — these must be installed to render correctly
- To create a PDF, open the .docx in Word/Pages and export

## Style Mapping

The script maps content structure to SPRIND styles:

| Input | SPRIND Style | Font |
|-------|-------------|------|
| `# Heading` (H1) | SPRIND - Überschrift | GT America Extended, 16.5pt |
| `## Subheading` (H2+) | SPRIND - Unterüberschrift | GT America Extended, 11pt |
| Body paragraphs | SPRIND - Paragraph | GT America Light, 11pt |
| `- Bullet items` | SPRIND - Auflistung | GT America Light, 11pt |
| `1. Numbered items` | SPRIND - Aufzählung | GT America Light, 11pt |
| Tables | Preserved with GT America Light, 11pt | |
| `**bold**` | Bold run | |
| `*italic*` | Italic run | |
| `__underline__` | Underline run (SPRIND - Hervorhebung) | |

When converting from .docx, the script also detects:
- `Heading 1`/`Heading 2` styles → SPRIND heading styles
- `List Paragraph` → SPRIND - Auflistung
- Large-font Normal paragraphs (16pt+) → SPRIND - Überschrift

## Footer Formats

Determined automatically by content language:

**English:** First page shows `Version X.Y from 19th March 2024, Page 1 (3)`, continuation pages show `1 (3)`

**German:** All pages show `15.02.2025   Seite 1/3`

## Formatting Reference

For detailed measurements (margins, spacing, XML structure), read `references/formatting-spec.md`.

## Dependencies

- `python-docx` (pip3 install python-docx)
- GT America fonts installed at `/Library/Fonts/` (GT-America-Extended-Regular.otf, GT-America-Light.otf)

## Important Notes

- This skill produces .docx only — the user converts to PDF manually
- The document is built from the official SPRIND .dotx template (included in assets), so all header images, styles, and numbering are inherited correctly
- Each page has the SPRIND logo; first page additionally shows the agency name in German and English
- The background page images are bundled as `assets/page_bg_first.png` and `assets/page_bg_continuation.png` for reference, but the actual template (`assets/SPRIND_Vorlage_GTAmerica.dotx`) is what gets used
