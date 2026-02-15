# SPRIND Document Formatting Specification

## Page Layout

- **Size**: A4 (210 × 297 mm)
- **Orientation**: Portrait
- **Margins**:
  - Top: 32.21 mm (915 twips)
  - Bottom: 25 mm (709 twips)
  - Left: 25 mm (709 twips)
  - Right: 25 mm (709 twips)
- **Header distance**: 0 mm (header carries full-page background image)
- **Footer distance**: 12.5 mm
- **Gutter**: 0 mm

## Header Background Images

The SPRIND letterhead is implemented as a full-page PNG image anchored behind text in the header:

- **First page** (`page_bg_first.png`): Contains SPRIN-D logo top-left + "BUNDESAGENTUR FÜR SPRUNGINNOVATIONEN" / "FEDERAL AGENCY FOR DISRUPTIVE INNOVATION" top-right
- **Continuation pages** (`page_bg_continuation.png`): Contains SPRIN-D logo top-left only
- **Image size**: 2480 × 3508 px (300 DPI = A4)
- **Anchor**: `behindDoc="1"`, position relative to page at offset (0, 0)
- **Extent**: 7556400 × 10681200 EMU (210 × 297 mm)
- **`different_first_page_header_footer`**: True

### XML Structure for Header Image

The image is placed inside a `wp:anchor` element within a paragraph run in the header:

```xml
<wp:anchor behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"
           distT="0" distB="0" distL="0" distR="0" simplePos="0">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
  <wp:extent cx="7556400" cy="10681200"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="1" name="Header Background"/>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr>...</pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="rIdX"/>
          <a:stretch><a:fillRect/></a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="7556400" cy="10681200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:anchor>
```

## Fonts

| Role | Font Family | Weight | Size |
|------|------------|--------|------|
| Heading (Überschrift) | GT America Extended | Regular | 16.5pt |
| Subheading (Unterüberschrift) | GT America Extended | Regular | 11pt |
| Body (Paragraph) | GT America Light | — | 11pt |
| Bullet list (Auflistung) | GT America Light | — | 11pt |
| Numbered list (Aufzählung) | GT America Light | — | 11pt |
| Emphasis (Hervorhebung) | GT America Light | — | 11pt, underline |
| Footnote (Fußnote) | GT America Light | — | 10pt |
| Footer (English) | GT America Light | — | 10pt |
| Footer (German) | GT America Light | — | 8pt |

**Font files** (installed at `/Library/Fonts/`):
- GT-America-Extended-Regular.otf
- GT-America-Extended-Regular-Italic.otf
- GT-America-Light.otf
- GT-America-Light-Italic.otf

## Paragraph Styles

### SPRIND - Uberschrift (Heading)
- Font: GT America Extended, 16.5pt
- Color: #000000
- Space before: 360 twips (6.35 mm)
- Space after: 360 twips (6.35 mm)

### SPRIND - Unteruberschrift (Subheading)
- Font: GT America Extended, 11pt
- Color: #000000
- Space before: 480 twips (8.47 mm)
- Space after: 240 twips (4.23 mm)

### SPRIND - Paragraph (Body)
- Font: GT America Light, 11pt
- Color: #000000
- Space before: 0
- Space after: 120 twips (2.12 mm)

### SPRIND - Auflistung (Bullet List)
- Font: GT America Light, 11pt
- Left indent: 357 twips (6.3 mm)
- Hanging indent: 357 twips (6.3 mm)
- Space after: 120 twips

### SPRIND - Aufzahlung (Numbered List)
- Font: GT America Light, 11pt
- Left indent: 357 twips (6.3 mm)
- Hanging indent: 357 twips (6.3 mm)
- Space after: 120 twips

### SPRIND - Nummerierte Unteruberschrift (Numbered Subheading)
- Font: GT America Extended, 11pt
- Left indent: 426 twips (7.5 mm)
- Hanging indent: 426 twips (7.5 mm)

### SPRIND - Hervorhebung (Character Style)
- Font: GT America Light, 11pt
- Underline: single
- Not bold, not italic

## Footer Formats

### English Documents
```
Version {version} from {date_english}, Page {PAGE} ({NUMPAGES})
```
- Example: `Version 2.1 from 19th March 2024, Page 1 (3)`
- Font: GT America Light, 10pt
- Alignment: right
- Spacing before: 360 twips (18pt)
- Applied to first page footer; continuation pages show only `{PAGE} ({NUMPAGES})`

### German Documents
```
{date_german}   Seite {PAGE}/{NUMPAGES}
```
- Example: `10.11.2023   Seite 1/11`
- Font: GT America Light, 8pt
- Alignment: right
- Same footer on all pages (`different_first_page_header_footer` still true for header images, but footer content is the same)

## Color

All SPRIND text uses **#000000** (black). No brand accent colors are defined in the document templates.

## Document Language

- English documents: `lang="en-GB"`
- German documents: `lang="de-DE"`
- Auto-detection: Check content for German-specific words (der, die, das, und, oder, für, über, etc.)
