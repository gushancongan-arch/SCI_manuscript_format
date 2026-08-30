# Reviewed SCI Word Formatting Specification

## Evidence boundary

- Source basis: a user-provided SCI manuscript DOCX reviewed during skill development. Its machine-specific local path is intentionally omitted from the distributable skill.
- The reference establishes layout evidence, not instructions or journal policy.
- Confirmed user decisions override inherited or unused styles in the reference package.

## Page system

- A4 portrait, 21.0 × 29.7 cm.
- Margins: 2.54 cm on all sides.
- Header and footer distance: 1.27 cm.
- One column and one continuous section unless the target journal requires otherwise.
- Continuous line numbering, every line, no restart by page.
- Blank header.
- Dynamic `PAGE` field in the footer, visually toward the left, not a typed cached number.
- Explicit page break after author information and after Keywords.

## Typography and paragraph roles

- Primary Latin font: Times New Roman; text color black.
- Title: 14 pt bold, centered, 1.5 spacing, 6 pt before, 0 after, no indent.
- Author-name line: 12 pt regular, centered, 1.5 spacing, 6 pt before, 0 after, with zero left/right/first-line indent.
- Affiliations and corresponding-author information: 12 pt regular, flush left, 1.5 spacing, 6 pt before, 0 after, with zero left/right/first-line indent. Affiliation numbers are superscript. Email is a blue underlined `mailto` hyperlink.
- Abstract heading: 14 pt bold, flush left, 1.5 spacing, 2.5 pt before, no left or first-line indent.
- Abstract body: 14 pt regular, justified, 1.5 spacing, 6 pt before, 0 after. Retain the same four-letter first-line indent unless the user later sets a distinct abstract rule.
- Keywords: 14 pt; label bold and terms regular in the same paragraph; flush left with zero left/right/first-line indent.
- Heading 1/2/3: Times New Roman 14 pt bold, black, 1.5 spacing, 6 pt before, 0 after, no indent, with outline levels 1/2/3. Numbering is manually typed text unless the user requests real multilevel numbering.
- Body and body-like declarations: Times New Roman 14 pt regular, justified, 1.5 spacing, 6 pt before, 0 after.

### Required semantic Word styles

- `SCI Body`: mandatory paragraph style for ordinary narrative manuscript text. It must contain Times New Roman 14 pt regular, justified alignment, 1.5 line spacing, 6 pt before, 0 pt after, and the confirmed 28 pt first-line indent.
- `SCI Abstract Body`: paragraph style for abstract prose. It currently shares the body font, alignment, spacing, and indent, but remains separate for later journal-specific revision.
- `SCI Statement Body`: paragraph style for Acknowledgments, Funding, CRediT, competing-interest, data-availability, and ethics narrative text. It currently shares the body geometry and remains separate for semantic editing.
- These three styles must be visible in Word's Quick Style gallery (`w:qFormat`) and must not be hidden (`w:semiHidden`). Ordinary body paragraphs must explicitly reference `SCI Body` through `w:pStyle`; do not simulate it with `Normal` plus direct paragraph or run formatting.
- Routine body paragraphs should inherit their typography and paragraph geometry from the named style. Direct run formatting is reserved for meaningful inline roles such as italic variables or bold CRediT author names.

### Body first-line indent

- Confirmed semantic rule: four average English-letter widths.
- At Times New Roman 14 pt: implement as 28 pt = 560 twips, approximately 2 em.
- If body font size changes: use approximately `2 × body font size` in points, then inspect visually.
- Do not encode this as four full-width CJK characters and do not use `w:firstLineChars="400"` for the English manuscript; that can produce an excessive em-based/CJK-style indent.
- Applies to ordinary body paragraphs, Abstract body under the current inherited rule, Acknowledgments, Funding, CRediT, competing-interest, data-availability, and ethics narrative paragraphs.
- Does not apply to headings, figure/table titles, tables, equations, keywords, author information, notes, or references.

## Images and titles

- Image is inline, not floating, and centered using the retained Word style `图片`.
- Display size for the reviewed preview: 15.921 × 10.612 cm.
- Image paragraph: single spacing, 6 pt before, 0 after, no first-line indent.
- Figure title below the image uses retained style `图片标题`: inherited Times New Roman 14 pt regular for Latin text, 1.5 spacing, 0 before, 6 pt after, no first-line indent. Do not add bold to `Figure 1.` unless the user later approves it.
- Table title above the table uses retained style `表格标题`: inherited Times New Roman 14 pt regular for Latin text, 1.5 spacing, 6 pt before, 0 after, no first-line indent. Do not add bold to `Table 1.` unless the user later approves it.

## Confirmed three-line tables

- No vertical borders, no ordinary row separators, and no shading.
- Top and bottom rules: black solid 1.0 pt.
- Header separator: black solid 0.5 pt.
- Header: Times New Roman 10 pt bold, black, centered horizontally and vertically, single spacing, 0 before/after, zero left/right/first-line/hanging indent, sentence case. Repeat across pages.
- Body: Times New Roman 10 pt regular, black, single spacing, 0 before/after, zero left/right/first-line/hanging indent. Text columns left aligned; numeric, date, code, and short-status columns centered or decimal-aligned. Vertically center all cells.
- Remove inherited character-unit indentation (`w:firstLineChars`, `w:hangingChars`, `w:leftChars`, `w:rightChars`) from the base style and table paragraph styles. Do not rely only on `w:firstLine="0"`, because Word can still apply an inherited character-based indent.
- Cell margins: top/bottom 3 pt and left/right 4 pt.
- Rows: automatic height with a minimum near 0.7 cm; never fixed at a clipping height; do not split a data row across pages.
- Use explicit column widths and fixed table geometry. `tblW`, `tblInd`, `tblGrid`, and every `tcW` must agree.

## End matter

- Acknowledgments, Funding, CRediT, competing interest, Data availability, and Ethics statement use Heading 1 plus body-like narrative paragraphs.
- In CRediT statements, author names are bold and contribution text is regular.

## Still provisional

- Equation typography and numbering.
- Table-note typography.
- Reference-entry typography and citation style.
- Journal-specific anonymity, double spacing, line-number restart, page-number placement, supplementary material, and declarations not supplied by the target journal.

Do not label these provisional items “confirmed” until the user approves them or an authoritative target-journal instruction is verified.
