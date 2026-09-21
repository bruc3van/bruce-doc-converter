# Conversion regression corpus

These small, synthetic documents contain no customer data. Regenerate the binary
fixtures with `python scripts/generate_fixtures.py` after intentionally changing
their content. Tests assert document structure and meaningful text rather than ZIP
timestamps or byte-for-byte Office output.

- `report.docx`: heading, paragraphs, and list content from the existing plugin fixture.
- `budget.xlsx`: numbers, an uncached formula at B4, and a cached formula at C4.
- `slides.pptx`: a title and checklist slide.
- `report.pdf`: a minimal text PDF with a real page/font/content stream.
- `semantic.md`: CommonMark escaping, code fences, loose/nested lists, links,
  tables, emphasis, and literal HTML; its generated DOCX XML is verified.
