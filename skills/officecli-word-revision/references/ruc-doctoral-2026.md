# RUC doctoral thesis profile (2026 baseline)

This profile encodes the locally supplied Chinese Renmin University doctoral-thesis requirements. It is versioned and must not be silently changed when requirements evolve. “2026 baseline” identifies the local requirements snapshot; do not describe it as the university's latest official publication unless an official dated source has been independently verified.

Core rules:

- A4 inner pages; 45 mm top, 40 mm bottom, 35 mm inner, and 30 mm outer margins; mirrored margins and different odd/even pages.
- Centered thesis-title header in SimSun 10.5 pt, 25 mm from the edge. Page number in Times New Roman 10.5 pt, 25 mm from the edge, on the outer footer.
- Front matter uses lower Roman numbers. The main text uses decimal numbering starting at 1. Full-thesis jobs require explicit front/body section anchors before changing numbering.
- Body: SimSun 12 pt, Times New Roman for Latin text, justified, 24 pt first-line indent, exact 20 pt line spacing, no paragraph spacing.
- Chapter title: SimHei 18 pt centered; first- and second-level section headings: SimHei 14/12 pt left aligned. Heading spacing follows the profile JSON.
- Captions and table text: SimSun 10.5 pt. Images remain inline and their containing paragraph must not use an exact line height that clips the image.
- Three-line tables have a top rule, header-bottom rule, and bottom rule only. Preserve intentional landscape sections and allow rows to grow rather than clipping them.
- Footnotes use SimSun 9 pt and page-bottom placement. Bibliographic content remains governed by Zotero/CSL; OfficeCLI applies only layout and checks field integrity.

The A3 cover, physical duplex printer setting, and substantive GB/T 7714 metadata correctness are outside automated formatting. Microsoft Word pagination and field refresh are the final authority on macOS.

The pipeline does not set Word's global `updateFields`/`recalcFields` flags in the clean output. Those flags can trigger external-field dialogs and may touch Zotero fields. Full-thesis jobs instead use the targeted field-refresh pass described in [thesis-format.md](thesis-format.md).
