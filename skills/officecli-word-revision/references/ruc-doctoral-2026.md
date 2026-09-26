# RUC doctoral thesis profile (2026 baseline)

This profile encodes the locally supplied Chinese Renmin University doctoral-thesis requirements. It is versioned and must not be silently changed when requirements evolve. “2026 baseline” identifies the local requirements snapshot; do not describe it as the university's latest official publication unless an official dated source has been independently verified.

Core rules:

- A4 inner pages; 45 mm top, 40 mm bottom, 35 mm inner, and 30 mm outer margins; mirrored margins and different odd/even pages.
- Centered thesis-title header in SimSun 10.5 pt, 25 mm from the edge. Page number in Times New Roman 10.5 pt, 25 mm from the edge, on the outer footer.
- The snapshot specifies Roman front matter and illustrates uppercase I/II/III. `lowerRoman` remains the compatibility profile default, not a verified case requirement; select case explicitly when the template or user settles it. The main text uses decimal numbering starting at 1. Full-thesis jobs require explicit front/body section anchors before changing numbering.
- Body: SimSun 12 pt, Times New Roman for Latin text, justified, 24 pt first-line indent, exact 20 pt line spacing, no paragraph spacing.
- Chapter title: SimHei 18 pt centered; first- and second-level section headings: SimHei 14/12 pt left aligned; tertiary heading: SimSun 12 pt. Profile 1.1.0 corrects the tertiary role and adds separate TOC, backmatter, table-caption and figure-caption roles. The snapshot's “occupied lines” and before/after values are not internally sufficient to derive one total height; the profile retains its earlier spacing translation pending a task-specific choice.
- Captions and table text: SimSun 10.5 pt. Images remain inline and their containing paragraph must not use an exact line height that clips the image.
- Three-line tables have a top rule, header-bottom rule, and bottom rule only. Preserve intentional landscape sections and allow rows to grow rather than clipping them.
- Footnotes use SimSun 9 pt, page-bottom placement and per-page numbering in the supplied snapshot. Inspect document defaults and section overrides. Bibliographic content remains governed by Zotero/CSL; OfficeCLI applies layout and checks field integrity. English-first/full-pinyin sorting is this user's project rule, not an institution-wide rule inferred from the snapshot.

See [full-thesis guidance](full-thesis-formatting.md), [source matrix](ruc-rule-sources.json) and
[the optional integrity contract](thesis-integrity.md). Separate content-block pagination from section-level settings;
table captions keep the first row and picture paragraphs keep the following figure caption. The generic caption role
does not force every caption to keep subsequent narrative text.

The A3 cover, physical duplex printer setting, and substantive GB/T 7714 metadata correctness are outside automated formatting. Microsoft Word pagination and field refresh are the final authority on macOS.

The pipeline does not set Word's global `updateFields`/`recalcFields` flags in the clean output. Those flags can trigger external-field dialogs and may touch Zotero fields. Full-thesis jobs instead use the targeted field-refresh pass described in [thesis-format.md](thesis-format.md).
