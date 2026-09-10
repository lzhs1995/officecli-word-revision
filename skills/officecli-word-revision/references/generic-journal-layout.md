# Generic journal figure and table layout

Use this fallback only when the journal or user has not supplied a conflicting format. Precedence is: explicit venue requirements, explicit user instructions, stable source-master conventions, then this policy. RUC doctoral-thesis formatting is separate and does not inherit these rules.

## Figures and pagination

- Remove artificial page breaks before resizing a figure or table. Render through Microsoft Word after every pagination change.
- Keep a figure within 90% of the text width. Its height ceiling is derived from the usable text box: usable height minus its caption block and two narrative lines. The usable box is header/footer aware even when those parts contain no visible text.
- Every page containing a figure or data table in the active revision scope must also contain at least two narrative lines or 40 narrative characters. A section-level job may declare exact start/end anchors through `pagination.object_scope`; a full-manuscript job must leave that scope unset. Locate raster figures with `pdfimages -list`; a caption in the PDF text layer is not proof that the figure was located.
- Repair in this order, for at most two cycles: reflow existing narrative; then reduce the object in 5% steps within the derived limit. Splitting or redrawing a figure can change cross-references and requires separate approval.
- Treat an 8 pt minimum label size as a manual warning unless the source plotting program or vector asset supplies a measurable font size.

## Data tables

- Classify a visible cell as left aligned when it contains Han characters. Also keep an English prose sentence left aligned when it has at least 25 characters and three spaces. Align short English labels, formulas, letters, numbers, dates, intervals, and their punctuation to the right. Apply the same literal rule to headers.
- Resolve empty-cell alignment from the majority semantic type of its real grid columns. Accumulate `gridSpan`; a cell's XML position is not necessarily its column number. Missing or tied column semantics are hard failures.
- Set paragraph first-line, hanging, left, and right indents to zero in BOTH length and character units. Explicitly write character-unit zeros; removing them re-enables inherited indents. Put compact horizontal padding in `tcMar` (102 twip on each side).
- Resolve effective table font size from direct runs, the paragraph style inheritance chain, then `docDefaults`. Use exact 12 pt spacing for effective sizes up to 10 pt and exact 14 pt above 10 pt; before/after spacing is zero.
- Mark each data row `cantSplit`. Repeat the table's actual header rows (including an N row when part of the header). Preserve three-line borders, merged cells and Unicode subscripts. Keep a categorical variable heading with its first level; do not bind all rows of a long descriptive table together.
- Unless an explicit user/venue width rule applies, use `publication_layout.table.fit_mode=content`. After typography and indentation are finalized, perform Microsoft Word's native **AutoFit to Contents**, save, and record the selected table indexes and completion marker. A `w:tblLayout` flag or `flextable::autofit()` alone is not this operation. Explicit width rules use `fixed`/`preserve`; never silently convert a content-fit table back to fixed narrow columns.
- Fit only the clean accepted-content candidate. Transfer its saved table/cell/grid geometry into the tracked document with native format revisions, then validate accept/reject semantics and final Word/PDF layout. Do not add hard line breaks or shrink fonts to fit confidence intervals. Long tables may wrap or continue with repeated headers.

## Diagram carriers and verification

A manuscript may use a table as a drawing canvas. When cells containing at least three arrow characters are a majority of its nonempty cells, preserve the whole table's alignment, indentation, spacing, margins, and header behavior. Report the table index and retained nonzero indents; do not silently count it as a formatted data table.

The final DOCX gate must first prove that its checked table/cell sets equal the target sets, then report zero violations. Mutation tests must demonstrate that wrong alignment, nonzero indentation, wrong line spacing, and an object-only page each fail. The final Word-rendered PDF still requires visual inspection for wrapping, clipping, overlap, excessive whitespace, and readable labels.

An explicit user policy may choose `alignment_mode=han-left-nonhan-right`, which keeps every non-Han cell right-aligned, including long English text. Do not reintroduce the prose exception in that mode. Resolve empty cells from actual grid columns including merged spans.

For uniform-font jobs, use `format_contract` and `manuscript_format.py`: specify script-specific fonts, point sizes and line spacing; do not hardcode a previous project's 9 pt or Arial. A false bold flag must remain false. Table/figure captions have a one-to-one object binding and must be above/below respectively. Verify both tracked accepted-view content and the Word-saved clean file, with all target cells covered. Source diagram geometry and school-specific thesis rules are not overridden by these manuscript settings.

The fallback incorporates durable ideas from `econ-writing-skill` (self-contained, information-dense figures/tables and consistent precision) and `academic-research-skills` (rendered visual QA, traceability, bounded repair). It does not import LaTeX macros, APA caption placement, generic journal page limits, or duplicated vendor instructions from `academic-research-skills-codex`.
