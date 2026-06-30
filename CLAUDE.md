# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the extractor

```bash
# Interactive (prompts for PDF path and output name)
python mosfet_extractor.py

# Non-interactive (preferred when running from scripts/Claude)
python mosfet_extractor.py datasheet.pdf --output result.xlsx --no-prompt
```

Current version: **v19**. Supported manufacturers: Infineon · ST · ON Semi · Toshiba · Nexperia · Vishay · ROHM · Wolfspeed · IXYS.

## Dependencies

```bash
pip install pdfplumber openpyxl pymupdf opencv-python numpy scipy Pillow pytesseract
```

Tesseract OCR engine is also required for log-axis calibration on raster figures:
- Windows: UB-Mannheim installer (`tesseract-ocr-w64-setup-*.exe`)
- Linux: `apt install tesseract-ocr`

`pymupdf`, `opencv-python`, and `Pillow` are optional — without them, graph sheets are skipped (`_HAS_GRAPH = False`). `pytesseract` is optional — without it, OCR-based axis calibration is skipped (`_HAS_OCR = False`).

## Architecture

The entire extractor is a single file (`mosfet_extractor.py`) with seven numbered sections:

### Section 1 — Graph-based rd extraction (lines ~107–420)
Pixel-level VF-IF curve digitisation using raster images. Finds the 25 °C forward-characteristics curve, fits a linear regression over the rated-current range, and returns `rd` (dynamic resistance).

### Section 1B — Temperature/energy/Zth/capacitance graph digitisation (lines ~422–2130)
The main graph engine. All four graph families share the same pipeline:

1. **`_find_plot_boxes`** — detects plot-frame rectangles from the PDF vector drawing layer (horizontal/vertical line groups). Handles side-by-side (2-up) Infineon layouts via `_split_columns` and 2×2 grid layouts (Nexperia) via `_split_rows`. After splitting, superset boxes (outer frames that contain smaller inner grid boxes) are removed so Y-axis clip positioning always targets the inner plot grid.
2. **`_figure_captions`** — extracts and joins "Figure N." caption lines. Delegates to `_infineon_diagram_captions` for Infineon's "Diagram N:" format.
3. **`_pair_caption_box`** — matches each caption to its plot box by vertical proximity. Scores two hypotheses (caption above vs below) and picks the winner.
4. **`_collect_curve_segments`** — pulls vector path segments from inside the plot box. Separates data curves from gridlines by stroke width.
5. **`_track_curves`** — column-sampled nearest-neighbour tracker with trajectory-extrapolation at crossings and fragment stitching. Returns `[({x_pdf: y_pdf}, coverage)]`.
6. **`_calibrate_axis`** — maps pixel coordinates to engineering values. Detection ladder (stops at first success):
   - Text-layer linear fit
   - OCR linear fit
   - Text-layer decade labels (`_decade_axis_labels`) — for `10^k` superscript notation
   - OCR decade labels (`_ocr_decade_axis_labels`) — for raster images where the text layer is empty
   - Strict log fit on exact powers-of-10 from text/OCR labels (`_try_log_fit`)
   - Wider Y-axis clip (100 px, then 160 px) — for Infineon two-up layouts where the shared Y-axis labels are far left of the right subbox
   - `_ocr_ok()` guard rejects contaminated OCR results for normalised-resistance Y-axes (requires ≥3 labels in [0, 5.5] and ≥50% majority)
   - User prompt (interactive mode only)
   - Normalised 0–1 fallback
7. **`extract_derating_curves`** — top-level dispatcher; calls the pipeline for every figure on every page, then calls `extract_raster_temp_graphs` as a fallback for datasheets with embedded-bitmap figures (e.g. some Infineon CoolSiC parts).

Figure classification (`_classify_fig`) routes each figure to one of: `energy`, `thermal_z`, `capacitance`, `power`, `current`, `resistance`, `relative`, `generic`.

### Section 1C — Word-baseline text scanner (lines ~2613–2876)
Fallback for narrow/bilingual/split-subscript parameter sheets where pdfplumber's table extractor misses rows. `baseline_lines` groups PDF word-level bounding boxes into lines; `scan_baseline_lines` then applies the same symbol/name matching as `scan_tables`.

### Section 2 — Parameter catalogue (lines ~2878–3262)
`PARAMS`: a list of dicts, one per tracked parameter (Ciss, RDS(on), Qg, Eon, …). Each dict carries:
- `symbol_res` — list of regex patterns matched against the Symbol column
- `name_res` — list of regex patterns matched against the Parameter Name column
- `unit`, `section`, `note`, `use_min`, `multi_cond`
- `valid_range` — `(min, max)` in base SI units; out-of-range values are discarded

Adding a new parameter: append a dict to `PARAMS` in the same format. The extraction loop in `extract()` iterates `PARAMS` automatically.

### Section 3 — Utilities + table parsing (lines ~3265–4242)
- `detect_structure` — infers column roles (sym, param_col, min, typ, max, unit, cond) from the first few rows of a pdfplumber table. Falls back to column-count heuristics for unlabelled tables.
- `scan_tables` — main matching loop: for each table × parameter, calls `row_matches` (symbol/name regex), then `row_values` (extracts typ/max, applies `to_base` unit normalisation, range-checks).
- `to_base` — universal unit normaliser: mΩ↔Ω, pF↔nF, nC↔µC, ns↔µs, µJ↔mJ, K/W↔°C/W, etc. All values are stored in base SI; `fmt_val` converts back to display units for the spreadsheet.
- `enrich_temps` — second-pass scan that fills in test temperatures missed by the primary scan (e.g. from stacked "Tvj = 25 °C / 150 °C" section headers).

### Section 4 — rd text/table fallback (lines ~4244–4380)
Five sequential strategies for extracting dynamic resistance when no graph is available: symbol match, name match, conductance inversion, VF-IF table scan, and formula reconstruction.

### Section 5 — Device info (lines ~4376–4428)
`get_device_info` — extracts part number and manufacturer from the first pages using ranked regex patterns for each supported manufacturer.

### Section 6 — Main extraction (lines ~4430–4585)
`extract(pdf_path)` orchestrates everything:
1. pdfplumber: read all text and tables
2. `get_device_info`
3. For each parameter in `PARAMS`: `scan_tables` → `scan_baseline_lines` → `scan_text` fallback
4. `extract_derating_curves` for graph sheets

### Section 7 — Excel output (lines ~4588–4855)
- `write_excel` — writes the "MOSFET Parameters" sheet. Row colour coding: green = found, yellow = not found, orange = max value, warm-yellow = high-temp row.
- `_write_graph_sheet` — generic writer for all four graph sheets (Energy, Temperature, Thermal Impedance, Capacitance). Each figure gets a header block, axis-calibration note, data table (max 15 rows, evenly sub-sampled if more), and the embedded original figure image.

## Key design decisions

**Log-axis calibration order matters.** `_try_log_fit` is intentionally strict: only values where `|log10(v) − round(log10(v))| < 0.05` are accepted. This prevents exponent digits (3, 4, 5 from "10^-3" etc.) from producing false 10^6 calibrations when OCR reads them out of context.

**Capacitance curve ordering.** After tracking, Ciss/Coss/Crss curves are sorted by descending median log₁₀(Y). Ciss is always largest at low VDS → assigned first. This overrides spatial-proximity label matching, which breaks when curves cross at high VDS.

**Vishay coloured Bézier curves.** Vishay capacitance curves are drawn as coloured (blue/orange/green) Bézier paths (`kind == "c"`), not black line segments. `_collect_curve_segments` rescues both `"l"` and `"c"` kind segments to capture these.

**Nexperia caption continuation guard.** `_figure_captions` skips text matching `r'^aaa[-_\s]?\d{4,}'` (NXP internal image reference codes). These codes sit in the 5 px inter-row gap between a caption and the adjacent figure, causing caption bleeding and wrong figure classification without this filter.

**Infineon CID font.** `_decode_cid_text` translates `(cid:XXXX)` placeholders in Infineon datasheets. This is applied to pdfplumber text but NOT to PyMuPDF text — the two libraries need to be kept consistent if new Infineon-specific code is added.

**`_HAS_GRAPH` / `_HAS_OCR` guards.** Every graph-related path checks these flags; all graph sheets are silently skipped when the optional dependencies are absent.
