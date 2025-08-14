# excel-to-ppt-explainer
Excel → PowerPoint generator that builds a summary slide and auto-links each number to a detail slide showing the exact Excel formula and a filtered raw-data snippet (openpyxl + python-pptx).

Generate a PowerPoint from an Excel workbook:

- **Slide 1:** a _Summary Table_ (numbers are **clickable**).
- **Linked slides:** show the **exact Excel formula** used and a **filtered snippet** of the raw table that the formula referenced.

Powered by **openpyxl** (reads stored formulas from Excel) and **python-pptx** (builds the deck).

---

## Features

- Reads **stored formulas** (incl. dynamic-array spills) directly from Excel via openpyxl.
- Numbers on the summary slide link to per-metric **detail slides**.
- Each detail slide shows:
  - the original **formula** (as stored in Excel);
  - a **filtered** slice of the raw Excel **Table** (ListObject), restricted to only the columns referenced by that formula (plus a key column, which defaults to the first raw table column but can be set via `--key_header`).
- Hyperlinks can be created as **overlay shapes** on top of each cell (default) or as **text-run** hyperlinks.
- Overlay positioning uses the table’s **actual** widths/heights after text is placed; overlays are forced 100% transparent.
- Numeric values are rounded to `--round_digits` decimal places (default 2).
- Can optionally append all generated slides to an **existing** PowerPoint file; when the template uses a wider slide size, navigation buttons and tables align with the title placeholders.

---

## Requirements

- Python **3.10+**
- Install Python deps:
  ```bash
  pip install -r requirements.txt
  ```

`requirements.txt`:
```
openpyxl>=3.1
python-pptx>=0.6.23
pandas>=2.0
```

---

## Quick start

```bash
python auto_generate_ppt_openpyxl.py ^
  --xlsx sample_sales_mix.xlsx ^
  --sheet Summary ^
  --summary_start A12 ^
  --key_header Product ^
  --pptx_in Power_point_input.pptx ^
  --out deck.pptx ^
  --link_mode overlay ^
  --table_font_pt 12 ^
  --round_digits 2 ^
  --slide_layout_idx 5 ^
  --skip_cols 2 4 ^
  --verbose
```

**Arguments**

- `--xlsx` : path to the workbook.
- `--sheet` : worksheet containing the **summary table** (raw tables may be on any sheet).
- `--summary_start` : the top-left **data** cell of the summary (first row below headers), e.g. `A12`.
- `--raw_table` : optional default Excel **Table** name. If omitted, the script auto-detects tables and each summary column may reference a different table.
- `--key_header` : column used as the key in detail tables (e.g., `Product`). Defaults to the first column of the raw table.
- `--pptx_in` : optional existing PowerPoint file to append the generated slides to.
- `--out` : output PPTX path. Defaults to the value of `--pptx_in` (if provided) or `deck.pptx`.
- `--link_mode` : `text` (hyperlink on the number text; no shapes; **default**) or `overlay` (transparent rectangles on top of each numeric cell).
- `--table_font_pt` : font size for table text (deprecated alias: `--header_font_pt`). When `--link_mode` is `text`, row heights scale with this value (0.4" at 18pt).
- `--round_digits` : decimal places for numeric values (default 2).
- `--slide_layout_idx` : index of the PowerPoint slide layout to use for generated slides (default 5).
- `--skip_cols` : one-based indexes of data columns (excluding the key column) to skip linking.
- `--verbose` : print debug info (useful while wiring up a new workbook).

## Summary formula guidelines

- Nested `SUM` and `FILTER` formulas are supported, e.g., `=SUM(FILTER(Raw_Data[Variants]*Raw_Data[Revenue per variant (USD)],Raw_Data[Category]=A12))`.
- Summary formulas must pull from data structured as an Excel Table (ListObject).
- Summary tables using `VLOOKUP` are untested and may not work.

---

## How it works (high level)

1. **openpyxl** opens the workbook and reads values and formulas.
2. The script locates your **Summary Table** based on `--summary_start` and parses each metric cell’s **formula anchor** (searches up/down the column if the cell is a spill result).
3. Structured references (like `Raw_Data[Revenue (USD)]`) are parsed to discover which **raw columns** are used by the formula.
4. For each summary cell:
   - A detail slide is created:
     - the **formula** is printed;
     - a **filtered** view of the raw table is shown (only the columns used by that formula + a configurable key column via `--key_header`).
   - On Slide 1, the numeric value is made **clickable** (overlay or text hyperlink) and jumps to the corresponding detail slide.

---

## Notes & tips

- Close the workbook in Excel before running (prevents file locks).
- If a summary cell was **pasted as a value** (no formula), the detail slide will show “(no formula found)”. The link still works; it just reflects what’s in the file.
- If your PowerPoint theme adjusts row heights/column widths, the script reads the **actual** sizes _after_ text is placed to keep overlays aligned.
- Overlay transparency is applied via python-pptx and also forced at the XML level; some PowerPoint UIs may still **display** “0% transparency”, but the rectangles will be visually invisible.

---

## Troubleshooting

- **Links don’t click (in edit mode).** Use **Slide Show** mode; PowerPoint sometimes ignores links in edit view.
- **“Repair” prompt on open.** Early versions used an action-based hyperlink; current code uses internal slide links and should not trigger repair.
- **Wrong rows in detail slide.** Ensure `--summary_start` points to the **first data row** (the row immediately below the summary headers).

---

## Example repo layout

```
.
├─ auto_generate_ppt_openpyxl.py
├─ requirements.txt
├─ sample_sales_mix.xlsx        # optional demo workbook
├─ README.md
└─ LICENSE
```

## Continuous integration

A GitHub Actions workflow installs dependencies and runs `verify_pptx.py` to ensure the generated deck has 21 slides and starts with a Summary Table slide.

---

## License

MIT — see [LICENSE](./LICENSE).
