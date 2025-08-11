# excel-to-ppt-explainer
Excel → PowerPoint generator that builds a summary slide and auto-links each number to a detail slide showing the exact Excel formula and a filtered raw-data snippet (xlwings + python-pptx).

Generate a PowerPoint from an Excel workbook:

- **Slide 1:** a _Summary Table_ (numbers are **clickable**).
- **Linked slides:** show the **exact Excel formula** used and a **filtered snippet** of the raw table that the formula referenced.

Powered by **xlwings** (reads real formulas from Excel) and **python-pptx** (builds the deck).

---

## Features

- Reads **live formulas** (incl. dynamic-array spills) directly from Excel via xlwings.
- Numbers on the summary slide link to per-metric **detail slides**.
- Each detail slide shows:
  - the original **formula** (as stored in Excel);
  - a **filtered** slice of the raw Excel **Table** (ListObject), restricted to only the columns referenced by that formula (plus a key column, which defaults to the first raw table column but can be set via `--key_header`).
- Hyperlinks can be created as **overlay shapes** on top of each cell (default) or as **text-run** hyperlinks.
- Overlay positioning uses the table’s **actual** widths/heights after text is placed; overlays are forced 100% transparent.
- Numeric values are rounded to `--round_digits` decimal places (default 2).

---

## Requirements

- Windows with **Microsoft Excel** installed (xlwings automates Excel).  
  > macOS with Excel generally works with xlwings too, but this project has been tested primarily on Windows.
- Python **3.10+**
- Install Python deps:
  ```bash
  pip install -r requirements.txt
  ```

`requirements.txt`:
```
xlwings>=0.30
python-pptx>=0.6.23
pandas>=2.0
```

---

## Quick start

```bash
python auto_generate_ppt_xlwings_final_v2.py ^
  --xlsx sample_sales_mix.xlsx ^
  --sheet Sheet1 ^
  --summary_start A12 ^
  --raw_table Raw_Data ^
  --key_header Product ^
  --out deck.pptx ^
  --link_mode overlay ^
  --table_font_pt 12 ^
  --round_digits 2 ^
  --verbose
```

**Arguments**

- `--xlsx` : path to the workbook.
- `--sheet` : worksheet containing the **raw table** and **summary**.
- `--summary_start` : the top-left **data** cell of the summary (first row below headers), e.g. `A12`.
- `--raw_table` : name of the Excel **Table** (ListObject) with raw data.
- `--key_header` : column used as the key in detail tables (e.g., `Product`). Defaults to the first column of the raw table.
- `--out` : output PPTX path.
- `--link_mode` : `text` (hyperlink on the number text; no shapes; **default**) or `overlay` (transparent rectangles on top of each numeric cell).
- `--table_font_pt` : font size for table text (deprecated alias: `--header_font_pt`). When `--link_mode` is `text`, row heights scale with this value (0.4" at 18pt).
- `--round_digits` : decimal places for numeric values (default 2).
- `--verbose` : print debug info (useful while wiring up a new workbook).

---

## How it works (high level)

1. **xlwings** opens the workbook and reads values and formulas (using `.Formula2`/`.Formula`).  
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
├─ auto_generate_ppt_xlwings_final_v2.py
├─ requirements.txt
├─ sample_sales_mix.xlsx        # optional demo workbook
├─ README.md
└─ LICENSE
```

---

## License

MIT — see [LICENSE](./LICENSE).
