# Power Query Only — Load & Unpivot from Excel (Refresh-Friendly)

**Source file:** `Interview_Data_Worksheet.xlsx`  
**Goal:** Build the fact table once in Power BI; when Look Ahead adds **March** (or any new month as new columns), click **Refresh** — no manual rework.

**Worksheet layout (do not change this logic):**

| Row (Excel) | Contents |
|-------------|----------|
| 1 | Empty / title |
| 2 | **Month-end dates** in columns G, J, M… (every **3** columns from column G) |
| 3 | **Headers:** Building Ref … Void Loss Target, then repeating **Void Loss / Rent / %** |
| 4+ | **One row per building** |

---

## Before you start

1. Install **Power BI Desktop**.
2. Put `Interview_Data_Worksheet.xlsx` in a stable folder (e.g. your Lookahead project folder).
3. Optional: in Excel, select the data block → **Insert → Table** → name it `VoidLossRaw` (makes the source range stable). If you skip this, Power BI still loads the sheet.

---

## Overview — two queries

| Query | Purpose |
|-------|---------|
| **MonthMap** | Reads the **date row** (Excel row 2) → one row per month |
| **FactVoidLoss** | Reads building rows, **unpivots** all month columns, attaches the correct month |

On refresh, new month columns are included automatically via **Unpivot Other Columns**.

---

## PART 1 — Create the report and connect to Excel

1. **File → New** in Power BI Desktop.
2. **Home → Get data → Excel**.
3. Select `Interview_Data_Worksheet.xlsx` → **Open**.
4. Tick the correct **sheet** (usually `Sheet1`) → **Transform Data** (not Load).

You are now in **Power Query Editor**.

---

## PART 2 — Query `MonthMap` (month dates)

### Step 1 — Duplicate the sheet query

1. In **Queries** pane (left), right-click the sheet query → **Duplicate**.
2. Rename the duplicate to **`MonthMap`**.
3. Rename the original (you’ll edit next) to **`FactVoidLoss`** later — or duplicate again from MonthMap when done.

### Step 2 — Keep only the date row

1. With **`MonthMap`** selected:
2. **Home → Remove Rows → Remove Top Rows** → **1**  
   (Now Excel row 2 is the first row — the dates row.)
3. **Home → Remove Bottom Rows** → enter a large number so only **one row** remains  
   **Or:** **Keep Rows → Keep Top Rows** → **1**.

### Step 3 — Unpivot the date cells (columns from G onward)

1. Click the **first metadata column** (column A / `Column1`).
2. **Shift+click** the **last metadata column** (column F / Void Loss Target) — selects A–F.
3. **Transform → Unpivot Columns → Unpivot Other Columns**.
4. You now have something like `Attribute` (old column names) and `Value` (dates).

### Step 4 — Keep only month blocks (every 3rd unpivoted column from the start of months)

The dates appear on the **first column of each 3-column month block** (Void Loss column), not on Rent or %.

1. **Add Column → Index Column → From 0**.
2. **Add Column → Custom Column**  
   - Name: `IsMonthStart`  
   - Formula: `Number.Mod([Index], 3) = 0`
3. **Filter** `IsMonthStart` = **true**.
4. Remove `Index` and `IsMonthStart` if you like.

### Step 5 — Clean `MonthMap`

1. Rename `Value` → **`Month`**.
2. Remove `Attribute`.
3. **Add Column → Index Column → From 0** → rename to **`MonthBlockIndex`** (0 = Apr 2025, 1 = May 2025, …).
4. Set **`Month`** type to **Date**.
5. **Fix February labelling** (David confirmed last period is Feb 2026, not Mar):

   - **Transform → Replace Values** on `Month`  
   - Find: `31/03/2026` or `2026-03-31` (match what you see)  
   - Replace: `28/02/2026`

6. **Add Column → Custom Column**  
   - Name: `MonthStart`  
   - Formula: `Date.StartOfMonth([Month])`
7. **Home → Close & Apply** is NOT yet — keep editor open.  
8. **Disable load** for now: right-click **`MonthMap`** → **Enable load** unchecked *(optional; you can load it for debugging)*.

---

## PART 3 — Query `FactVoidLoss` (buildings + unpivot)

Select query **`FactVoidLoss`** (the main sheet query, not MonthMap).

### Step 1 — Remove header/date rows; promote column names

1. **Remove Top Rows** → **2**  
   (Removes Excel rows 1–2; old row 3 becomes row 1.)
2. **Home → Use First Row as Headers** (Promote Headers).
3. You should see: `Building Ref`, `Building`, `Specialism`, … `Monthly Void Loss`, `Monthly Rent Charged`, …

Power BI may rename duplicates as:

- `Monthly Void Loss`, `Monthly Rent Charged`, `Monthly Void Loss %`
- `Monthly Void Loss.1`, `Monthly Rent Charged.1`, …

That **`.1`, `.2` suffix** = month block index (important later).

### Step 2 — Set types on metadata columns

1. Set **`Building Ref`** → Whole number (or text if mixed).
2. **`Void Loss Target`** → Decimal number.
3. Leave month columns as text/any for now.

### Step 3 — Remove blank building rows

1. Filter **`Building Ref`** → uncheck **null** / blank.

### Step 4 — Unpivot all month columns (refresh-friendly)

1. Select columns **`Building Ref`** through **`Void Loss Target`** (first **6** columns only).
2. **Transform → Unpivot Columns → Unpivot Other Columns**.
3. Columns: **`Attribute`**, **`Value`**.

When **March 2027** is added in Excel as 3 new columns to the right, refresh will unpivot them too — **no query edit**.

### Step 5 — Derive `MonthBlockIndex` from attribute name

1. **Add Column → Custom Column**  
   - Name: `MonthBlockIndex`  
   - Formula:

```powerquery
let
    s = [Attribute],
    dot = Text.PositionOf(s, ".", Occurrence.First)
in
    if dot = -1 then 0 else Number.FromText(Text.End(s, Text.Length(s) - dot - 1))
```

- `Monthly Void Loss` → **0** (April)  
- `Monthly Void Loss.1` → **1** (May)  
- etc.

### Step 6 — Derive metric type (Void / Rent / %)

1. **Add Column → Custom Column**  
   - Name: `Metric`  
   - Formula:

```powerquery
if Text.Contains([Attribute], "Rent") then "Rent"
else if Text.Contains([Attribute], "%") then "Pct"
else "Void"
```

### Step 7 — Merge month dates from `MonthMap`

1. **Home → Merge Queries**  
   - Left: **`FactVoidLoss`**  
   - Right: **`MonthMap`**  
   - Column: **`MonthBlockIndex`** = **`MonthBlockIndex`**  
   - Join kind: **Left Outer**  
   - OK  
2. Expand **`MonthMap`** — tick **`Month`**, **`MonthStart`** only.

### Step 8 — Pivot metrics to columns

1. **Transform → Pivot Column**  
   - Values column: **`Value`**  
   - Column: **`Metric`**  
   - Advanced: **Don’t aggregate** (you already have one value per cell).
2. Rename pivoted columns:
   - `Void` → **`Void Loss`**
   - `Rent` → **`Rent Charged`**
   - `Pct` → **`Void Loss Pct`**
3. Set types: **Void Loss** and **Rent Charged** → Decimal; **Void Loss Pct** → Decimal.

### Step 9 — Final column tidy-up

Keep (and rename if needed):

| Column | Use |
|--------|-----|
| Building Ref, Building, Specialism, Manager, Local Authority, Void Loss Target | Dimensions |
| MonthStart, Month | Time |
| Void Loss, Rent Charged, Void Loss Pct | Facts |

Remove: `Attribute`, `MonthBlockIndex`, `Metric`, duplicate columns.

### Step 10 — Apply

1. Rename query to **`FactVoidLoss`**.
2. **Home → Close & Apply**.

---

## PART 4 — Validate after load

In **Report** view, create a quick table:

| Check | Expected |
|-------|----------|
| Row count | **462** (42 buildings × 11 months) |
| Distinct MonthStart | **11** months (Apr 2025 – Feb 2026) |
| Sum Void Loss | **~£844,114** |
| YTD % measure | **~10.78%** |

**DAX (after you create measures):**

```dax
YTD Void Loss £ = SUM ( FactVoidLoss[Void Loss] )
YTD Rent Charged £ = SUM ( FactVoidLoss[Rent Charged] )
YTD Void Loss % = DIVIDE ( [YTD Void Loss £], [YTD Rent Charged £] )
```

---

## PART 5 — When March (or a new month) is added in Excel

Look Ahead adds **3 columns** to the right of the sheet:

`Monthly Void Loss | Monthly Rent Charged | Monthly Void Loss %`

with the correct date in **row 2** above the void column.

**You do:**

1. Save the updated Excel file (same path, or update a **Parameter** path).
2. Power BI → **Home → Refresh**.
3. Power Query re-runs **MonthMap** + **FactVoidLoss** unpivot.
4. Dashboard updates (new point on line chart; YTD includes March).

**You do NOT:**

- Re-select column lists manually  
- Rebuild the report  
- Change DAX (unless you add a static calendar table with a fixed end date)

---

## PART 6 — Optional: parameter for file path

1. **Home → Manage Parameters → New Parameter**  
   - Name: `ExcelFilePath`  
   - Type: Text  
   - Current value: full path to `Interview_Data_Worksheet.xlsx`
2. In **Source** step of both queries, use:

```powerquery
Excel.Workbook(Parameter.ExcelFilePath, null, true)
```

Then only update the parameter when the file moves.

---

## PART 7 — Optional: static `DimDate` (replace fixed end date)

If you create `DimDate` with `CALENDAR ( DATE(2025,4,1), DATE(2026,2,28) )`, extend the end when new months arrive:

```dax
DimDate =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2025, 4, 1 ), EOMONTH ( MAX ( FactVoidLoss[MonthStart] ), 0 ) ),
    "Month Name", FORMAT ( [Date], "MMM yyyy" ),
    "Month Sort", YEAR ( [Date] ) * 100 + MONTH ( [Date] )
)
```

Relationship: `DimDate[Date]` → `FactVoidLoss[MonthStart]` (many-to-one).

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Wrong row count | Check **Remove Top Rows** = 2 before promote headers |
| Months all null after merge | `MonthBlockIndex` formula must match `.1`, `.2` suffixes |
| February shows as March | **Replace Values** in `MonthMap` (see Part 2 Step 5) |
| Refresh breaks | New month must follow **3-column** pattern; date in row 2 above void column |
| `Pct` column text | Change type to Decimal; `%` may be stored as 0.11 not 11 |
| Duplicate buildings | Filter null `Building Ref` |

---

## Advanced: single-query M template (copy into Advanced Editor)

Use only if you’re comfortable with M. Creates **`FactVoidLoss`** in one query (includes MonthMap logic inline). Update the file path.

```powerquery
let
    FilePath = "C:\FULL\PATH\TO\Interview_Data_Worksheet.xlsx",
    Source = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],

    // --- Month map from row 2 (index 1) ---
    DateRow = Sheet{1},
    DateRowTable = Table.FromList(Record.FieldValues(DateRow), Splitter.SplitByNothing(), {"Value"}),
    DateRowWithIndex = Table.AddIndexColumn(DateRowTable, "ColIndex", 0, 1),
    MonthDates = Table.SelectRows(DateRowWithIndex, each [ColIndex] >= 6 and Number.Mod([ColIndex] - 6, 3) = 0),
    MonthMap = Table.AddIndexColumn(
        Table.TransformColumnTypes(
            Table.RenameColumns(Table.RemoveColumns(MonthDates, {"ColIndex"}), {{"Value", "Month"}}),
            {{"Month", type date}}
        ),
        "MonthBlockIndex", 0, 1
    ),
    MonthMapFixed = Table.ReplaceValue(MonthMap, #date(2026,3,31), #date(2026,2,28), Replacer.ReplaceValue, {"Month"}),
    MonthMapFinal = Table.AddColumn(MonthMapFixed, "MonthStart", each Date.StartOfMonth([Month]), type date),

    // --- Building data from row 4+ ---
    DataRows = Table.Skip(Sheet, 2),
    Promoted = Table.PromoteHeaders(DataRows, [PromoteAllScalars=true]),
    MetaCols = List.FirstN(Table.ColumnNames(Promoted), 6),
    Unpivoted = Table.UnpivotOtherColumns(Promoted, MetaCols, "Attribute", "Value"),

    AddBlockIndex = Table.AddColumn(Unpivoted, "MonthBlockIndex", each
        let s = [Attribute], dot = Text.PositionOf(s, ".", Occurrence.First)
        in if dot = -1 then 0 else Number.FromText(Text.End(s, Text.Length(s) - dot - 1))
    ),
    AddMetric = Table.AddColumn(AddBlockIndex, "Metric", each
        if Text.Contains([Attribute], "Rent") then "Rent"
        else if Text.Contains([Attribute], "%") then "Pct"
        else "Void"
    ),

    Merged = Table.NestedJoin(AddMetric, {"MonthBlockIndex"}, MonthMapFinal, {"MonthBlockIndex"}, "MonthMap", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "MonthMap", {"Month", "MonthStart"}, {"Month", "MonthStart"}),

    Pivoted = Table.Pivot(Expanded, List.Distinct(Expanded[Metric]), "Metric", "Value"),
    Renamed = Table.RenameColumns(Pivoted, {{"Void", "Void Loss"}, {"Rent", "Rent Charged"}, {"Pct", "Void Loss Pct"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Void Loss", type number}, {"Rent Charged", type number}, {"Void Loss Pct", type number},
        {"Void Loss Target", type number}, {"MonthStart", type date}, {"Month", type date}
    }),
    Filtered = Table.SelectRows(Typed, each [Building Ref] <> null and [Building Ref] <> "")
in
    Filtered
```

After pasting: fix `FilePath`, sheet name, and February replace date if needed → **Close & Apply**.

In the single-query version, replace hardcoded `FilePath` with `Parameter.ExcelFilePath` once parameters are created (Part 8).

---

## PART 8 — Parameters, template & new financial year (don’t start again)

### Choose your approach

| Approach | Best for | Year comparison |
|----------|----------|-----------------|
| **A — One `.pbix` per year** (Steps below) | Archiving annual packs | Open two files side by side |
| **B — One `.pbix` + Financial Year slicer** (**Part 9**) | **Same dashboard, switch & compare years** | **Slicer + optional YoY line chart** |

**You preferred a year slicer** → build **Part 9** (add `FinancialYear` column, append files, slicer on report). Keep Part 8 for parameters and monthly refresh.

### The two refresh scenarios

| When | What changes | What you do |
|------|----------------|-------------|
| **New month, same year** | 3 columns added to the **same** Excel | **Refresh** — same `.pbix`, queries unchanged |
| **New financial year (e.g. 2026/27)** | **New Excel file** in folder | **Approach A:** new `.pbix` from template · **Approach B:** drop file in folder → **Refresh** (Part 9) |

Power Query **steps stay the same** if the worksheet **layout** is the same (cols A–F + 3 columns per month).

---

### Step 1 — Create parameters (once)

In Power BI Desktop: **Home → Manage Parameters → New parameter**

Create these three:

| Name | Type | Suggested current value | Used for |
|------|------|-------------------------|----------|
| `ExcelFilePath` | Text | Full path to `Interview_Data_Worksheet.xlsx` | Source file |
| `FinancialYearLabel` | Text | `2025/26` | Report title / text boxes |
| `OrgTargetPct` | Decimal | `0.065` | DAX target measure & reference lines |

**Important:** For `ExcelFilePath`, click **Current Value** and browse to your file, or paste:

`/Users/kwiffstaff/Downloads/Aks-CV/Look_ahead/Lookahead/Interview_Data_Worksheet.xlsx`

---

### Step 2 — Wire parameters into Power Query

1. Open **Transform data**.
2. Select **`FactVoidLoss`** (and **`MonthMap`** if separate) → **Advanced Editor** (or edit **Source** step).
3. Replace hardcoded path with the parameter.

**Source step should look like:**

```powerquery
Source = Excel.Workbook(Parameter.ExcelFilePath, null, true)
```

4. **Done** on each query that loads the file.

**Optional:** Add a column in Power Query:

```powerquery
FinancialYear = Parameter.FinancialYearLabel
```

Useful if you later combine multiple years in one model (advanced).

---

### Step 3 — Wire `OrgTargetPct` into DAX (instead of hardcoded 0.065)

Replace:

```dax
Org Target % = 0.065
```

With:

```dax
Org Target % = SELECTEDVALUE ( OrgTarget[Target], 0.065 )
```

**Simple approach for interview (no extra table):**

```dax
Org Target % = [Org Target % Value]
```

Create a **What-if parameter** or a one-row table — **easiest:**

1. **Modeling → New measure**:

```dax
Org Target % = 0.065
```

2. Each year, edit the measure in the **2026/27** copy of the file to `0.09`, **or** use a parameter:

```dax
Org Target % = SELECTEDVALUE ( 'Target Parameter'[Org Target % Parameter], 0.065 )
```

(Create via **Modeling → New parameter** → Decimal → default `0.065`, name it `Org Target % Parameter`.)

**Constant lines** on charts: use the measure `Org Target %` instead of typing `0.065`.

---

### Step 4 — Use `FinancialYearLabel` on the report

1. **Insert → Text box** on the dashboard.
2. Type: `Void Loss Performance — YTD ` then insert the parameter:
   - **Insert → Card** with a measure, **or** manually type `2025/26` and change it when you copy the file.
3. **Practical approach:** title text =  
   `"Void Loss Performance — " & [FinancialYearLabel]`  
   requires a measure:

```dax
Report Title = "Void Loss Performance — YTD " & SELECTEDVALUE ( FY[Label], "2025/26" )
```

**Simplest for submission:** hardcode title `2025/26`; change text when you copy to next year.

---

### Step 5 — Save as template (.pbit) — reuse without rebuilding

1. Finish the 2025/26 report (queries + dashboard).
2. **File → Export → Power BI template**.
3. Save as `Lookahead_Void_Loss_Template.pbit`.

**Next year:**

1. Double-click `Lookahead_Void_Loss_Template.pbit`.
2. Power BI prompts for parameter values → set:
   - `ExcelFilePath` → `Interview_Data_Worksheet_2026-27.xlsx`
   - `FinancialYearLabel` → `2026/27`
   - `OrgTargetPct` → `0.09` (or whatever SMT agrees)
3. **File → Save As** → `Lookahead_Void_Loss_2026-27.pbix`.
4. **Refresh** → check KPIs and titles.

Your **2025/26 file stays untouched** in the folder as the archived pack.

---

### Step 6 — Recommended folder structure

```
Lookahead/
├── Interview_Data_Worksheet_2025-26.xlsx    ← current year source
├── Interview_Data_Worksheet_2026-27.xlsx    ← next year (when issued)
├── Lookahead_Void_Loss_2025-26.pbix         ← frozen after year-end
├── Lookahead_Void_Loss_2026-27.pbix         ← active year
├── Lookahead_Void_Loss_Template.pbit        ← master template
└── Archive/
    └── Lookahead_Void_Loss_2025-26.pdf      ← optional SMT snapshot
```

---

### Step 7 — Monthly workflow (same financial year)

1. Performance team publishes updated Excel (new month columns).
2. Save over the same file **or** save as `..._2025-26_May.xlsx` and update `ExcelFilePath` once.
3. Open **`Lookahead_Void_Loss_2025-26.pbix`** → **Refresh**.
4. Spot-check: row count grew by ~42 rows per month; YTD % moved slightly.

**No Power Query edits** if layout unchanged.

---

### Step 8 — Year-end workflow

1. **Final refresh** after last month of FY.
2. **Export PDF** of report page for records.
3. **Save** final `Lookahead_Void_Loss_2025-26.pbix` → move copy to `Archive/` (read-only).
4. **Do not** keep refreshing the archived file next year.

---

### Step 9 — Optional: compare two years in one report (advanced)

Only if SMT wants **side-by-side** 25/26 vs 26/27 in **one** app:

1. Put both Excel files in a folder.
2. Power Query: **Folder** source → combine files → add `FinancialYear` column from filename.
3. Append into one `FactVoidLoss` with a `FinancialYear` column.
4. Report: **Slicer** on `FinancialYear` or separate pages per year.

For the interview, **one pbix per year** is simpler and easier to explain.

---

### What breaks automation (when you *would* redo Power Query)

| Change | Action |
|--------|--------|
| New month columns (same layout) | Refresh only |
| New file, same layout | New parameter path or new pbix from template |
| Columns A–F renamed or moved | Update promoted headers / MetaCols list |
| Month blocks not in groups of 3 | Redesign unpivot logic |
| Data moves from Excel to database | Replace source with SQL; keep measure logic |

---

### Interview talking points (year versioning — Approach A)

> “For monthly reporting within the year, I’d refresh the same Power BI file — the unpivot picks up new columns automatically. At year-end I’d archive the 25/26 report and spin up a new 26/27 file from a template, updating the Excel path and organisational target parameter, so we never overwrite last year’s dashboard. The transformation logic is reusable unless the source layout changes.”

---

## PART 9 — Financial Year slicer (one report, year-on-year comparison)

**Goal:** One dashboard. Slicer selects **2025/26** or **2026/27**. KPIs and charts filter to that year. Optional chart compares **same calendar month** across years.

**Today:** You only have one Excel file — add `FinancialYear = "2025/26"` now. When **2026/27** file arrives, append it — no new dashboard.

---

### Step 1 — Folder for all year files

Create a folder, e.g.:

```
Lookahead/Data/
├── Interview_Data_Worksheet_2025-26.xlsx
└── (later) Interview_Data_Worksheet_2026-27.xlsx
```

Same sheet layout in every file.

---

### Step 2 — Turn single-file query into a function (Power Query)

1. Open **Transform data**.
2. Select your working **`FactVoidLoss`** query (after unpivot logic works on one file).
3. **Right-click → Create Function** (or **Advanced** → create function `TransformVoidLossFile`).
4. Name it **`TransformVoidLossFile`** with parameter **`FilePath`** (text).

Inside the function: paste your full transformation (from Excel load through to typed columns). Last step must return the table.

**Test:** Invoke with your 2025/26 path → same 462 rows as before.

*If **Create Function** is greyed out, duplicate the query, rename to `TransformVoidLossFile`, add `(FilePath as text)` in Advanced Editor as first line after `let`.*

---

### Step 3 — Load folder and append all years

1. **New Source → Folder** → point to `Lookahead/Data/`.
2. **Combine → Combine & Transform** (or **Edit** then filter `.xlsx` only).
3. For each file, call the function:

**Pattern in Power Query:**

```powerquery
let
    Source = Folder.Files("C:\...\Lookahead\Data"),
    XlsxOnly = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsx")),
    AddData = Table.AddColumn(XlsxOnly, "Data", each TransformVoidLossFile([Folder Path] & [Name])),
    Combined = Table.Combine(AddData[Data])
in
    Combined
```

4. **Add `FinancialYear` from filename** (if not already inside function):

```powerquery
= Table.AddColumn(Combined, "FinancialYear", each
    if Text.Contains([Name], "2026-27") then "2026/27"
    else if Text.Contains([Name], "2025-26") then "2025/26"
    else "Unknown"
)
```

*Better:* derive from filename with `Text.BetweenDelimiters([Name], "_", ".xlsx")` → `2025-26` → replace `-` with `/`.

5. Rename final query **`FactVoidLoss`** → **Close & Apply**.

**Refresh later:** Drop `Interview_Data_Worksheet_2026-27.xlsx` in folder → **Refresh** → second year appears in slicer.

---

### Step 4 — Table of targets per year (for correct target line)

**Enter data** (Home → Enter data):

| FinancialYear | OrgTargetPct |
|---------------|--------------|
| 2025/26       | 0.065        |
| 2026/27       | 0.09         |

Name table **`DimFinancialYear`**.

**Model view:** Relate `DimFinancialYear[FinancialYear]` → `FactVoidLoss[FinancialYear]` (one-to-many).

**Measures:**

```dax
Org Target % =
SELECTEDVALUE ( DimFinancialYear[OrgTargetPct], 0.065 )
```

When the slicer selects **2026/27**, target line becomes **9%** automatically.

---

### Step 5 — Add Fiscal Month for YoY line charts

Months must align across years (April = month 1 of FY, etc.).

**Option A — Custom column in Power Query:**

```powerquery
FiscalMonthNumber = Date.Month([MonthStart])
FiscalMonthName = Date.ToText([MonthStart], "MMM")
```

(Apr–Mar order works if all years start April; sort by `MonthStart` within year.)

**Option B — DAX calculated column** (if FY always starts April):

```dax
Fiscal Month Sort =
VAR M = MONTH ( FactVoidLoss[MonthStart] )
RETURN IF ( M >= 4, M, M + 12 )
```

---

### Step 6 — Financial Year slicer on the report

1. **Insert → Slicer**.
2. Field: **`FinancialYear`** (from `FactVoidLoss` or `DimFinancialYear`).
3. Style: **Dropdown**; turn on **Select all** if you want multi-year views.
4. Place at top right (above Specialism slicer).

**Behaviour:**

- Select **2025/26** → all cards/charts show that year’s YTD and months.
- Select **2026/27** when file exists → switches dataset.
- **Ctrl+click** two years → filters to both (useful for combined tables; line chart needs Legend — below).

Set **default:** **Edit interactions** / bookmark, or slicer default filter **2025/26** only.

---

### Step 7 — Update DAX measures (respect slicer automatically)

These already filter by slicer if `FinancialYear` is on the fact table:

```dax
YTD Void Loss £ = SUM ( FactVoidLoss[Void Loss] )
YTD Rent Charged £ = SUM ( FactVoidLoss[Rent Charged] )
YTD Void Loss % = DIVIDE ( [YTD Void Loss £], [YTD Rent Charged £] )
In-Month Void Loss % = DIVIDE ( SUM ( FactVoidLoss[Void Loss] ), SUM ( FactVoidLoss[Rent Charged] ) )
```

No change needed — **slicer filters fact rows**.

**Optional — show selected year in title:**

```dax
Selected FY Label =
"Void Loss Performance — YTD " & SELECTEDVALUE ( FactVoidLoss[FinancialYear], "All years" )
```

Use in a **Card** or text box (needs Card visual with measure, or manual title updated each demo).

---

### Step 8 — Year-on-year comparison line chart

**Chart type:** Line chart

| Well | Field |
|------|--------|
| **X-axis** | `MonthStart` **or** `FiscalMonthName` (sort by month number) |
| **Y-axis** | `In-Month Void Loss %` |
| **Legend** | **`FinancialYear`** |
| **Slicer** | Set **Financial Year** to **multi-select** both years, **or** clear slicer and use Legend only |

**How to read it:** Two lines — **2025/26** vs **2026/27** — same months on X-axis (Apr, May, …).

**Tip:** If X-axis duplicates (Apr 2025 and Apr 2026 as separate labels), use **FiscalMonthName** on X-axis and **FinancialYear** on Legend, with **Fiscal Month Sort** as sort column.

---

### Step 9 — Optional YoY comparison card

```dax
YTD Void Loss % (Prior Year) =
CALCULATE (
    [YTD Void Loss %],
    SAMEPERIODLASTYEAR ( FactVoidLoss[MonthStart] )
)
```

*Only works if both years exist in the model and dates align — for different FY files, simpler approach:*

```dax
YTD Void Loss % 2025-26 =
CALCULATE ( [YTD Void Loss %], FactVoidLoss[FinancialYear] = "2025/26" )

YTD Void Loss % 2026-27 =
CALCULATE ( [YTD Void Loss %], FactVoidLoss[FinancialYear] = "2026/27" )

YoY Change pp =
[YTD Void Loss % 2026-27] - [YTD Void Loss % 2025-26]
```

Show as cards when both years loaded.

---

### Step 10 — What you do each year (slicer approach)

| Event | Action |
|-------|--------|
| **New month (2025/26)** | Update that year’s Excel → **Refresh** |
| **New year (2026/27)** | Add `Interview_Data_Worksheet_2026-27.xlsx` to **Data** folder → **Refresh** |
| **New org target** | Add row to **`DimFinancialYear`** table |
| **Dashboard** | **Same `.pbix`** — slicer gains new year |

**Do not overwrite** old Excel files — keep them in the folder for history.

---

### Step 11 — For your interview submission (one file only)

You only have 2025/26 data today:

1. Add column in Power Query: **`FinancialYear = "2025/26"`** (custom column).
2. Add **`DimFinancialYear`** with one row: `2025/26`, `0.065`.
3. Build **Financial Year slicer** (one value for now).
4. Say in presentation: *“When 26/27 data is added to the folder, refresh adds a second slicer value and enables YoY on the legend chart.”*

---

### Interview talking point (year slicer)

> “I’ve modelled financial year on the fact table so SMT can slice one dashboard by year or compare years on the same trend chart. New annual files go into a standard folder; Power Query appends them on refresh without rebuilding the report. Targets sit in a small dimension table so the 6.5% and 9% lines follow the selected year.”

---

## Link to dashboard build

Once **`FactVoidLoss`** loads correctly, follow:

- **`Dashboard_From_Scratch.md`** — visuals and DAX  
- **`Key_Findings_and_Recommendations.md`** — narrative for submission  

---

## Interview talking points

**Refresh & data quality**

> “The report connects directly to the Excel void loss worksheet. Power Query unpivots any new monthly columns on refresh, maps the date row to each month block, and the DAX KPIs recalculate without remodelling. We confirmed the February period with Performance when the column had been mislabelled as March.”

**New financial year (slicer / single report)**

> “Financial year is a slicer on one report — new annual workbooks drop into a folder and append on refresh, so we can compare 25/26 and 26/27 on the same visuals without overwriting history. Targets are year-specific via a small dimension table.”
