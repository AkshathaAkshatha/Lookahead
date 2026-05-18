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

---

## Link to dashboard build

Once **`FactVoidLoss`** loads correctly, follow:

- **`Dashboard_From_Scratch.md`** — visuals and DAX  
- **`Key_Findings_and_Recommendations.md`** — narrative for submission  

---

## Interview talking point

> “The report connects directly to the Excel void loss worksheet. Power Query unpivots any new monthly columns on refresh, maps the date row to each month block, and the DAX KPIs recalculate without remodelling. We confirmed the February period with Performance when the column had been mislabelled as March.”
