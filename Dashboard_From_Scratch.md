# Void Loss Dashboard — Build From Scratch

**One table. One page. ~60–90 minutes.**

File to import: `powerbi_void_loss_long.csv`  
Do **not** import the other CSVs (avoids confusion).

---

## PHASE 1 — New file & load data (10 min)

1. Open **Power BI Desktop**.
2. **File → New** (blank report).
3. **Home → Get data → Text/CSV**.
4. Select `powerbi_void_loss_long.csv` → **Transform Data** (not Load).

### Power Query — fix column types

5. Click **`MonthStart`** column → **Date** (not Text).
6. Click **`Month`** column → **Date**.
7. **`Void Loss`** and **`Re,nt Charged`** → **Decimal Number**.
8. **`Void Loss Target`** → **Decimal Number**.
9. **Home → Close & Apply**.

### Rename table

10. **Model** view (left) → right-click table `powerbi_void_loss_long` → **Rename** → `FactVoidLoss`.

---

## PHASE 2 — Measures (15 min)

Click table **`FactVoidLoss`**. For **each** measure: **Modeling → New measure** → paste **one** formula → Enter.

```dax
YTD Void Loss £ = SUM ( FactVoidLoss[Void Loss] )
```

```dax
YTD Rent Charged £ = SUM ( FactVoidLoss[Rent Charged] )
```

```dax
YTD Void Loss % = DIVIDE ( [YTD Void Loss £], [YTD Rent Charged £] )
```

```dax
Org Target % = 0.065
```

```dax
YTD vs Target pp = [YTD Void Loss %] - [Org Target %]
```

```dax
In-Month Void Loss % = DIVIDE ( SUM ( FactVoidLoss[Void Loss] ), SUM ( FactVoidLoss[Rent Charged] ) )
```

### Format measures (click each measure → Measure tools)

| Measure | Format |
|---------|--------|
| YTD Void Loss £ | Currency £, 0 decimals |
| YTD Rent Charged £ | Currency £, 0 decimals |
| YTD Void Loss % | **Percentage**, 2 decimals |
| Org Target % | Percentage |
| YTD vs Target pp | Percentage, 2 decimals |
| In-Month Void Loss % | Percentage, 2 decimals |

**Test:** Card with `YTD Void Loss %` → **10.78%**. Card with `YTD Void Loss £` → **£844,114**.

---

## PHASE 3 — Page title (2 min)

1. **Insert → Text box**.
2. Type: **Void Loss Performance — YTD Apr 2025 to Feb 2026**.
3. Large font, bold, top of page.

---

## PHASE 4 — Four KPI cards (15 min)

For each: **Insert → Card** → drag **one** measure to **Fields**.

| Card label (add via Format → General → Title) | Field |
|-----------------------------------------------|-------|
| YTD Void Loss % | `YTD Void Loss %` |
| YTD Void Loss £ | `YTD Void Loss £` |
| vs Target | `YTD vs Target pp` |
| Org Target | `Org Target %` |

Arrange in a row under the title.

**Optional colour:** Select card → Format → Callout value → **Color** (red if YTD % card > 6.5%).

---

## PHASE 5 — Line chart: monthly trend (15 min)

1. **Insert → Line chart**.
2. **X-axis:** `MonthStart` (from FactVoidLoss).
3. **Y-axis:** `In-Month Void Loss %`.
4. Sort: click `MonthStart` in X-axis well → **⋯ → Sort axis → Ascending**.
5. **Format → X-axis → Type:** Continuous (if available).
6. **Analytics** (on visual) → **Constant line** → Value **0.065** → Label "Target 6.5%".

**Title:** In-month void loss % (organisation)

---

## PHASE 6 — Bar chart: specialism (10 min)

1. **Insert → Clustered bar chart**.
2. **Y-axis:** `Specialism`.
3. **X-axis:** `YTD Void Loss %`.
4. Sort bars by X-axis descending (worst at top).
5. **Analytics → Constant line** at 0.065 (target).

**Title:** YTD void loss % by specialism

---

## PHASE 7 — Bar chart: top buildings by £ (10 min)

1. **Insert → Clustered bar chart**.
2. **Y-axis:** `Building`.
3. **X-axis:** `YTD Void Loss £` (use the £ measure, not %).
4. **Filters on this visual → Building → Filter type: Top N → Top 10 → By value: YTD Void Loss £**.

**Title:** Top 10 buildings by void loss (£)

---

## PHASE 8 — Table (10 min)

1. **Insert → Table**.
2. Add columns: `Building`, `Specialism`, `Manager`, `Void Loss Target`, then measures `YTD Void Loss %`, `YTD Void Loss £`.
3. **Format → Cell elements → Values → YTD Void Loss %** → **Conditional formatting** → Background colour → Rules vs `Org Target %` or manual 6.5%.

**Title:** Building detail (YTD)

---

## PHASE 9 — Slicers (5 min)

1. **Insert → Slicer** → `Specialism`.
2. Another slicer → `Manager`.

Place on the right side. Selecting a value filters **all** visuals.

---

## PHASE 10 — Save & present (5 min)

1. **File → Save as** → `Lookahead_Void_Loss.pbix`.
2. **View → Reading view** or **F5** to test.
3. **View → Bookmarks** (optional) for presentation sections.

---

## Expected numbers (sanity check)

| Item | Value |
|------|--------|
| YTD Void Loss % | 10.78% |
| YTD Void Loss £ | £844,114 |
| Org Target | 6.50% |
| vs Target | +4.28 pp |
| Mental Health (slicer) | ~9.03% |
| Homelessness (slicer) | ~12.46% |

---

## 10-minute presentation order

1. Title + four cards (org KPI vs target).
2. Line chart (improving or not month to month).
3. Specialism bar (who drives the problem).
4. Top 10 buildings (where to act).
5. Recommend **9.0%** target for 2026/27 (say aloud; optional extra card).

