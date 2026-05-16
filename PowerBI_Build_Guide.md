# Power BI Dashboard — Build Guide

## 1. Import data

**Option A (recommended):** Use prepared CSVs in this folder:

- `powerbi_void_loss_long.csv` — one row per building per month (fact table)
- `powerbi_building_ytd.csv` — building-level YTD aggregates
- `powerbi_org_monthly.csv` — organisation monthly totals
- `powerbi_specialism_ytd.csv` — specialism YTD totals

**Option B:** Import `Interview_Data_Worksheet.xlsx` and unpivot in Power Query (see section 6).

In Power BI Desktop: **Get data → Text/CSV** (or Excel) → load all four tables.

---

## 2. Data model (star schema)

```
DimDate (optional) ──< FactVoidLoss (powerbi_void_loss_long)
DimBuilding      ──<     │
DimSpecialism    ──<     │
DimManager       ──<     │
```

**Relationships:**

- `powerbi_void_loss_long[MonthStart]` → `DimDate[Date]` (or use MonthStart on axis directly)
- Create **Building** dimension from distinct `Building Ref` + `Building` + `Specialism` + `Manager` + `Void Loss Target`

Mark `powerbi_org_monthly` and `powerbi_specialism_ytd` as **dimension/helper** tables OR use only measures on the long table (preferred for consistency).

---

## 3. Core DAX measures

Create a **Measures** table and add:

```dax
// --- Organisation KPIs ---
YTD Void Loss £ = 
SUM ( 'powerbi_void_loss_long'[Void Loss] )

YTD Rent Charged £ = 
SUM ( 'powerbi_void_loss_long'[Rent Charged] )

YTD Void Loss % = 
DIVIDE ( [YTD Void Loss £], [YTD Rent Charged £] )

Org Target % = 0.065

YTD vs Target (pp) = 
[YTD Void Loss %] - [Org Target %]

Excess Void Loss £ vs Target = 
[YTD Void Loss £] - ( [Org Target %] * [YTD Rent Charged £] )

// --- In-month (respects month filter on visual) ---
In-Month Void Loss £ = [YTD Void Loss £]

In-Month Void Loss % = [YTD Void Loss %]

// --- Rolling trend ---
3M Rolling Void Loss % = 
VAR MaxDate = MAX ( 'powerbi_void_loss_long'[MonthStart] )
VAR WindowStart = EDATE ( MaxDate, -2 )
RETURN
    CALCULATE (
        DIVIDE ( SUM ( 'powerbi_void_loss_long'[Void Loss] ), SUM ( 'powerbi_void_loss_long'[Rent Charged] ) ),
        'powerbi_void_loss_long'[MonthStart] >= WindowStart,
        'powerbi_void_loss_long'[MonthStart] <= MaxDate
    )

// --- Specialism (with matrix / bar) ---
Specialism YTD % = [YTD Void Loss %]  // use in matrix with Specialism on rows

// --- Building vs target ---
Building Beat Target = 
VAR YTDPct = DIVIDE ( SUM ( 'powerbi_void_loss_long'[Void Loss] ), SUM ( 'powerbi_void_loss_long'[Rent Charged] ) )
VAR Tgt = MAX ( 'powerbi_void_loss_long'[Void Loss Target] )
RETURN IF ( YTDPct <= Tgt, 1, 0 )

Buildings Meeting Target = 
SUMX (
    VALUES ( 'powerbi_void_loss_long'[Building Ref] ),
    [Building Beat Target]
)
```

Format **YTD Void Loss %**, **In-Month Void Loss %**, and **3M Rolling Void Loss %** as **percentage, 2 decimal places**.

---

## 4. Suggested dashboard layout (one page)

```
┌─────────────────────────────────────────────────────────────────┐
│  VOID LOSS PERFORMANCE — YTD Apr 2025–Mar 2026                │
├──────────────┬──────────────┬──────────────┬──────────────────┤
│ YTD Void %   │ YTD Void £   │ vs Target    │ Buildings on     │
│  10.78%      │  £844k       │  +4.28 pp    │  target 12/42    │
│  (card, RAG) │  (card)      │  (card)      │  (card)          │
├──────────────┴──────────────┴──────────────┴──────────────────┤
│  LINE: In-Month Org Void Loss % by Month                        │
│  + constant reference line at 6.5%                              │
│  + optional 3M rolling line                                     │
├────────────────────────────┬────────────────────────────────────┤
│  BAR: YTD % by Specialism  │  BAR: Top 10 buildings by Void £   │
│  (sorted, target line 6.5%)│  (horizontal, click to filter)     │
├────────────────────────────┴────────────────────────────────────┤
│  TABLE: Building | Specialism | Manager | Target | YTD % | £     │
│  (conditional formatting on YTD % vs Target)                    │
└─────────────────────────────────────────────────────────────────┘
```

### Visual tips

- **KPI cards:** Use **red** if YTD Void Loss % > 6.5%, **amber** if 6.5–9%, **green** if ≤ 6.5%.  
- **Line chart:** `MonthStart` on X, `[In-Month Void Loss %]` on Y; add **Analytics → Constant line** at 0.065.  
- **Specialism bar:** `Specialism` on Y, `[YTD Void Loss %]` on X; data colours by distance from target.  
- **Slicers:** Specialism, Manager, Local Authority (for Q&A drill-down).  
- **Tooltip:** Show void loss £ and rent £ alongside %.  

---

## 5. Recommended 26/27 target line

Add a second constant line or card at **9.0%** labelled *“Proposed 2026/27 target”* to frame the recommendation visually.

---

## 6. Power Query unpivot (if using Excel only)

1. Load sheet → skip top rows until headers row.  
2. Select building metadata columns (A–F).  
3. Unpivot other columns.  
4. Split attribute column into **Month**, **Metric** (Void Loss / Rent / %) from the 3-column repeating pattern.  
5. Pivot Metric to columns.  
6. Set `Month` as date type.

---

## 7. Presentation mode

- Use **Bookmarks**: “Overview” | “Specialism” | “Buildings” | “Recommendation”.  
- **10-minute flow:** Start on Overview cards + line chart → click Specialism bookmark → top buildings → final slide with 9.0% target recommendation.

---

## 8. Files checklist for interview

- [ ] `.pbix` dashboard saved and tested in Present mode  
- [ ] `Key_Findings_and_Recommendations.md` printed or in speaker notes  
- [ ] Backup: PDF export of report page if Teams screen-share fails  
