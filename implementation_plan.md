# AppSheet-Powered Mobile Pharmacy Product Evaluation System

Build a complete Google Apps Script backend engine + AppSheet configuration guide for a 3-evaluator hospital pharmacy product evaluation workflow.

---

## Proposed Changes

### Component 1: Apps Script Backend Engine

#### [NEW] [`Code.gs`](file:///c:/Users/PC 127/Desktop/PPMP Project/Code.gs)

The single, comprehensive Apps Script file containing all backend logic. Functions:

| Function | Purpose |
|---|---|
| `initializeEvaluationWorkbook()` | Creates & formats all 6 sheets with headers, styling, validations, frozen rows/columns |
| `syncMasterDrugList()` | Pulls generic drug names from the PPMP Master List (`indicative 1` tab, Column A from row 4) into `Medicine_Master` |
| `onEvaluationsSubmitTrigger(e)` | `onEdit` trigger that mirrors new/updated rows from `Evaluations_Master` → individual evaluator tabs, then recalculates `Consolidated_Summary` |
| `refreshConsolidatedSummary()` | Rebuilds the pivot/consensus dashboard comparing all 3 evaluators per drug/brand |
| `setupTriggers()` | Installs the `onEdit` trigger programmatically |
| `createCustomMenu()` | Adds a "PPMP Evaluation" menu to the spreadsheet UI for manual sync/refresh |
| Helper utilities | ID generation, score calculation, consensus logic |

**Sheet schemas created by `initializeEvaluationWorkbook()`:**

##### Sheet 1: `Medicine_Master`
| Column | Type | Notes |
|---|---|---|
| `Generic_Name` | Text (Key) | Populated by `syncMasterDrugList()` |
| `Category_Type` | Text | Manual or auto-filled |
| `Procurement_Mode` | Text | Manual or auto-filled |

##### Sheet 2: `Evaluations_Master` (AppSheet data source — 38 columns)
- **Metadata block** (5 cols): `Evaluation_ID`, `Timestamp`, `Evaluator`, `Generic_Name`, `Brand_Name`, `Manufacturer`
- **Part I block** (19 cols): `P1_01` through `P1_19` — Labeling & Regulatory (Yes/No/N/A)
- **Part II block** (6 cols): `P2_01` through `P2_06` — Physical Packaging (Yes/No/N/A)
- **Part III conditional block** (10 cols): `Requires_Reconstitution` (Yes/No), `P3_01` through `P3_09` — Dilution/Reconstitution (Yes/No/N/A)
- **Output block** (4 cols): `Part_I_Score`, `Part_II_Score`, `Remarks`, `Recommendation`

##### Sheets 3–5: `Evaluator 1`, `Evaluator 2`, `Evaluator 3`
- Mirror the exact same schema as `Evaluations_Master`
- Auto-populated via trigger — rows filtered by `Evaluator` column

##### Sheet 6: `Consolidated_Summary`
| Column | Purpose |
|---|---|
| `Generic_Name` | Drug identifier |
| `Brand_Name` | Product brand |
| `Manufacturer` | Manufacturer name |
| `Eval1_Recommendation` | Evaluator 1's verdict |
| `Eval2_Recommendation` | Evaluator 2's verdict |
| `Eval3_Recommendation` | Evaluator 3's verdict |
| `Eval1_PartI_Score` | Score breakdown |
| `Eval1_PartII_Score` | Score breakdown |
| `Eval2_PartI_Score` | Score breakdown |
| `Eval2_PartII_Score` | Score breakdown |
| `Eval3_PartI_Score` | Score breakdown |
| `Eval3_PartII_Score` | Score breakdown |
| `Consensus_Status` | Unanimous Recommended / Unanimous Not Recommended / Split Decision |
| `Consensus_Detail` | e.g., "2 of 3 Recommended" |
| `Last_Updated` | Timestamp of last recalculation |

**Key design decisions in `Code.gs`:**

1. **Score calculation**: `Part_I_Score` = count of "Yes" answers / (19 − count of "N/A" answers) × 100%. Same formula adapted for Part II (6 items).
2. **Consensus logic**: 
   - All 3 "Recommended" → `Unanimous Recommended`
   - All 3 "Not Recommended" → `Unanimous Not Recommended`
   - Mixed → `Split Decision` with detail string
   - Missing evaluator(s) → `Pending` with count
3. **Trigger strategy**: `onEdit` trigger on `Evaluations_Master` (compatible with AppSheet's webhook/sync writes).
4. **Idempotent sync**: `syncMasterDrugList()` clears and re-populates `Medicine_Master` each run to avoid duplicates.
5. **Error handling**: Try/catch with `Logger.log()` and `SpreadsheetApp.getUi().alert()` for user-facing errors.

---

### Component 2: AppSheet Configuration Guide

#### [NEW] [`AppSheet_Setup_Guide.md`](file:///c:/Users/PC 127/Desktop/PPMP Project/AppSheet_Setup_Guide.md)

A comprehensive, step-by-step blueprint document covering:

1. **Data Source Connection** — Connect the Google Sheet as data source, table-by-table
2. **Table Configuration** — Column types, keys, Ref relationships (`Generic_Name` → `Medicine_Master`)
3. **Slices** — `Slice_Evaluator1/2/3` with filter formulas
4. **Form View (5-page mobile form)**:
   - Page 1: Header Info (Evaluator selector, Generic_Name dropdown, Brand, Manufacturer)
   - Page 2: Part I — Labeling & Regulatory (19 Yes/No/N/A fields)
   - Page 3: Part II — Physical Packaging (6 Yes/No/N/A fields)
   - Page 4: Part III — Reconstitution (conditional, `Show_If = [Requires_Reconstitution] = "Yes"`)
   - Page 5: Verdict (Remarks, Recommendation)
5. **UX Views (Bottom Navigation)**:
   - 📝 New Evaluation (Form view → `Evaluations_Master`)
   - 👤 My Evals 1 (Deck view → `Slice_Evaluator1`)
   - 👤 My Evals 2 (Deck view → `Slice_Evaluator2`)
   - 👤 My Evals 3 (Deck view → `Slice_Evaluator3`)
   - 📊 Summary (Table view → `Consolidated_Summary`)
6. **Format Rules**:
   - `Recommendation = "Recommended"` → Green background/badge
   - `Recommendation = "Not Recommended"` → Red background/badge
   - `Consensus_Status = "Unanimous Recommended"` → Green row highlight
   - `Consensus_Status = "Split Decision"` → Yellow/amber row highlight
7. **Virtual Columns** — `Part_I_Score` and `Part_II_Score` AppSheet expressions
8. **Security & User Filtering** — Optional `USEREMAIL()` based row filtering
9. **Automation** — Bot/webhook to call Apps Script on form save

---

### Component 3: Deployment & Setup Instructions

#### [NEW] [`README.md`](file:///c:/Users/PC 127/Desktop/PPMP Project/README.md)

- Prerequisites (Google Workspace account, AppSheet access)
- Step-by-step deployment instructions for Apps Script
- How to run `initializeEvaluationWorkbook()` and `syncMasterDrugList()`
- How to set up the `onEdit` trigger
- Link to the AppSheet setup guide

---

## Open Questions

> [!IMPORTANT]
> **Target Spreadsheet**: Should `Code.gs` create a brand-new Google Spreadsheet, or should it initialize sheets inside an existing spreadsheet that you'll bind the script to? The current plan assumes it runs as a **container-bound script** (opened via Extensions → Apps Script from within the target spreadsheet).

> [!IMPORTANT]
> **Score formula interpretation**: The plan calculates scores as `(Yes count) / (Total items − N/A count) × 100%`. Should "No" answers carry negative weight, or is it purely a compliance percentage?

> [!NOTE]
> **Evaluator identity**: The plan uses a simple dropdown (`Evaluator 1/2/3`) in the AppSheet form. Would you prefer evaluator identification tied to Google account email via `USEREMAIL()` instead?

> [!NOTE]
> **Category_Type and Procurement_Mode** in `Medicine_Master`: These columns are in the schema but the PPMP Master List source only provides `Generic_Name`. Should these be manually entered, or do you have another source for them?

---

## Verification Plan

### Manual Verification
1. **Paste `Code.gs`** into a container-bound Apps Script project
2. **Run `initializeEvaluationWorkbook()`** — confirm all 6 sheets are created with correct headers, styling, and validations
3. **Run `syncMasterDrugList()`** — confirm `Medicine_Master` is populated with generic drug names from the PPMP Master List
4. **Manually add a test row** to `Evaluations_Master` — confirm the `onEdit` trigger mirrors it to the correct evaluator tab and updates `Consolidated_Summary`
5. **Follow the AppSheet setup guide** to configure the mobile app and test end-to-end flow

### Structural Checks
- Verify all 38+ columns in `Evaluations_Master` match the schema exactly
- Verify data validation dropdowns (Yes/No/N/A, Evaluator enum, Recommendation enum) are applied
- Verify score formulas compute correctly with edge cases (all N/A, mixed, etc.)
- Verify consensus logic handles 0, 1, 2, and 3 evaluators correctly
