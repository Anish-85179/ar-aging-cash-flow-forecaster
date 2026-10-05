# Accounts Receivable (AR) Aging & Cash Flow Forecaster

A dynamic, multi-tab Excel financial model designed to track Accounts Receivable balances, classify aging buckets, evaluate collection probabilities, and forecast expected weekly cash inflows.

---

## 📁 Workbook Structure

* **`AR_Invoice_Master`**: Core dataset containing raw invoice details, client information, issue dates, and original balances.
* **`Aging_Analysis`**: Executive dashboard with top-level KPI summary cards (`Total Outstanding`, `Total Overdue`, `Portfolio DSO`) and calculated columns for days overdue and aging buckets.
* **`Cash_Flow_Forecaster`**: Risk-weighted cash flow modeling tab utilizing an Aging Risk Matrix to calculate expected cash realization by target week.
* **`Documentation_Setup`**: Standardized documentation cover page outlining data dictionaries, formulas, and version control.

---

## 🛠️ Key Excel Formulas & Logic

1. **Days Overdue Calculation (`Aging_Analysis!E6`)**:
   ```excel
   =IF(D6=0, 0, MAX(0, TODAY() - C6))
