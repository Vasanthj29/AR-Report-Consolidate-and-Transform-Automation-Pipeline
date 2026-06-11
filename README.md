# AR Report Consolidate & Transform Automation Pipeline

> **⚠️ DISCLAIMER:** All datasets used in this project are completely dummy/random data created solely for demonstration purposes. No organization's confidential, proprietary, or real patient data is used anywhere in this code or examples. Any resemblance to actual persons, organizations, or data is purely coincidental.

## Background

In a typical hospital finance setup, Accounts Receivable (AR) outstanding reports are generated from multiple business units, each producing their own Excel file. Additionally, reference data like customer master, claim submissions, and accounting self‑pay transactions sit in separate files. Manually consolidating these every month involves countless VLOOKUPs, copy‑pasting, and formula checks – and is prone to errors.

This pipeline automates the entire process: it reads all outstanding reports from a folder, merges them, enriches the data with standardised names and IDs, calculates ageing buckets and financial year, applies business rules to classify invoice status, and outputs a single ready‑to‑use Excel file.

## What This Code Does

1. **Reads** all Excel files from a specified source folder.
2. **Consolidates** multiple files and sheets into a single pandas DataFrame.
3. **Standardises** business unit names (e.g., "AHEL Bhubaneswar BU" → "Bhubaneswar").
4. **Merges** customer master data to get standard payer names and categories (VLOOKUP equivalent).
5. **Converts** bill dates from string to datetime, then calculates ageing in days.
6. **Assigns** ageing buckets (0‑30D, 31‑60D, 61‑90D, 91‑120D, 121‑180D, 181D‑1Y, 1Y‑2Y, 2Y‑3Y, 3Y+).
7. **Derives** financial year based on India logic (April‑March); marks dates before April 2021 as "Prior".
8. **Joins** claim submission report to add Claim IDs.
9. **Joins** accounting self‑pay transactions to capture self‑paid amounts.
10. **Applies** a rule‑based engine to classify invoice status:  
    `Fully Unpaid`, `Partially Paid`, `Non-AR`, `Unaccounted Payments`, or `Other`.
11. **Retains** audit columns (differences between Bill Amount, Outstanding, and Self Paid Amount).
12. **Drops** unnecessary intermediate columns and reorders the final columns logically.
13. **Exports** the final reconfigured data to an Excel file with a timestamp.

## Input Requirements

- All source AR outstanding Excel files must be placed in a single folder (e.g., `D:\Learning\AR\Outstanding`).
- Each file should contain the following columns (names exactly as shown):

| Business Unit | Customer Number | Customer Name | Customer Category | Bill Number | Bill Date | Type | Patient Name | Bill Amount | Settlement | Outstanding |

- Three reference files must exist at the paths defined in the notebook:
  - **Customer Master** – `Payer_Customer_Master.xlsx` with columns: `Payer_Name_Raw`, `Std_Payer_Name`, `Std_Payer_Category`
  - **Claim Submission** – `Claim_Submission_Report_Dummy.xlsx` with columns: `Claim Bill Number`, `Claim ID`
  - **Accounting Self‑Pay** – `Accounting_Report_Dummy.xlsx` with columns: `Accounting Bill Number`, `Accounted Amount`, `Customer Category` (must contain rows where `Customer Category = "Self Pay"`)

- `Bill Date` in source files should be in format like `12-Mar-2026` (the script handles `%d-%b-%Y`).

## Output

One Excel file per run, named:  
`Consolidated Outstanding with working_DD-MM-YY-HHMMSS.xlsx`

The output file contains the following columns (in this order):

| Column Name |
|-------------|
| Business Unit |
| Std Unit |
| Customer Number |
| Customer Name |
| Customer Category |
| Bill Number |
| Bill Date |
| Type |
| Patient Name |
| Bill Amount |
| Settlement |
| Outstanding |
| Std_Payer_Name |
| Std_Payer_Category |
| Financial Year |
| Month_Year |
| Ageing |
| Ageing Bucket |
| Claim ID |
| BillAmount_vs_Outstanding |
| Self Paid Amount |
| SelfPaid_vs_BillAmount |
| Diff_vs_Outstanding |
| Invoice Status |

## Dependencies

- Python 3.7+
- pandas
- xlsxwriter
- glob (built‑in)
- datetime (built‑in)
- os (built‑in)

Install with:  
`pip install pandas xlsxwriter`

## How to Run

1. Update the file paths inside the notebook:
   - `Source` – folder containing raw outstanding reports
   - `customer_master` – path to `Payer_Customer_Master.xlsx`
   - `claim_submission` – path to `Claim_Submission_Report_Dummy.xlsx`
   - `accounting` – path to `Accounting_Report_Dummy.xlsx`
   - Output path in the `pd.ExcelWriter` line (optional)
2. Run all cells in the Jupyter notebook.
3. Find the output Excel file in the specified output location.

## Why This Pipeline Saves Time

- Eliminates manual VLOOKUPs across 3‑4 different files.
- No more copy‑pasting rows from multiple unit reports.
- Ageing buckets and financial year are calculated automatically – no Excel formulas to drag.
- Invoice status logic is consistent and repeatable.
- Audit columns let you spot data mismatches instantly.

## Customisation Notes

- To add a new business unit, extend the `std_unit_assign()` function.
- To modify ageing buckets, change the thresholds in `Ageing_Bucket()`.
- To change invoice status rules, edit the `partial_tags()` function.
- To include additional reference files, add another `df.merge()` step.

## License

This project is provided as‑is for internal automation. Modify and reuse freely.
