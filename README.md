# UCES Margin Analyzer

A Streamlit dashboard for managing UCES margin data. It helps you load a master Excel file, review margin health, filter records, edit rows, and export the updated workbook.

## Features

- Load an existing Excel file or map mismatched columns
- Add, edit, and delete margin records
- Track margins with color-coded status bands
- Filter by status, project, vendor, site ID, and margin range
- Review data in a styled table with formatted columns
- Download the updated Excel file with preserved styling and validation
- Save session data locally as JSON for continuity between runs

## Tech Stack

- Python
- Streamlit
- Pandas
- OpenPyXL

## Required Input Columns

The app is designed around these core fields:

- Quotation No
- Po Huawei
- Linked PR Subcon
- Date of PR
- Vendor Name
- Project
- Site ID
- Line Items
- Po Huawei (Unit Price)
- Requested Qty
- Total
- Po Subcon (Unit Price)
- Qty
- Sub Total
- Profit
- Margin%
- Status
- Margin Reason

If your spreadsheet uses slightly different headers, the app can map many of them during upload.

## Local Setup

1. Clone the repository.

```bash
git clone https://github.com/EkhsanFitri94/UCES-Margin-Analyzer.git
cd UCES-Margin-Analyzer
```

2. Install dependencies.

```bash
pip install -r requirements.txt
```

3. Start the dashboard.

```bash
streamlit run master_file_app.py
```

## How To Use

1. Open the app in your browser.
2. Upload your master Excel file.
3. Confirm column mapping if the headers do not match exactly.
4. Add or edit entries using the form.
5. Use filters to review margin risk.
6. Download the updated workbook when finished.

## Margin Guide

- Green: margin is at or above 30%
- Yellow: margin is between 20% and 29%
- Red: margin is below 20%

## Notes

- The app stores session data in `uces_app_data.json`.
- The workbook export keeps formatting, column widths, and project validation.
- This repo is a strong candidate for continued improvement, especially if you want to add screenshots and a sample workbook.
