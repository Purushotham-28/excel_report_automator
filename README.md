# Excel Report Automator 📊

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Pandas](https://img.shields.io/badge/Pandas-2.0+-150458?style=flat-square&logo=pandas&logoColor=white)](https://pandas.pydata.org)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.1+-217346?style=flat-square&logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io)
[![Flask](https://img.shields.io/badge/Flask-3.0+-000000?style=flat-square&logo=flask&logoColor=white)](https://flask.palletsprojects.com)
[![React](https://img.shields.io/badge/React-18-61DAFB?style=flat-square&logo=react&logoColor=black)](https://react.dev)

An automated data cleaning, validation, and executive report generator available as both a **Python CLI tool** and an **interactive React/Flask web dashboard**.

---

## 📌 Problem Solved

Business data exported from CRMs, ERPs, or legacy databases is frequently messy—containing duplicate records, missing values, negative prices/quantities, and inconsistent column headers. Manually auditing and re-formatting these datasets in Excel is time-consuming, repetitive, and error-prone.

**Excel Report Automator** automates this pipeline end-to-end:
1. Standardizes headers and cleans text fields.
2. Removes exact duplicate rows.
3. Segregates clean transactions from invalid rows (flagging specific failure reasons).
4. Calculates category totals, averages, and processing KPIs.
5. Exports a publication-ready Excel report (`.xlsx`) formatted with custom openpyxl styling.

---

## 🛠 Tech Stack

- **Core Engine & Data Processing**: Python 3, Pandas, NumPy
- **Excel Formatting**: OpenPyXL
- **CLI Interface**: `argparse`
- **Backend API**: Flask, Flask-CORS
- **Frontend Dashboard**: React 18, HTML5/CSS3 (Glassmorphism UI), Lucide Icons

---

## ⚙️ Quickstart & Setup

### 1. Clone & Navigate to Project Directory
```bash
git clone https://github.com/your-username/excel-report-automator.git
cd excel-report-automator
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 🚀 How to Run

### Option A: Command-Line Interface (CLI)

Run the automator directly on any CSV or Excel file:

```bash
# Basic usage (auto-names report output)
python excel_report_automator.py --input sample_input.csv

# Custom output file name
python excel_report_automator.py --input sample_input.csv --output custom_report.xlsx
```

**Terminal Output Example**:
```text
Reading input file: sample_input.csv
Standardized columns: ['transaction_id', 'date', 'category', 'product_name', 'quantity', 'unit_price', 'total_amount', 'region']
Selected category column for summary: 'category'
Generating Excel report: custom_report.xlsx

=======================================================
 EXCEL REPORT AUTOMATION COMPLETED
=======================================================
  Input File:          sample_input.csv
  Output Report:       custom_report.xlsx
  Total Raw Rows:      20
  Duplicates Removed:  2
  Valid Cleaned Rows:  12
  Flagged Error Rows:  6
-------------------------------------------------------
  Sheets Created:
   1. Cleaned Data  (Valid transactions)
   2. Errors        (Flagged transactions with error reasons)
   3. Summary       (KPI metrics & Category aggregations)
=======================================================

12 rows cleaned, 6 rows flagged as errors, output saved to custom_report.xlsx
```

---

### Option B: Web Application Dashboard

Launch the Flask backend server to use the React drag-and-drop interface:

```bash
python app.py
```
Open **`http://127.0.0.1:5000`** in your web browser to:
- Drag and drop CSV or Excel files.
- Click **"Test Sample Dataset"** to run an instant demonstration.
- Preview **Cleaned Data**, **Flagged Errors**, and **Executive Summaries** interactively.
- Download the generated `.xlsx` report with a single click.

---

## 📊 Before & After Data Transformation

### **Before** (`sample_input.csv` - 20 Raw Rows):
- **Raw Headers**: `Transaction ID`, `Unit Price ($)`, `Total Amount`, `Region` (inconsistent casing and symbols).
- **Messy Issues**:
  - 2 exact duplicate rows (`TXN1001`, `TXN1004`).
  - Missing categories, missing regions, or missing quantities (e.g. `TXN1006`, `TXN1010`, `TXN1014`).
  - Invalid negative values in quantities and prices (e.g. `TXN1003` has `-3` quantity, `TXN1007` has `-$75` price).

### **After** (`sample_input_report.xlsx` - 3 Styled Worksheets):
- **Sheet 1: `Cleaned Data`**: 12 fully validated, deduplicated records with auto-fitted columns, bold navy headers, right-aligned currency formatting (`$#,##0.00`), and clean headers (`transaction_id`, `unit_price`, `total_amount`).
- **Sheet 2: `Errors`**: 6 flagged invalid rows with an explicit `error_reason` column (e.g., `"Negative value in 'unit_price' (-75.0); Negative value in 'total_amount' (-150.0)"`).
- **Sheet 3: `Summary`**:
  - **Processing Overview**: High-level KPIs (Total Input, Duplicates Removed, Valid Cleaned Rows, Error Count, Error Rate %).
  - **Category Breakdown**: Aggregated totals and averages grouped by Category (`Electronics`, `Furniture`, `Office Supplies`) featuring a styled **Grand Total** row with double-line borders.

---

## 📁 Project Structure

```text
excel_report_automator/
├── excel_report_automator.py  # Core CLI script & data processing engine
├── app.py                     # Flask REST API backend
├── requirements.txt           # Python project dependencies
├── sample_input.csv           # Sample messy CSV dataset (~20 rows)
├── templates/
│   └── index.html             # React single-page frontend application
└── reports_output/            # Output folder for generated Excel reports
```

---

## 📄 License

MIT License © 2026
