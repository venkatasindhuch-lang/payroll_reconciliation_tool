# Payroll Reconciliation Checker
## Live Demo

🔗 [Open the Payroll Reconciliation Checker](https://smart-payroll-reconciliation.streamlit.app)

*Demo only:* Please use fictional or sample payroll data. Do not upload real employee or confidential payroll information.

*Developed by Sindhu | Payroll Reconciliation Portfolio Project*
A simple payroll reconciliation web app built with **Python, Pandas and Streamlit**.

It compares an **expected payroll file** with a **processed payroll file** and highlights:
- missing employees
- unexpected employees
- gross pay differences
- PAYE differences
- Employee NI differences
- pension differences
- other deduction differences
- net pay differences

It also creates an **exceptions report** that can be downloaded as CSV or Excel.

## Why I built this

Payroll teams need to make sure employees are paid accurately and on time.  
This project demonstrates a practical payroll control: comparing expected payroll values with processed payroll values before or after payroll completion.

It is designed as a portfolio project for payroll, finance and HR operations roles.

## Features

- Upload CSV or Excel files
- Reconcile employees using Employee ID
- Compare selected payroll fields
- Set a variance tolerance
- Detect duplicate Employee IDs
- Detect missing or unexpected employees
- Show payroll variance by field
- Show reconciliation KPIs
- Export a full Excel reconciliation report
- Export an exceptions-only CSV

## Required columns

Both payroll files should contain:

```text
Employee ID
Employee Name
Gross Pay
PAYE
Employee NI
Pension
Other Deductions
Net Pay
```

You can choose which available payroll fields to compare in the app.

## Example

Expected payroll:

| Employee ID | Employee Name | Gross Pay | PAYE | Employee NI | Pension | Other Deductions | Net Pay |
|---|---|---:|---:|---:|---:|---:|---:|
| 1001 | Alex Smith | 2500 | 300 | 145 | 100 | 0 | 1955 |

Processed payroll:

| Employee ID | Employee Name | Gross Pay | PAYE | Employee NI | Pension | Other Deductions | Net Pay |
|---|---|---:|---:|---:|---:|---:|---:|
| 1001 | Alex Smith | 2500 | 300 | 145 | 100 | 0 | 1955 |

If the processed payroll differs from the expected payroll by more than the selected tolerance, the employee is flagged as an exception.

## Run locally

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/payroll-reconciliation-tool.git
cd payroll-reconciliation-tool
```

### 2. Create a virtual environment

Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the app

```bash
streamlit run app.py
```

## Test data

The `sample_data` folder contains:
- `expected_payroll.csv`
- `processed_payroll.csv`

The processed sample intentionally contains reconciliation differences so the exception report can be tested.

## Possible future improvements

- payroll period selector
- department/cost-centre reconciliation
- employer NI reconciliation
- statutory payments such as SSP/SMP/SPP
- pension scheme reconciliation
- payroll-to-GL reconciliation
- configurable column mapping
- audit trail and sign-off
- PDF reconciliation summary
- dashboard charts
- secure user login
- database storage

## Important note

This project is a **reconciliation and portfolio tool**, not payroll calculation or tax advice software.  
Real payroll systems should use validated payroll rules, secure controls, authorised access and appropriate data-protection practices.

## Tech stack

- Python
- Streamlit
- Pandas
- OpenPyXL

## License

MIT
