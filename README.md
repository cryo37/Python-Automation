# Python Excel Automation

## Overview
This project automates the creation of an Excel report using Python.
It generates a structured Excel file from tabular data and automatically opens it in Microsoft Excel using Windows COM automation.

## Features
- Creates Excel files programmatically using `openpyxl`
- Uses `pandas` for structured data handling
- Automates Microsoft Excel launch via `win32com`
- Demonstrates end-to-end office automation

## Tech Stack
- Python
- Pandas
- OpenPyXL
- win32com (Windows COM Automation)

## How It Works
1. Data is created using a Pandas DataFrame
2. The DataFrame is written to an Excel file
3. Excel is automatically launched and the file is opened

## How to Run
```bash
pip install -r requirements.txt
python src/excel_automation.py
