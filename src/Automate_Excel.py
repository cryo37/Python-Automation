import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from win32com.client import Dispatch
from pathlib import Path

data = {
    "Asset Name" : ["Asset 1", "Asset 2"],
    "Month 1" : [15,30],
    "Month 2" : [5,35]
}

df = pd.DataFrame(data)

workbook = Workbook()
sheet = workbook.active

for row in dataframe_to_rows(df, index=False, header=True):
    sheet.append(row)

output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

file_path = output_dir / "sample_output.xlsx"

workbook.save(filename=file_path)

x1 = Dispatch("Excel.Application")
x1.Visible = True

wb = x1.Workbooks.Open(r'C:\Users\Anurag\Desktop\AWS\Python\Python Automation Projects\output\sample_output.xlsx')

