import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# Function to open or create an Excel workbook
def open_or_create_workbook(filename):
    try:
        workbook = load_workbook(filename, data_only=True)
    except FileNotFoundError:
        workbook = Workbook()
    return workbook

script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
excel_file = os.path.join(parent_dir, 'clocks.xlsx')
text_file = os.path.join(parent_dir, 'clocks.txt')


# Open or create the workbook
workbook = open_or_create_workbook(excel_file)

# Add new sheet with specific format
sheet_name = datetime.now().strftime('%b%y')
if sheet_name not in workbook.sheetnames:
    if len(workbook.sheetnames) == 1 and workbook.sheetnames[0] == 'Sheet': # Newly created workbook
        workbook.active.title = sheet_name
    else:
        workbook.create_sheet(title=sheet_name)

sheet = workbook[sheet_name]

# Set headers
headers = ["Date", "Clock In", "Clock Out", "", "Hours"]
for col, header in enumerate(headers, 2):  # Starting from column B
    if not header:
        continue
    sheet[f"{get_column_letter(col)}1"] = header
    sheet[f"{get_column_letter(col)}1"].style = 'Check Cell'

# Read data from clock.txt and copy it to the Excel sheet
open(text_file, 'a')
with open(text_file, 'r+') as file:
    times = file.readlines()
    for row_index, line in enumerate(times, 2):  # Starting from row 2
        *date, clock_in, clock_out = line.strip().split('\t')
        if not clock_in or not clock_out:
            continue
        sheet[f"B{row_index}"] = date[0] if date else ""
        sheet[f"C{row_index}"] = clock_in
        sheet[f"C{row_index}"].number_format = 'hh:mm'
        sheet[f"D{row_index}"] = clock_out
        sheet[f"D{row_index}"].number_format = 'hh:mm'

        # Formula for hours calculation
        sheet[f"F{row_index}"] = f"=D{row_index}-C{row_index}"
        sheet[f"F{row_index}"].number_format = 'hh:mm'

        clock_in_hour, _ = clock_in.split(':')
        clock_out_hour, clock_out_min = clock_out.split(':')
        if (int(clock_in_hour) > int(clock_out_hour)):
            sheet[f"D{row_index}"] = f"{int(clock_out_hour)+24}:{clock_out_min}"

if times:
    sheet[f"F{row_index+1}"] = f'=SUM(F2:F{row_index})'
    sheet[f"F{row_index+1}"].style = '40 % - Accent4'
    sheet[f"F{row_index+1}"].number_format = '[h]:mm'

# Save the workbook
workbook.save(excel_file)
