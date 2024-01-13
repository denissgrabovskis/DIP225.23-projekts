from datetime import datetime
from helper import save_clocks_copy, roundTo5min, get_all_clocks, get_last_date, get_month_from_str, clocks_file_path
from update_excel import generate_excel

save_clocks_copy()

all_clocks = get_all_clocks()
last_date = get_last_date(all_clocks)

# Get today's date and rounded time
today = datetime.now()
clock_in_time = roundTo5min(today).strftime('%H:%M')

if get_month_from_str(last_date) != today.month:
    generate_excel()
    all_clocks = []

today = today.strftime('%d/%m/%Y')

# Edit the source file if the last line date is different from today
with open(clocks_file_path, 'a') as file:
    if all_clocks:
        file.write('\n')
    if last_date != today:
        file.write(f"{today}")
    file.write(f"\t{clock_in_time}\t")