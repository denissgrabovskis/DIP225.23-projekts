import os
import shutil
from datetime import timedelta

# Mainīgais, kurš glabā līdz kurām tuvākām vērtībām jānoapaļo
ROUND_TO = timedelta(minutes=5)
def roundTo5min(dt): # saņem datetime, atgriež to noapaļoto līdz tuvākām ROUND_TO
    dateInSeconds = dt - dt.min
    roundedSecondsPassed = round(dateInSeconds / ROUND_TO) * ROUND_TO
    return dt.min + roundedSecondsPassed

# Py skripta ceļs (../src)
script_dir = os.path.dirname(__file__)

parent_dir = os.path.abspath(os.path.join(script_dir, os.pardir)) # ../src mapes ceļš (..)
clocks_file_path = os.path.join(parent_dir, 'clocks.txt') # ../clock.txt
excel_file_path = os.path.join(parent_dir, 'clocks.xlsx') # ../clock.xlsx

log_dir = os.path.join(script_dir, 'log') # ../src/log
previous_clocks_file_path = os.path.join(log_dir, 'previous_clocks.txt') # ../src/log/previous_clocks.txt

# saglabā clock.txt kopiju
def save_clocks_copy():
    # izveido direktoriju, ja vēl nav izveidota (src/log)
    os.makedirs(log_dir, exist_ok=True)

    # (izveido un) kopē failu src/log
    open(clocks_file_path, 'a')
    shutil.copy(clocks_file_path, previous_clocks_file_path)

# atgriež faila clock.txt rindus
def get_all_clocks():
    with open(clocks_file_path, 'r') as file:
        return file.readlines()
    
# saņem clock.txt rindus un atrod pedējo fiksēto datumu. Ja nav - tukšs
def get_last_date(clocks):
    last_line_with_date = len(clocks)-1 # sākt ar pedejo rindu

    while clocks and last_line_with_date >= 0:
        clock_line = clocks[last_line_with_date]
        if ('/' in clock_line): # '/' būs tikai ja ir datums ([dd/mm/yyyy]  hh:mm   hh:mm)
            return clock_line.strip().split('\t')[0] # atgriež tikai datumu
        else:
            last_line_with_date -= 1

    return ""

# saņem datumu kā virkni un atgriež mēnesi kā veselo skaitli ('dd/mm/yyyy' -> int(mm))
def get_month_from_str(str_date):
    return int(str_date.split('/')[1]) if str_date else ""