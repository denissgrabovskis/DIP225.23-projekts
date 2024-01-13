import os
import shutil
from datetime import timedelta

# Function to round time to the nearest 5 minutes
ROUND_TO = timedelta(minutes=5)
def roundTo5min(dt):
    dateInSeconds = dt - dt.min
    roundedSecondsPassed = round(dateInSeconds / ROUND_TO) * ROUND_TO
    return dt.min + roundedSecondsPassed

# Paths
script_dir = os.path.dirname(__file__)

parent_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
clocks_file_path = os.path.join(parent_dir, 'clocks.txt')
excel_file_path = os.path.join(parent_dir, 'clocks.xlsx')

log_dir = os.path.join(script_dir, 'log')
previous_clocks_file_path = os.path.join(log_dir, 'previous_clocks.txt')

def save_clocks_copy():
    # Ensure the destination directory exists
    os.makedirs(log_dir, exist_ok=True)

    # Copy and rename the file. Create if not exists
    open(clocks_file_path, 'a')
    shutil.copy(clocks_file_path, previous_clocks_file_path)

def get_all_clocks():
    with open(clocks_file_path, 'r') as file:
        return file.readlines()
    
def get_last_date(clocks):
    last_line_with_date = len(clocks)-1 # start with the last line

    while clocks and last_line_with_date >= 0:
        clock_line = clocks[last_line_with_date]
        if ('/' in clock_line): # has date
            return clock_line.strip().split('\t')[0]
        else:
            last_line_with_date -= 1

    return ""

def get_month_from_str(str_date):
    return int(str_date.split('/')[1]) if str_date else ""