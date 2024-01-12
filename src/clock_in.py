import shutil
import os
from datetime import datetime, timedelta

# Function to round time to the nearest 5 minutes
ROUND_TO = timedelta(minutes=5)
def roundTo5min(dt):
    dateInSeconds = dt - dt.min
    roundedSecondsPassed = round(dateInSeconds / ROUND_TO) * ROUND_TO
    return dt.min + roundedSecondsPassed

# Paths
script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
source_file = os.path.join(parent_dir, 'clocks.txt')
destination_dir = os.path.join(script_dir, 'log')
destination_file = os.path.join(destination_dir, 'previous_clocks.txt')

# Ensure the destination directory exists
os.makedirs(destination_dir, exist_ok=True)

# Copy and rename the file. Create if not exists
open(source_file, 'a')
shutil.copy(source_file, destination_file)

# Read the last line of the file to check the date
with open(source_file, 'r') as file:
    times = file.readlines()
    if not times:
        last_line_with_date = ""
    else:
        last_line_with_date = -1
        while True:
            if times and -len(times) <= last_line_with_date:
                if ('/' in times[last_line_with_date].strip()):
                    last_line_with_date = times[last_line_with_date].strip()
                    break
                else:
                    last_line_with_date -= 1
            else:
                last_line_with_date = ""
                break


# Get today's date and rounded time
today = datetime.now()
clock_in_time = roundTo5min(today).strftime('%H:%M')
today = today.strftime('%d/%m/%Y')

# Edit the source file if the last line date is different from today
with open(source_file, 'a') as file:
    if times:
        file.write('\n')
    if not last_line_with_date.startswith(today):
        file.write(f"{today}")
    file.write(f"\t{clock_in_time}\t")
