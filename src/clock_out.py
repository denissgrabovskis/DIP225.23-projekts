from datetime import datetime
from helper import save_clocks_copy, get_all_clocks, get_last_date, roundTo5min, get_month_from_str, clocks_file_path
from update_excel import generate_excel

# saglāba clock.txt kopiju pirms darbībām ar to
save_clocks_copy()

all_clocks = get_all_clocks() # nolasa rindus no clock.txt
last_date = get_last_date(all_clocks) # atrod pēdējo datumu kā virkni

# saņem šodienas datumu un saglāba noapoļoto laiku kā virkni
today = datetime.now()
clock_out_time = roundTo5min(today).strftime('%H:%M')

# jā ir sākusies jauns mēness
if get_month_from_str(last_date) != today.month:
    generate_excel() # jāizveido Excel un jānodzēs clock.txt, lai sāktu pierakstus no jauna
    all_clocks = []

# saglaba datumu kā virkni
today = today.strftime('%d/%m/%Y')

# pieraksta jauno darba pabeigšanas laiku
with open(clocks_file_path, 'a') as file:
    if last_date != today: # ja pēdēja laika ieraksts ir cits datums, tad jāuzraksta jaunu datumu
        if all_clocks:  # ...un jā ir ieraksti pirms šīm, jāsāk ar jauno rindu
            file.write('\n')
        file.write(f"{today}\t\t")
    file.write(f"{clock_out_time}") # beidzot jāuzraksta pabeigšanas laiku