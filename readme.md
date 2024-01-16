# Deniss Grabovskis 231RDB217

## Projekta apraksts
Es strādāju attālināti, tāpēc darba devējs lūdz pierakstīt nostrādātās stundas un mēneša beigās atsūtīt Excel failu ar apkopotām stundām par visu nostrādāto mēnesi. Lai to izdarītu, sāku darba sākumā un beigās ierakstīt datumu un laiku (noapaļots līdz tuvākajām piecām minūtēm) TXT failā, bet mēneša beigās visas stundas kopēt programmā Excel, formatēt un nosūtīt.  
Es to daru, jo atšķirībā no vienkārša teksta faila Excel faila palaišana katru dienu mana datorā aizņemtu ilgu laiku. Tāpat, lai manuāli nepārrakstītu visus datumus un stundas Excel failā meneša beigās, es pieradu tos sākotnēji ierakstīt teksta failā īpašā formātā, kas ļauj vienkārši atlasīt visu tekstu, kopēt un ielīmēt Excel failā:   
`{datums}\t{sākuma laiks}\t{beigas laiks}`   
Ja pa dienu strādāju ar pārtraukumu, tad šādā formātā pierakstu tikai pirmo rindiņu. Nākamās rindas izskatās šādi:   
`\t{sākuma laiks}\t{beigas laiks}`   

Tātad mans tipiskais teksta fails mēneša beigās izskatītos šādi:  
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYTU2is9B20yd3DyZtCtq3-7WVcqueEONVccHNAEzJaNH0Fh4qo-gnfwVyUQgO3E1MwxC9snfvMqnG84ucE7pYqlNPIb9A=w1920-h955">
</p>

Uz ta pamata es veidoju Excel failu, kas izskatās šādi:
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYR18_K_7xsG_WIrzzroi-xow02mIA5FUHgfX-oT_st3BROT5ggFuw0fJVHYNpP7aGFEGE1XU9wBPrSGYBHzFFt02Qxvkw=w1920-h955">
</p>

Dažreiz es pabeidzu darbu pēc pusnakts, tāpēc "Clock Out" laiks var būt mazāks par "Clock In". Šādas situācijas jārisina, pieskaitot 24 stundas "Clock Out" laikam, lai Excel formula strādātu pareizi:
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYSIBQlzwHoHLEebpzSfTvs45J6uc9Q-keJ13cf335md_gODvEq3oG25V8aU9fyzS_dPSLPATFZi8kqTbRcomLI8q-1J=w1920-h955">
</p>

## Faili
Apkopojot iepriekš minētās prasības, sāku rakstīt kodu. Es nolēmu, ka visi faili atradīsies mapē [`/src`](src). Mape tiek ievietota direktorijā, kurā tiks izveidoti un atjaunināti `clock.txt` un `clock.xlsx`.

### Clock In & Clock Out
Pirmie divi Python faili ir [`clock_in.py`](src/clock_in.py) un [`clock_out.py`](src/clock_out.py). Viņi veic šādas darbības:
* saglabā pašreizējā faila `clock.txt` kopiju mapē `src/log/previous_clocks.txt`. Tas tika darīts, lai, ja programma pēkšņi sabojā `clock.txt` vai sabojā tā formatējumu, es varētu atjaunot tās iepriekšējo versiju.
* izlasa visus ierakstus un atroda pēdējo ar datumu, lai, ierakstot, saprastu, vai jāraksta jauns datums, vai vienkārši jāpievieno jauno laika ierakstu
* ja kāda iemesla dēļ Excel fails netika izveidots mēneša beigās, tas tiks izveidots šajā darbībā, lai teksta fails tiktu izveidots no sākuma (jauns mēness). Funkcija `generate_excel` tiks izsaukta no faila [update_excel.py](src/update_excel.py), kas ir aprakstīts [tālāk](###update-excel).
* rakstīšanai tiks atvērts teksta fails `clock.txt`. Un tiks pievienots šodienas datums (ja nepieciešams) un laiks (vai nu darba sākums, vai beigas).

### Update Excel
[update_excel.py](src/update_excel.py) veic šādas darbības:
* atver clock.xlsx rakstīšanai (vai izveido)
* izveido jaunu izklājlapu ar nosaukumu **"MMYY"** *(piemēram, "Apr24")*
* pievieno galvenes: Date, Clock In, Clock Out, Hours ar stiliem **"Check Cell"**
* atver failu `clock.txt` un nolasa visus laikus ar datumiem. Katrs datums un laiks tiek ievadīti savā rindā un kolonnā, un formāts ir iestatīts uz **"hh:mm"**.
  * Ja Clock Out vērtība ir mazāka par Clock In, tad pievieno 24 stundas
* Ievada formulu katra rindā Hours kolonnā, kas atņem laiku darba beigās no tā sākuma. Tas aprēķina, cik daudz laika es pavadīju darbā noteiktā dienā.
* Tabulas beigās visas stundas tiek summētas ar funkciju Sum kolonnā Hours, stils ir iestatīts uz **"40% - Accent4"** un formāts **"[h]:mm"**
* Excel fails tiek saglabāts, teksta fails tiek izdzēsts

### Helper
Dažās vietās kods failos atkārtojas, tāpēc šīs koda daļas tika pārvietotas uz atsevišķu failu [helper.py](src/helper.py). Tajā tiek saglabāti šādi mainīgie, kurus izmanto citi faili:
* clocks_file_path - absolūtais ceļš uz teksta failu
* excel_file_path — absolūtais ceļš uz Excel failu
Un šīs metodes:
* roundTo5min(datetime): saņem datumu un atgriež to noapaļotu līdz tuvākajām 5 minūtēm. [Algoritms ir aprakstīts paņemts no šeit](https://stephenallwright.com/python-round-time-15-minutes/)
* save_clocks_copy(): saglabā teksta faila kopiju mapē `/src/log`
* get_all_clocks(): nolasa visus laikus no teksta faila un atgriež tos
* get_last_date(clocks): saņem visas stundas ar datumiem (virkņu masīvs) un atgriež pēdējo datumu.
* get_month_from_str(str_date): parsē virkni (dd/mm/gggg) un atgriež mēnesi kā vesels skaitlis.

### Bibliotēkas
* **datetime** - darbam ar datumiem un laikiem
* **os** - darbam ar ceļiem un direktorijiem
* **shutil** - darbam ar failiem. Izmanto, lai kopētu failu clock.txt uz `/src/log`.

## Automatizācija
Lai gan šie faili paši automatizē manu darbu, es varēju joprojām vairāk to vienkāršot. Es nolēmu izmantot *Task Scheduler*, lai automātiski palaistu skriptus. Lai to izdarītu, vispirms man bija jāsaprot, kuru manu darbību izmantot kā programmu palaišanas aktivizētāju. Es daru savu darbu programmā PhpStorm, tāpēc nolēmu, ka varu palaist `clock_in.py` un `clock_out.py` failus, kad attiecīgi atvēru un aizvēru programmu. Savukārt `update_excel.py` tiks palaists katra mēneša pirmajā dienā.

### PhpStorm
Es atveru *Task Scheduler -> Create Task*. Es piešķiru uzdevumam nosaukumu un atlasu opciju **"Run with highest priveleges"**.
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYS2Y8z6mHapmXStHoc4zCj8_7Dl36Xweh2BgLXflPOcwRwas4YbPd9-DoM6Q7gunfwCMRxqHyzjCimWzbONZXca9TQTKA=w1920-h955">
</p>

Cilnē **"Actions"** es izveidoju jaunu darbību un atlasu programmu `clock_in.py` (un `clock_out.py` otrām uzdevumam).
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYT-5WJLKpzRGrofGIHS1UFK2XRHRCVJtKF2NBJEmFRsU2SLDg9W7gGfyX6rX5s-a3f2IaW_4rrLheXwhFs4RQWQ0qRdSw=w1920-h955">
</p>

Cilnē **"Triggers"** es izveidoju jaunu trigeri. Es izvēlos palaist trigeri **"On an Event"**. Iestatījumos atlasu opciju **"Custom"** un noklikšķinu uz **"New Event Filter"**. Šajā posmā rodas problēma: operētājsistēmā Windows nav iebūvēta trigeri programmas palaišanai. Es aprakstīšu, kā es atrisināju šo problēmu nākamajā sadaļā [Event Viewer](##event-viewer) bet pagaidām šeit ir ko es uzrakstīju XML sadaļā.
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYTmlaWtKBLG2axaVYQcY6MQVfvcnytlfBu6sC2qORqneG7gVoWHkfcFfOUUQHfrCGhqennw0v4Q-OPLExTlZ2yumvfr7w=w1920-h955">
</p>

Es izveidoju to pašu trigeri, lai palaistu `clock_out.py`, bet XML failā nomainīju EventID uz 4689.

## Event Viewer
Lai izveidotu iepriekš minēto XML, ir jāatrod divi mainīgie: **"EventID"** un **"ProcessName"**. **"EventID"** — notikuma identifikators. Pasākumiem "Programma ir atvērta" un "Programma ir slēgta" ir savi identifikatori. **"ProcessName"** saglabā ceļu uz šo programmu. Lai uzzinātu šos identifikatorus, vispirms ir jāieslēdz notikumu reģistrēšana. Nospiedu **Win+R** un ievadu **"secpol.msc"**. Mapē **"Local Policies/Audit Policy"** ir jāieslēdz **"Audit process tracking"**.
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYTJgPO7ooYNtcNznSYwoRUcmO1aBJpYQTOa8sIp2uUAAI5t2NKJO9sXdeQdqcXi9LK6f3r5sbuK2wvxHC8iaKGfRgW95A=w1920-h955">
</p>

Pēc tam es atvēru **"Event Viewer"**. Tagad, apskatot visus notikumus, es atvēru un uzreiz aizvēru vēlamo programmu - PhpStorm. Sarakstā parādījās tās atvēršanas un slēgšanas notikumi. Atverot to, redzu vajadzīgo **"EventID"** un **"ProcessName"**. Tagad tos var kopēt un ielīmēt XML. Manā gadījumā programmas atvēršanas un aizvēršanas identifikatori bija attiecīgi 4688 un 4689.
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYTsXPhwDOGl7hZihJRtOibQpJbrDJHohhuX1OVjTtLX12ldUH0CNo7Nc81GzLO6P1MMmwttpV3xXxBUAkSaWOH5gFoF_A=w1920-h955">
</p>

## Excel
Beigās es izveidoju uzdevumu, kas katru mēnesi palaiž `update_excel.py` failu. Šim nolūkam es izveidoju atbilstošo aktivizētāju:
<p align="center"> 
<img src="https://lh3.googleusercontent.com/u/0/drive-viewer/AEYmBYSBCtebAJgw1LG0vk3-M3a8vERo4HT8_gb2Jrso46qBi7YMmcjvzeaFnXbrbl3lFPr564Jz4s5FVXsAkUhu6O2682m_=w1920-h955">
</p>

Un arī iestatījumos ieslēdzu parametru **"Run Task as soon as possibel after a scheduled start is missed"**. Tas ļauj uzdevumu sākt pat tad, ja uzdevuma palaišanas laiks bija palaists garām. Tas tiks palaists kad dators pirmo reizi startēs.
Taču pamanīju, ka šis parametrs ne vienmēr darbojas, tāpēc failiem `clock_in.py` un `clock_out.py` pievienoju papildu pārbaudi, vai nav pienācis jauns mēnesis.

## Secinājums
Galu galā viss darbojas tā, kā biju domājis. Manas stundas tagad tiek automātiski reģistrētas vēlamajā formātā, tiklīdz es sāku vai beidzu darbu. Ja kaut kas tiks pierakstīts nepareizi (piemēram, aizveru PhpStorm nejauši), tad varu laicīgi labot teksta failu manuāli, kā arī apskatīties tā iepriekšējo versiju.  
Katra mēneša sākumā tiks izveidots man nepieciešamais Excel fails, kurā visas manas stundas būs skaisti saplānotas un aprēķinātas.   
Manuprāt, viss darbojas tieši tā, kā man vajag, un es veiksmīgi automatizēju vienu no saviem ikdienas uzdevumiem.