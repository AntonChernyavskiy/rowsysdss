import pandas as pd
import datetime
import os
from winreg import *

# splitFun = "500 m"
# splitCount = 4
cn = "LR kausa izcīņa"
cd = "2024. gada 31. маijs - 01. jūnijs"
prog = ""
# cn = input("compName: ")
# cd = input("compDate: ")
# prog = input("progression: ")

aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
aKey = OpenKey(aReg, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', 0, KEY_READ)
chromePath = QueryValue(aKey, None)
CloseKey(aKey)
CloseKey(aReg)

home_dir = os.path.expanduser("~")
race_docs_dir = os.path.join(home_dir, "Documents", "raceDocs")
heatsheetFILE = os.path.join(race_docs_dir, "r13586-heatsheet.csv")
summaryFILE = os.path.join(race_docs_dir, "r13586.csv")
raceSuffixesFILE = os.path.join(race_docs_dir, "compInfo/race_suffixes.csv")
eventNumFILE = os.path.join(race_docs_dir, "compInfo/event_num.csv")
boatSuffixesFILE = os.path.join(race_docs_dir, "compInfo/boat_suffixes.csv")

df = pd.read_csv(heatsheetFILE, skip_blank_lines=True, na_filter=True)
df = df[df["Event"].notna()]
df = df.reset_index()
fl = pd.read_csv(summaryFILE, skip_blank_lines=True, na_filter=True)
fl.dropna(how="all", inplace=True)
fl = fl.fillna("")

sl = pd.read_csv(raceSuffixesFILE, skip_blank_lines=True, na_filter=True)
suf_list = {}
for index, row in sl.iterrows():
    suf_list[row["Shortcut"]] = row["Full"]

cl = pd.read_csv(eventNumFILE, skip_blank_lines=True, na_filter=True)
cat_list = {}
for index, row in cl.iterrows():
    cat_list[row["Category"]] = row["EventNum"]

bl = pd.read_csv(boatSuffixesFILE, skip_blank_lines=True, na_filter=True, encoding='latin1')
boat_list = {}
for index, row in bl.iterrows():
    boat_list[row["Shortcut"]] = row["Full"]

y = ""

for j, i in enumerate(df["Day"]):
    if isinstance(i, str):
        y = i
    df.loc[j, "Day"] = y

f = open("html/res_with_qual.html", "r")
html = f.read()
f.close()

f = open("html/start_lists.html", "r")
start_html = f.read()
f.close()

start = datetime.datetime.now()
print("Start: ", start)
info = []
data = []
start_data = []
st_info = []

selected_event_nums = ['9', '10', '11', '12', '13', '18', '19', '3', '5']
selected_event_nums = sorted(selected_event_nums, key=int)  # Сортируем номера событий по возрастанию

# Отфильтруем информацию только для выбранных событий
info = []
for i, en in enumerate(df["EventNum"]):
    if str(en) in selected_event_nums:
        info.append([en, boat_list[df["Event"][i].split()[1]], df["Event"][i].split()[1], df["Event"][i].split()[2], suf_list[df["Event"][i].split()[2]], df["Day"][i], df["Start"][i], cat_list[df["Event"][i].split()[1]], df["Prog"][i]])

# Отфильтруем данные только для выбранных событий
data = []
start_data = []
for j, en in enumerate(fl["EventNum"]):
    if str(en) in selected_event_nums:
        if fl["Crew"][j] != "Empty":
            data.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], " ", " ", fl["Qual"][j], en])
            start_data.append([str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), en])

print("Filtered info:", info)
print("Filtered data:", data)

current_date = datetime.datetime.now().strftime('%Y-%m-%d')
current_time = datetime.datetime.now().strftime('%H:%M:%S')

html = html.format(compName=cn, compDates=cd, cDate=current_date, cTime=current_time)
start_html = start_html.format(compName=cn, compDates=cd, cDate=current_date, cTime=current_time)

f = open("html/tbody_res_with_qual.txt", "r")
tr = f.read()
f.close()

f = open("html/race_header.txt", "r")
inf = f.read()
f.close()

f = open("html/race_headerFIRST.txt", "r")
infOne = f.read()
f.close()

last = ''
last_id = 0

first_insert = True
for j, a in enumerate(data):
    if a[-1] == last:
        html = html.replace("[rinda]", tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
        last = a[-1]
    else:
        if not first_insert:
            html = html.replace("[rinda]", inf.format(info[last_id - 1][8], current_date, current_time, info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
        else:
            html = html.replace("[rinda]", infOne.format(info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
            first_insert = False

        last = a[-1]
        last_id += 1

html = html.replace("[rinda]", "")

last_prog = ""
for sublist in info:
    last_prog = sublist[-1]

html = html.replace("[prog_sys]", last_prog)

ft = open("html/log.html", "w", encoding='utf-8')
ft.write(html)
ft.close()

f = open("html/start_lists.txt", "r")
tr = f.read()
f.close()

f = open("html/start_header.txt", "r")
inf = f.read()
f.close()

f = open("html/start_headerFIRST.txt", "r")
infOne = f.read()
f.close()

last = ''
last_id = 0

first_insert_start = True

for j, a in enumerate(start_data):
    if a[-1] == last:
        start_html = start_html.replace("[st_rinda]", tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
        last = a[-1]
    else:
        if not first_insert_start:
            start_html = start_html.replace("[st_rinda]", inf.format(info[last_id - 1][8], current_date, current_time, info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
        else:
            start_html = start_html.replace("[st_rinda]", infOne.format(info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
            first_insert_start = False

        last = a[-1]
        last_id += 1

start_html = start_html.replace("[st_rinda]", "")
start_html = start_html.replace("[prog_sys]", last_prog)

ft = open("html/start_log.html", "w", encoding='utf-8')
ft.write(start_html)
ft.close()

absolute_path = os.path.dirname(__file__)

os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/result_list.pdf file:///{absolute_path}/html/log.html")
os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/start_list.pdf file:///{absolute_path}/html/start_log.html")

end = datetime.datetime.now()
print("End: ", end)
print("Total: ", (end - start))

