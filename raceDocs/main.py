import pandas as pd
import datetime
import os
from winreg import *
import shutil

# splitFun = "500 m"
# splitCount = 4
cn = "LR kausa izcīņa"
cd = "2024. gada 31. maijs - 01. jūnijs"
prog = ""
# cn = input("compName: ")
# cd = input("compDate: ")
# prog = input("progression: ")


# def split_time(Adjtime):
#     if not Adjtime:
#         return "0"
#
#     # Преобразование строки в объект времени
#     time_obj = datetime.datetime.strptime(Adjtime, '%M:%S.%f')
#
#     # Деление времени на 4
#     result_time = time_obj + datetime.timedelta(milliseconds=(time_obj.microsecond / 1000) / splitCount)
#
#     # Преобразование объекта времени обратно в строку
#     result_str = result_time.strftime('%M:%S.%f')[:-4]
#
#     return result_str


aReg = ConnectRegistry(None,HKEY_LOCAL_MACHINE)

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
#df = df.dropna(how="all", subset=["Event"])
df = df[df["Event"].notna()]
df = df.reset_index()
#df = df.fillna("")
#print(df["Event"])
fl = pd.read_csv(summaryFILE, skip_blank_lines=True, na_filter = True)
fl.dropna(how="all", inplace=True)
fl = fl.fillna("")

sl = pd.read_csv(raceSuffixesFILE, skip_blank_lines=True, na_filter = True)
suf_list = {}
for index, row in sl.iterrows():
    suf_list[row["Shortcut"]] = row["Full"]

cl = pd.read_csv(eventNumFILE, skip_blank_lines=True, na_filter=True)
cat_list = {}
for index, row in cl.iterrows():
    cat_list[row["Category"]] = row["EventNum"]

#print(cat_list)

bl = pd.read_csv(boatSuffixesFILE, skip_blank_lines=True, na_filter = True, encoding='latin1')
boat_list = {}
for index, row in bl.iterrows():
    boat_list[row["Shortcut"]] = row["Full"]

#print(boat_list)
y = ""

for j,i in enumerate(df["Day"]):
    if isinstance(i,str):
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
#prog = []
data = []
start_data = []
st_info = []


# print(df["EventNum"])
for i,en in enumerate(df["EventNum"]):
    info.append([en, boat_list[df["Event"][i].split()[1]], df["Event"][i].split()[1], df["Event"][i].split()[2], suf_list[df["Event"][i].split()[2]], df["Day"][i], df["Start"][i], cat_list[df["Event"][i].split()[1]]])
    #prog.append([[df["Prog"][i]]])

#print(info)\
# print(df["EventNum"])
for i,en in enumerate(df["EventNum"]):

    # print(i)
    #print(i, df["Event"][i])

    if isinstance(df["Event"][i], str):
        a = df["Event"][i].split()
    else:
        a = ["","",""]
for j, en in enumerate(fl["EventNum"]):
    if fl["Crew"][j] != "Empty": data.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], " ", " ", fl["Qual"][j], en])
    if fl["Crew"][j] != "Empty": start_data.append([str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), en])

html = html.format(compName=cn, compDates=cd,
                   cDate=datetime.datetime.now().strftime('%Y-%m-%d'),
                   cTime=datetime.datetime.now().strftime('%H:%M:%S'))
start_html = start_html.format(compName=cn, compDates=cd,
                   cDate=datetime.datetime.now().strftime('%Y-%m-%d'),
                   cTime=datetime.datetime.now().strftime('%H:%M:%S'))

f = open("html/tbody_res_with_qual.txt", "r")
tr = f.read()
f.close()

f = open("html/race_header.txt", "r")
inf = f.read()
f.close()

# for j, a in enumerate(info):
#     #html = html.replace("[header]", inf.format("boat", a[2], a[1], a[4], a[3], a[0], a[5], a[6], "series")+"\ntest")
#     html = html.replace("[rinda]",inf.format("boat", j, j, j, j, j, j, j, "series")+"\n[rinda]")
#     #print(j, inf.format("boat", a[2], a[1], a[4], a[3], a[0], a[5], a[6], "series"))

last = ''
last_id = 0
# print(info)
for j, a in enumerate(data):
    if a[-1] == last:
        html = html.replace("[rinda]", tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
        last = a[-1]
    else:
        #print(a[-1], last_id, info[last_id])
        html = html.replace("[rinda]", inf.format(info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
        last = a[-1]
        last_id+=1
html = html.replace("[rinda]", "")
# "<p style=\"page-break-before: always;\">&nbsp;</p>" +

ft = open("html/log.html", "w", encoding='utf-8')
ft.write(html)
ft.close()

f = open("html/start_lists.txt", "r")
tr = f.read()
f.close()

f = open("html/start_header.txt", "r")
inf = f.read()
f.close()

last = ''
last_id = 0
for j, a in enumerate(start_data):
    if a[-1] == last:
        start_html = start_html.replace("[st_rinda]", tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
        last = a[-1]
    else:
        start_html = start_html.replace("[st_rinda]", inf.format(info[last_id][7], info[last_id][2], info[last_id][1], info[last_id][4], info[last_id][3], info[last_id][0], info[last_id][5], info[last_id][6]) + tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
        #start_html = start_html.replace("[prog_sys]", ft.format(prog[last_id][0]))
        last = a[-1]
        last_id+=1
start_html = start_html.replace("[st_rinda]", "")

ft = open("html/start_log.html", "w", encoding='utf-8')
ft.write(start_html)
ft.close()

absolute_path = os.path.dirname(__file__)

os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/result_list.pdf file:///{absolute_path}/html/log.html")
os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/start_list.pdf file:///{absolute_path}/html/start_log.html")
# os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={desktop_path}/log.pdf file:///{absolute_path}/html/log.html")
# os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={desktop_path}/start_log.pdf file:///{absolute_path}/html/start_log.html")
end = datetime.datetime.now()
print("End: ", end)
print("Total: ", (end-start))