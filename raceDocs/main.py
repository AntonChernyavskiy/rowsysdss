import pandas as pd
import datetime
import os
from winreg import *

aReg = ConnectRegistry(None,HKEY_LOCAL_MACHINE)

aKey = OpenKey(aReg, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', 0, KEY_READ)
chromePath = QueryValue(aKey, None)
CloseKey(aKey)
CloseKey(aReg)

df = pd.read_csv("csv/r13322-heatsheet.csv", skip_blank_lines=True, na_filter=True)
#df = df.dropna(how="all", subset=["Event"])
df = df[df["Event"].notna()]
df = df.reset_index()
#df = df.fillna("")
#print(df["Event"])
fl = pd.read_csv("csv/r13322.csv", skip_blank_lines=True, na_filter = True)
fl.dropna(how="all", inplace=True)
fl = fl.fillna("")

sl = pd.read_csv("race_suffixes.csv", skip_blank_lines=True, na_filter = True)
suf_list = {}
for index, row in sl.iterrows():
    suf_list[row["Shortcut"]] = row["Full"]

bl = pd.read_csv("boat_suffixes.csv", skip_blank_lines=True, na_filter = True, encoding='latin1')
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

cn = input("compName: ")
cd = input("compDate: ")

start = datetime.datetime.now()
print("Start: ", start)
info = []
data = []
start_data = []
st_info = []

for i,en in enumerate(df["EventNum"]):
    #print(i, df["Event"][i])
    info.append([en, boat_list[df["Event"][i].split()[1]], df["Event"][i].split()[1], df["Event"][i].split()[2], suf_list[df["Event"][i].split()[2]], df["Day"][i], df["Start"][i]])

    #print(info)
    if isinstance(df["Event"][i], str):
        a = df["Event"][i].split()
    else:
        a = ["","",""]
    for j, en in enumerate(fl["EventNum"]):
        if fl["Crew"][j] != "Empty": data.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], "split", "split time", "next race", en])
        if fl["Crew"][j] != "Empty": start_data.append([[str(fl["Bow"][j]).split(sep=".")[0]], fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>")])

    html = html.format(compName=cn, compDates=cd,
                       cDate=datetime.datetime.now().strftime('%Y-%m-%d'),
                       cTime=datetime.datetime.now().strftime('%H:%M:%S'))
    # start_html = start_html.format(compName=cn, compDates=cd,
    #                    cDate=datetime.datetime.now().strftime('%Y-%m-%d'),
    #                    cTime=datetime.datetime.now().strftime('%H:%M:%S'), rProgression="N/A")

    f = open("html/tbody_res_with_qual_copy.txt", "r")
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
    for j, a in enumerate(data):
        if a[-1] == last:
            html = html.replace("[rinda]", tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
            last = a[-1]
        else:
            html = html.replace("[rinda]", inf.format("boat", j, j, j, j, j, j, j, "series") + tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
            last = a[-1]
    html = html.replace("[rinda]", "")
    # "<p style=\"page-break-before: always;\">&nbsp;</p>" +

    ft = open("html\log.html", "w", encoding='utf-8')
    ft.write(html)
    ft.close()

    f = open("html/start_lists.txt", "r")
    tr = f.read()
    f.close()

    # last = ''
    # for j, a in enumerate(start_data):
    #     if a[-1] == last:
    #         start_html = start_html.replace("[st_rinda", tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
    #         last = a[-1]
    #     else:
    #         print(":)")
    #         start_html = start_html.replace("[st_rinda", tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
    #         #start_html = start_html.replace("[st_rinda]", tr.format("<p style=\"page-break-before: always;\">&nbsp;</p>" + a[0], a[1], a[2]) + "\n[st_rinda]")
    #         #last = a[-1]
    # start_html = start_html.replace("[st_rinda]", "")
    #
    # ft = open("html\start_log.html", "w", encoding='utf-8')
    # ft.write(start_html)
    # ft.close()

absolute_path = os.path.dirname(__file__)
print(absolute_path)
os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/log.pdf file:///{absolute_path}/html/log.html")
#os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf={absolute_path}/start_log.pdf file:///{absolute_path}/html/start_log.html")
end = datetime.datetime.now()
print("End: ", end)
print("Total: ", (end-start))