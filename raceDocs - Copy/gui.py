import os
import datetime
import pandas as pd
from winreg import *
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView, FileChooserIconView
from kivy.uix.popup import Popup
from kivy.uix.checkbox import CheckBox
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.spinner import Spinner
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.uix.dropdown import DropDown
from kivy.uix.behaviors import ButtonBehavior
from kivy.uix.button import Button


class DriveButton(Button, ButtonBehavior):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


class RaceApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.shift_down = False
        self.last_active_checkbox = None

        left_layout = BoxLayout(orientation='vertical', size_hint=(0.5, 1))

        self.comp_name = TextInput(hint_text='Competition Name', multiline=False)
        left_layout.add_widget(self.comp_name)

        self.comp_date = TextInput(hint_text='Competition Date', multiline=False)
        left_layout.add_widget(self.comp_date)

        self.file_chooser_btn = Button(text='Choose file')
        self.file_chooser_btn.bind(on_press=self.choose_file)
        left_layout.add_widget(self.file_chooser_btn)

        self.print_btn = Button(text='Print Files')
        self.print_btn.bind(on_press=self.show_export_popup)
        left_layout.add_widget(self.print_btn)

        self.add_widget(left_layout)

        # Right side layout for displaying events
        self.right_layout = BoxLayout(orientation='vertical', size_hint=(0.5, 1))

        self.event_layout = GridLayout(cols=1, size_hint_y=None)
        self.event_layout.bind(minimum_height=self.event_layout.setter('height'))
        self.scroll_view = ScrollView(size_hint=(1, 1))
        self.scroll_view.add_widget(self.event_layout)
        self.right_layout.add_widget(self.scroll_view)

        self.add_widget(self.right_layout)

        self.file_path = ''
        self.df = None
        self.selected_events = []
        self.event_checkboxes = []

        Window.bind(on_key_down=self.shift_pressed)
        Window.bind(on_key_up=self.shift_unpressed)

    def shift_pressed(self, window, key, scancode, codepoint, modifier):
        if 'shift' in modifier:
            self.shift_down = True

    def shift_unpressed(self, window, key, scancode):
        if key == 304:
            self.shift_down = False

    def choose_file(self, instance):
        content = BoxLayout(orientation='vertical')

        file_chooser = FileChooserListView()
        content.add_widget(file_chooser)

        dropdown = DropDown()

        for drive in ['C:', 'D:', 'E:', 'F:', 'G:']:
            btn = DriveButton(text=drive, size_hint_y=None, height=dp(44))
            btn.bind(on_release=lambda btn: self.select_drive(btn.text, file_chooser))
            dropdown.add_widget(btn)

        drive_button = Button(text='Select Drive', size_hint=(1, None), height=dp(44))
        drive_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(drive_button, 'text', x))
        content.add_widget(drive_button)

        file_chooser.bind(on_submit=self.readFile)

        popup = Popup(title='Choose file', content=content, size_hint=(0.9, 0.9))
        popup.open()

    def readFile(self, filechooser, selection, touch):
        try:
            self.file_path = selection[0]
            self.df = pd.read_csv(self.file_path, skip_blank_lines=True, na_filter=True)
            self.df = self.df[self.df["Event"].notna()]
            self.df = self.df.reset_index()
            self.file_chooser_btn.text = f'File: {os.path.basename(self.file_path)}'
            self.showEvents()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error loading file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()

    def showEvents(self):
        self.event_layout.clear_widgets()
        self.event_checkboxes.clear()
        if self.df is not None:
            for idx, row in self.df.iterrows():
                event_num = row.get('EventNum', '')
                event = row.get('Event', '')
                display_text = f"{event_num} - {event}" if event_num and event else str(event or event_num)
                box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30)
                checkbox = CheckBox()
                checkbox.bind(active=self.eventSelection)
                box.add_widget(checkbox)
                label = Label(text=display_text, size_hint_x=0.8, halign='center', valign='middle')
                label.bind(size=label.setter('text_size'))
                box.add_widget(label)
                self.event_layout.add_widget(box)
                self.event_checkboxes.append((checkbox, idx))

    def eventSelection(self, checkbox, value):
        index = next(idx for idx, (cb, _) in enumerate(self.event_checkboxes) if cb == checkbox)
        if self.shift_down and self.last_active_checkbox:
            current_index = index
            last_index = next(
                idx for idx, (cb, _) in enumerate(self.event_checkboxes) if cb == self.last_active_checkbox)
            start, end = sorted((current_index, last_index))
            for idx in range(start, end + 1):
                cb, event_idx = self.event_checkboxes[idx]
                cb.active = value
                if value:
                    if event_idx not in self.selected_events:
                        self.selected_events.append(event_idx)
                else:
                    if event_idx in self.selected_events:
                        self.selected_events.remove(event_idx)
        else:
            if value:
                if index not in self.selected_events:
                    self.selected_events.append(index)
            else:
                if index in self.selected_events:
                    self.selected_events.remove(index)
            self.last_active_checkbox = checkbox if value else None

    def show_export_popup(self, instance):
        content = BoxLayout(orientation='vertical')

        file_name_input = TextInput(hint_text='Enter file name', multiline=False)
        content.add_widget(file_name_input)

        file_spinner = Spinner(
            text='Select File',
            values=('Results', 'Startlists'),
            size_hint=(1, None),
            height=dp(44)
        )
        content.add_widget(file_spinner)

        file_chooser = FileChooserListView()
        content.add_widget(file_chooser)

        export_btn = Button(text='Export', size_hint=(1, 0.2), height=dp(44))
        export_btn.bind(on_press=lambda x: self.print_files(file_name_input.text, file_spinner.text, file_chooser.path))
        content.add_widget(export_btn)

        dropdown = DropDown()

        for drive in ['C:', 'D:', 'E:', 'F:', 'G:']:
            btn = DriveButton(text=drive, size_hint_y=None, height=dp(44))
            btn.bind(on_release=lambda btn: self.select_drive(btn.text, file_chooser))
            dropdown.add_widget(btn)

        drive_button = Button(text='Select Drive', size_hint=(1, None), height=dp(44))
        drive_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(drive_button, 'text', x))
        content.add_widget(drive_button)

        popup = Popup(title='Export File', content=content, size_hint=(0.8, 0.8))
        popup.open()

    def print_files(self, file_name, selected_file, folder_path):
        try:
            if not self.selected_events:
                raise ValueError("No events selected")

            if not folder_path:
                raise ValueError("No folder selected")

            selected_event_nums = [str(self.df.iloc[idx]['EventNum']) for idx in self.selected_events]
            selected_event_nums = sorted(selected_event_nums, key=int)  # Sort event numbers
            print("Selected Event Numbers:", selected_event_nums)

            self.process_files(selected_event_nums)

            file_path = os.path.join(folder_path, f"{file_name}.pdf")

            chrome_path = self.get_chrome_path()
            absolute_path = os.path.dirname(__file__)
            if selected_file == 'Results':
                os.system(
                    f"\"{chrome_path}\" --headless --disable-gpu --print-to-pdf={file_path} file:///{absolute_path}/html/log.html")
            elif selected_file == 'Startlists':
                os.system(
                    f"\"{chrome_path}\" --headless --disable-gpu --print-to-pdf={file_path} file:///{absolute_path}/html/start_log.html")

            popup = Popup(title='Print Complete', content=Label(text=f'Files printed successfully!'),size_hint=(0.6, 0.4))
            popup.open()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error printing file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()

    def select_drive(self, drive, file_chooser):
        file_chooser.path = drive

    def get_chrome_path(self):
        aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
        aKey = OpenKey(aReg, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', 0, KEY_READ)
        chrome_path = QueryValue(aKey, None)
        CloseKey(aKey)
        CloseKey(aReg)
        return chrome_path

    def process_files(self, selected_event_nums):
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

        with open("html/res_with_qual.html", "r") as f:
            html = f.read()

        with open("html/start_lists.html", "r") as f:
            start_html = f.read()

        start = datetime.datetime.now()
        print("Start: ", start)

        info = []
        data = []
        start_data = []

        for i, en in enumerate(df["EventNum"]):
            if str(en) in selected_event_nums:
                info.append(
                    [en, boat_list[df["Event"][i].split()[1]], df["Event"][i].split()[1], df["Event"][i].split()[2],
                     suf_list[df["Event"][i].split()[2]], df["Day"][i], df["Start"][i],
                     cat_list[df["Event"][i].split()[1]], df["Prog"][i]])

        for j, en in enumerate(fl["EventNum"]):
            if str(en) in selected_event_nums:
                if fl["Crew"][j] != "Empty":
                    data.append(
                        [str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j],
                         fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], " ", " ",
                         fl["Qual"][j], en])
                    start_data.append(
                        [str(fl["Bow"][j]).split(sep=".")[0], fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), en])

        current_date = datetime.datetime.now().strftime('%Y-%m-%d')
        current_time = datetime.datetime.now().strftime('%H:%M:%S')

        html = html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                           cTime=current_time)
        start_html = start_html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                       cTime=current_time)

        with open("html/tbody_res_with_qual.txt", "r") as f:
            tr = f.read()

        with open("html/race_header.txt", "r") as f:
            inf = f.read()

        with open("html/race_headerFIRST.txt", "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert = True

        for j, a in enumerate(data):
            if a[-1] == last:
                html = html.replace("[rinda]",
                                    tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda]")
                last = a[-1]
            else:
                if not first_insert:
                    html = html.replace("[rinda]",
                                        inf.format(info[last_id - 1][8], current_date, current_time, info[last_id][7],
                                                   info[last_id][2], info[last_id][1], info[last_id][4],
                                                   info[last_id][3], info[last_id][0], info[last_id][5],
                                                   info[last_id][6]) + tr.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                                 a[6], a[7], a[8]) + "\n[rinda]")
                else:
                    html = html.replace("[rinda]", infOne.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                 info[last_id][4], info[last_id][3], info[last_id][0],
                                                                 info[last_id][5], info[last_id][6]) + tr.format(a[0],
                                                                                                                 a[1],
                                                                                                                 a[2],
                                                                                                                 a[3],
                                                                                                                 a[4],
                                                                                                                 a[5],
                                                                                                                 a[6],
                                                                                                                 a[7],
                                                                                                                 a[
                                                                                                                     8]) + "\n[rinda]")
                    first_insert = False

                last = a[-1]
                last_id += 1
        print(info)
        html = html.replace("[rinda]", "")
        last_prog = ""
        for sublist in info:
            last_prog = sublist[-1]
        html = html.replace("[prog_sys]", last_prog)

        with open("html/log.html", "w", encoding='utf-8') as ft:
            ft.write(html)

        with open("html/start_lists.txt", "r") as f:
            tr = f.read()

        with open("html/start_header.txt", "r") as f:
            inf = f.read()

        with open("html/start_headerFIRST.txt", "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert_start = True

        for j, a in enumerate(start_data):
            if a[-1] == last:
                start_html = start_html.replace("[st_rinda]", tr.format(a[0], a[1], a[2]) + "\n[st_rinda]")
                last = a[-1]
            else:
                if not first_insert_start:
                    start_html = start_html.replace("[st_rinda]",
                                                    inf.format(info[last_id - 1][8], current_date, current_time,
                                                               info[last_id][7], info[last_id][2], info[last_id][1],
                                                               info[last_id][4], info[last_id][3], info[last_id][0],
                                                               info[last_id][5], info[last_id][6]) + tr.format(a[0],
                                                                                                               a[1], a[
                                                                                                                   2]) + "\n[st_rinda]")
                else:
                    start_html = start_html.replace("[st_rinda]",
                                                    infOne.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                  info[last_id][4], info[last_id][3], info[last_id][0],
                                                                  info[last_id][5], info[last_id][6]) + tr.format(a[0],
                                                                                                                  a[1],
                                                                                                                  a[
                                                                                                                      2]) + "\n[st_rinda]")
                    first_insert_start = False

                last = a[-1]
                last_id += 1

        start_html = start_html.replace("[st_rinda]", "")
        start_html = start_html.replace("[prog_sys]", last_prog)

        with open("html/start_log.html", "w", encoding='utf-8') as ft:
            ft.write(start_html)

        end = datetime.datetime.now()
        print("End: ", end)
        print("Total: ", (end - start))


class MyApp(App):
    def build(self):
        return RaceApp()


if __name__ == '__main__':
    MyApp().run()