import os
import subprocess
import datetime
import pandas as pd
import glob
from PyPDF2 import PdfMerger
from winreg import *
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
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

    def get_download_path(self):
        """Returns the default downloads path for linux or windows"""
        if os.name == 'nt':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
            return location
        else:
            return os.path.join(os.path.expanduser('~'), 'downloads')

    def choose_file(self, instance):
        content = BoxLayout(orientation='vertical')

        file_path = "C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv"

        try:
            self.df = pd.read_csv(file_path, skip_blank_lines=True, na_filter=True)
            self.df = self.df[self.df["Event"].notna()]
            self.df = self.df.reset_index()
            self.file_chooser_btn.text = f'File: {os.path.basename(file_path)}'
            self.showEvents()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error loading file: {str(e)}'), size_hint=(0.6, 0.4))
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
            values=('Results', 'Results (no qual)', 'Master results', 'Master results (no qual)', 'Startlists', 'Master startlists', 'Short startlists', 'Entry list by event', 'Atlase'),
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

            chrome_path = self.get_chrome_path()
            absolute_path = os.path.dirname(__file__)
            html_dir = os.path.join(absolute_path, 'html')  # Directory for HTML files

            # Ensure the folder path exists
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # Load event_num.csv and create cat_list
            home_dir = os.path.expanduser("~")
            race_docs_dir = os.path.join(home_dir, "Documents", "raceDocs")
            eventNumFILE = os.path.join(race_docs_dir, "compInfo/event_num.csv")
            cl = pd.read_csv(eventNumFILE, skip_blank_lines=True, na_filter=True)

            cat_list = {}
            for index, row in cl.iterrows():
                cat_list[row["Category"]] = row["EventNum"]

            temp_pdf_files = []
            start_short_log_handled = False
            for ev in selected_event_nums:
                event_row = self.df[self.df['EventNum'] == int(ev)].iloc[0]
                event_name = event_row['Event'].split()[1]  # Extract name after first space
                eta = event_row['Event'].split()[2] if len(event_row['Event'].split()) > 2 else event_row['EventNum']

                # Get event number from category name
                ev_num = cat_list.get(event_name)

                if selected_file == 'Results':
                    html_file = f"log_{ev}.html"
                elif selected_file == 'Results (no qual)':
                    html_file = f"log_noq_{ev}.html"
                elif selected_file == 'Master results':
                    html_file = f"log_mast_{ev}.html"
                elif selected_file == 'Master results (no qual)':
                    html_file = f"log_noq_master_{ev}.html"
                elif selected_file == 'Startlists':
                    html_file = f"start_log_{ev}.html"
                elif selected_file == 'Master startlists':
                    html_file = f"start_log_master_{ev}.html"
                elif selected_file == 'Entry list by event':
                    html_file = f"entry_by_events_log_{ev}.html"
                elif selected_file == 'Short startlists':
                    if not start_short_log_handled:
                        html_file = "start_short_log.html"
                        start_short_log_handled = True
                    else:
                        continue
                elif selected_file == 'Atlase':
                    html_file = f"atlase_{ev}.html"
                else:
                    raise ValueError("Invalid file selection")

                # Формируем имя файла
                output_file_name = f"race-{ev}___event-{ev_num}___name-{event_name}___eta-{eta}.pdf"
                output_file_path = os.path.join(folder_path, output_file_name)
                temp_pdf_files.append(output_file_path)

                print(f"Printing {html_file} to {output_file_path}")

                input_file_path = os.path.join(html_dir, html_file)

                if not os.path.isfile(input_file_path):
                    raise FileNotFoundError(f"{input_file_path} does not exist")

                input_file_path_url = input_file_path.replace('\\', '/').replace(' ', '%20')
                input_file_path_url = f"file:///{input_file_path_url}"

                cmd = [
                    chrome_path,
                    "--headless",
                    "--disable-gpu",
                    f"--print-to-pdf={output_file_path}",
                    input_file_path_url
                ]
                result = subprocess.run(cmd, capture_output=True, text=True)

                if result.returncode != 0:
                    error_msg = f"Command failed with exit code {result.returncode}: {result.stderr}"
                    print(error_msg)
                    raise RuntimeError(error_msg)

            merged_output_path = os.path.join(folder_path, f"{file_name}.pdf")
            merger = PdfMerger()

            for pdf in temp_pdf_files:
                merger.append(pdf)

            merger.write(merged_output_path)
            merger.close()

            html_files = glob.glob(os.path.join(html_dir, '*.html'))
            exceptions = {'res_no_qual.html', 'res_with_qual.html', 'start_lists.html', 'entry_lists.html',
                          'start_lists_short.html', 'atlase.html'}
            for html_file_path in html_files:
                if os.path.basename(html_file_path) not in exceptions:
                    os.remove(html_file_path)

            popup = Popup(title='Print Complete', content=Label(text='Files printed and merged successfully!'),
                          size_hint=(0.6, 0.4))
            popup.open()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error printing file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()
            print(f"Error: {str(e)}")

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
        eventNumFILE = os.path.join(race_docs_dir, "compInfo/event_num.csv")
        heatsheetFILE = os.path.join(race_docs_dir, "heatsheet.csv")
        summaryFILE = os.path.join(race_docs_dir, "results.csv")
        raceSuffixesFILE = os.path.join(race_docs_dir, "compInfo/race_suffixes.csv")

        boatSuffixesFILE = os.path.join(race_docs_dir, "compInfo/boat_suffixes.csv")
        affilationFlagFILE = os.path.join(race_docs_dir, "compInfo/affilation_flags.csv")
        affilationListFILE = os.path.join(race_docs_dir, "compInfo/affilations.csv")
        coachListFile = os.path.join(race_docs_dir, "compInfo/coaches.csv")

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

        coachList = pd.read_csv(coachListFile, skip_blank_lines=True, na_filter=True)
        coach_list = {}
        for index, row in coachList.iterrows():
            coach_list[row["Stroke"]] = row["Coach"]

        al = pd.read_csv(affilationListFILE, skip_blank_lines=True, na_filter=True)
        aff_list = {}
        for index, row in al.iterrows():
            aff_list[row["Shortcut"]] = row["Full"]

        cl = pd.read_csv(eventNumFILE, skip_blank_lines=True, na_filter=True)
        cat_list = {}
        for index, row in cl.iterrows():
            cat_list[row["Category"]] = row["EventNum"]

        bl = pd.read_csv(boatSuffixesFILE, skip_blank_lines=True, na_filter=True, encoding='latin1')
        boat_list = {}
        for index, row in bl.iterrows():
            boat_list[row["Shortcut"]] = row["Full"]

        fll = pd.read_csv(affilationFlagFILE, skip_blank_lines=True, na_filter=True, encoding='latin1')
        flag_list = {}
        for index, row in fll.iterrows():
            flag_list[row["Shortcut"]] = row["Flag"]

        with open("html/res_with_qual.html", "r") as f:
            html = f.read()

        with open("html/res_no_qual.html", "r") as f:
            htmlQ = f.read()

        with open("html/start_lists.html", "r") as f:
            start_html = f.read()

        with open("html/res_with_qual.html", "r") as f:
            htmlMaster = f.read()

        with open("html/res_no_qual.html", "r") as f:
            htmlMasterQ = f.read()

        with open("html/start_lists.html", "r") as f:
            start_htmlMaster = f.read()

        with open("html/entry_lists.html", "r") as f:
            entry_html = f.read()

        with open("html/start_lists_short.html", "r") as f:
            start_short = f.read()

        with open("html/atlase.html", "r") as f:
            atlase_html = f.read()

        start = datetime.datetime.now()
        print("Start: ", start)

        info = []
        infoShort = []
        data = []
        dataQ = []
        atlase = []

        dataMaster = []
        dataMasterQ = []

        start_data = []
        start_data_master = []
        entry_data = []

        for i, en in enumerate(df["EventNum"]):
            if str(en) in selected_event_nums:

                def surnameCheck(input_string):
                    if not input_string or len(input_string) == 0:
                        return ' '
                    # Убираем все числа и закрывающие скобки
                    clean_string = ''.join(c for c in input_string if not c.isdigit() and c not in "())")
                    # Разбиваем строку по пробелам
                    parts = clean_string.split()
                    # Ищем слово, написанное заглавными буквами
                    uppercase_parts = [part for part in parts if part.isupper()]
                    if uppercase_parts:
                        surname = uppercase_parts[0]
                        # Форматируем фамилию так, чтобы первая буква была заглавной, остальные строчными
                        return surname.capitalize()
                    return ' '

                def determine_indices(lane_split):
                    # Проверяем, есть ли в первом элементе числа
                    if lane_split[0] and lane_split[0][0].isdigit():
                        return 1, 2
                    return 0, 1

                info.append([
                    en,
                    boat_list[df["Event"][i].split()[1]],
                    df["Event"][i].split()[1],
                    df["Event"][i].split()[2],
                    suf_list[df["Event"][i].split()[2]],
                    df["Day"][i],
                    df["Start"][i],
                    cat_list[df["Event"][i].split()[1]],
                    df["Prog"][i]
                ])

                def safe_split(value):
                    parts = str(value).split()
                    return parts if len(parts) >= 3 else [None] * 3

                event_split = safe_split(df["Event"][i])
                lane1_split = safe_split(df["Lane 1"][i])
                lane2_split = safe_split(df["Lane 2"][i])
                lane3_split = safe_split(df["Lane 3"][i])
                lane4_split = safe_split(df["Lane 4"][i])
                lane5_split = safe_split(df["Lane 5"][i])
                lane6_split = safe_split(df["Lane 6"][i])

                # Определяем индексы для фамилии и клуба
                lane1_indices = determine_indices(lane1_split)
                lane2_indices = determine_indices(lane2_split)
                lane3_indices = determine_indices(lane3_split)
                lane4_indices = determine_indices(lane4_split)
                lane5_indices = determine_indices(lane5_split)
                lane6_indices = determine_indices(lane6_split)

                infoShort.append([
                    df["EventNum"][i],
                    cat_list.get(event_split[1]),
                    event_split[1],
                    df["Day"][i],
                    df["Start"][i],
                    event_split[2],
                    f'<img src="flags/{flag_list.get(lane1_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane1_split[lane1_indices[0]], " "),
                    surnameCheck(lane1_split[lane1_indices[1]]),
                    f'<img src="flags/{flag_list.get(lane2_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane2_split[lane2_indices[0]], " "),
                    surnameCheck(lane2_split[lane2_indices[1]]),
                    f'<img src="flags/{flag_list.get(lane3_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane3_split[lane3_indices[0]], " "),
                    surnameCheck(lane3_split[lane3_indices[1]]),
                    f'<img src="flags/{flag_list.get(lane4_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane4_split[lane4_indices[0]], " "),
                    surnameCheck(lane4_split[lane4_indices[1]]),
                    f'<img src="flags/{flag_list.get(lane5_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane5_split[lane5_indices[0]], " "),
                    surnameCheck(lane5_split[lane5_indices[1]]),
                    f'<img src="flags/{flag_list.get(lane6_split[0], "none.jpg")}" style="max-width: 6mm; max-height: 6mm">',
                    aff_list.get(lane6_split[lane6_indices[0]], " "),
                    surnameCheck(lane6_split[lane6_indices[1]]),
                    df["Prog"][i],
                    en
                ])

        def format_time(seconds):
            """Format seconds into mm:ss.00."""
            if seconds is None:
                return ""
            minutes = int(seconds // 60)
            seconds = seconds % 60
            return f"{minutes:02}:{seconds:05.2f}"

        def time_to_seconds(time_str):
            """Convert a time string in mm:ss.00 or s.00 format to seconds."""
            if isinstance(time_str, str):
                time_str = time_str.strip()
                parts = time_str.split(':')
                if len(parts) == 2:  # mm:ss.00 format
                    try:
                        minutes, seconds = map(float, parts)
                        return minutes * 60 + seconds
                    except ValueError:
                        print(f"Error converting time: {time_str}")
                        return None
                elif len(parts) == 1:  # s.00 format
                    try:
                        return float(parts[0])
                    except ValueError:
                        print(f"Error converting time: {time_str}")
                        return None
            print(f"Invalid time format: {time_str}")
            return None

        def rank_and_delta(time_column, df, index, event_num):
            """Calculate rank and delta from a given time column and row index."""
            current_time = time_to_seconds(df[time_column][index])
            if current_time is None:
                return " ", " "

            # Collect all times for the same event
            times = [time_to_seconds(df[time_column][i]) for i in df.index
                     if str(df["EventNum"][i]) == event_num and time_to_seconds(df[time_column][i]) is not None]

            if not times:
                return " ", " "

            sorted_times = sorted(set(times))
            print(f"Sorted times for event {event_num} in column {time_column}: {sorted_times}")

            # Calculate rank
            rank = sum(t < current_time for t in sorted_times) + 1
            delta = current_time - min(sorted_times) if sorted_times else None
            delta_str = format_time(delta)
            delta_str = f"+ {delta_str}"

            return f"({rank})", delta_str

        def split_time(prev_split, current_split):
            """Calculate the split time between two segments and return it with a label."""
            prev_seconds = time_to_seconds(prev_split)
            current_seconds = time_to_seconds(current_split)
            if prev_seconds is None or current_seconds is None:
                return " ", " "

            split_seconds = current_seconds - prev_seconds
            split_time_str = format_time(split_seconds)

            return f"500m: {split_time_str}", " "

        # Determine columns to check for missing data
        time_columns = ["500m", "1000m", "1500m"]

        # Collect valid indices for each time column
        valid_indices = {col: [] for col in time_columns}
        for col in time_columns:
            valid_indices[col] = [i for i in fl.index if not pd.isna(fl[col][i])]

        # Determine which columns have data for all participants
        columns_with_data = {col: len(valid_indices[col]) == len(fl.index) for col in time_columns}
        columns_with_data["Finish"] = True  # Always include Finish

        # Process data
        data = []
        for j, en in enumerate(fl["EventNum"]):
            if str(en) in selected_event_nums:
                if fl["Crew"][j] != "Empty":
                    # Check if data is available for all required splits
                    valid = all(fl[col][j] is not None for col in time_columns if columns_with_data[col])
                    if valid:
                        def masterFun(a):
                            try:
                                penalty_code = str(a) if a is not None else ""
                                if penalty_code:
                                    age_part, handicap_part = penalty_code.split("(-")
                                    av_age = age_part.split()[1]
                                    handicap = handicap_part.split(")")[0]
                                    return f'AV AGE: {av_age} / HANDICAP: {handicap}'
                                else:
                                    return " "
                            except (IndexError, ValueError):
                                return " "

                        adj_time = fl["AdjTime"][j]
                        penalty_code = fl["PenaltyCode"][j]
                        adj_time_seconds = time_to_seconds(adj_time)
                        penalty_code_seconds = time_to_seconds(penalty_code)

                        if adj_time_seconds is not None and penalty_code_seconds is not None:
                            modeltime = round((penalty_code_seconds / adj_time_seconds) * 100, 2)
                        else:
                            modeltime = None

                        print(f"Processing row {j}, event number {en}")
                        print(f"500m time: {fl['500m'][j]}")
                        print(f"1000m time: {fl['1000m'][j]}")
                        print(f"1500m time: {fl['1500m'][j]}")

                        m500rank, m500delta = rank_and_delta("500m", fl, j, str(en))
                        m1000rank, m1000delta = rank_and_delta("1000m", fl, j, str(en)) if columns_with_data[
                            "1000m"] else ("", "")
                        m1500rank, m1500delta = rank_and_delta("1500m", fl, j, str(en)) if columns_with_data[
                            "1500m"] else ("", "")

                        m1000split, _ = split_time(fl["500m"][j], fl["1000m"][j]) if columns_with_data["1000m"] else (
                        "500m: ", " ")
                        m1500split, _ = split_time(fl["1000m"][j], fl["1500m"][j]) if columns_with_data["1500m"] else (
                        "500m: ", " ")
                        mFinsplit, _ = split_time(fl["1500m"][j], fl["AdjTime"][j])

                        data.append([str(fl["Place"][j]).split(sep=".")[0],
                            str(fl["Bow"][j]).split(sep=".")[0],
                            f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                            fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["500m"][j])) if columns_with_data["500m"] else "",
                            m500rank, m500delta, format_time(time_to_seconds(fl["1000m"][j])) if columns_with_data["1000m"] else "",
                            m1000rank, m1000delta, m1000split, format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                            m1500rank, m1500delta, m1500split, fl["AdjTime"][j],
                            f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else '', mFinsplit, fl["Qual"][j], coach_list.get(fl["Stroke"][j]), en])

                        dataQ.append([str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                                        str(fl["Bow"][j]).split(sep=".")[0],
                                        f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                        fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["500m"][j])) if columns_with_data["500m"] else "",
                                        m500rank, m500delta, format_time(time_to_seconds(fl["1000m"][j])) if columns_with_data["1000m"] else "",
                                        m1000rank, m1000delta, m1000split, format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                        m1500rank, m1500delta, m1500split, fl["AdjTime"][j],
                                        f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else '', mFinsplit, coach_list.get(fl["Stroke"][j]), en])

                        atlase.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0],
                                       f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                       fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j],
                                       fl["Delta"][j], " ",
                                       " ", fl["Qual"][j], " ", " ",
                                       " ", modeltime, " ", " ",
                                       " ", en])

                        dataMaster.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0],
                                           f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                           fl["Crew"][j],
                                           fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                           m1500rank, m1500delta, fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                                           f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else '',
                                           fl["Qual"][j], masterFun(str(fl["PenaltyCode"][j])), en])

                        dataMasterQ.append([str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                             str(fl["Bow"][j]).split(sep=".")[0],
                             f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                             fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                             m1500rank, m1500delta, fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                             f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else '', masterFun(str(fl["PenaltyCode"][j])), en])

                        start_data.append([str(fl["Bow"][j]).split(sep=".")[0],
                                           f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                           fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), coach_list.get(fl["Stroke"][j]), en])

                        start_data_master.append([str(fl["Bow"][j]).split(sep=".")[0],
                                           f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                           fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                           masterFun(str(fl["PenaltyCode"][j])), coach_list.get(fl["Stroke"][j]), en])

                        entry_data.append([f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                              fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), en])

        start_data.sort(key=lambda x: (int(x[5]), int(x[0])))
        start_data_master.sort(key=lambda x: (int(x[6]), int(x[0])))
        current_date = datetime.datetime.now().strftime('%Y-%m-%d')
        current_time = datetime.datetime.now().strftime('%H:%M:%S')

        html = html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                           cTime=current_time)
        htmlQ = htmlQ.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                           cTime=current_time)
        htmlMaster = htmlMaster.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                           cTime=current_time)
        htmlMasterQ = htmlMasterQ.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                             cTime=current_time)
        start_html = start_html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                       cTime=current_time)
        start_htmlMaster = start_htmlMaster.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                       cTime=current_time)
        entry_html = entry_html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                                   cTime=current_time)
        start_short = start_short.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                       cTime=current_time)
        atlase_html = atlase_html.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                         cDate=current_date, cTime=current_time)

        html_dir = os.path.join(os.path.dirname(__file__), 'html')

        with open(os.path.join(html_dir, "tbody_res_with_qual.txt"), "r") as f:
            tr = f.read()

        with open(os.path.join(html_dir, "race_header.txt"), "r") as f:
            inf = f.read()

        with open(os.path.join(html_dir, "race_headerFIRST.txt"), "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert = True
        for j, a in enumerate(data):
            if a[-1] == last:
                html = html.replace("[rinda]",
                                    tr.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                last = a[-1]
            else:
                if not first_insert:
                    html = html.replace("[rinda]", "")
                    last_prog = info[last_id-1][8]
                    html = html.replace("[prog_sys]", last_prog)
                    html = html.replace("[prog_sys]", "")
                    with open(os.path.join(html_dir, f"log_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(html)

                    with open(os.path.join(html_dir, "res_with_qual.html"), "r") as f:
                        html = f.read()

                    html = html.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                       cTime=current_time)

                    with open(os.path.join(html_dir, "tbody_res_with_qual.txt"), "r") as f:
                        tr = f.read()

                    with open(os.path.join(html_dir, "race_header.txt"), "r") as f:
                        inf = f.read()

                    with open(os.path.join(html_dir, "race_headerFIRST.txt"), "r") as f:
                        infOne = f.read()

                    html = html.replace("[rinda]", tr.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                             a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                    html = html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                  info[last_id][4], info[last_id][3], info[last_id][0],
                                                                  info[last_id][5], info[last_id][6]))
                else:
                    html = html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                  info[last_id][4], info[last_id][3], info[last_id][0],
                                                                  info[last_id][5], info[last_id][6]))
                    html = html.replace("[rinda]", tr.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                             a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                    first_insert = False

                last = a[-1]
                last_id += 1
        html = html.replace("[rinda]", "")
        last_prog = info[last_id-1][8]
        html = html.replace("[prog_sys]", last_prog)
        html = html.replace("[prog_sys]", "")
        with open(os.path.join(html_dir, f"log_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(html)
#--------------------------------------------------------------------------------------
        with open(os.path.join(html_dir, "tbody_res_with_qual_masters.txt"), "r") as f:
            trMast = f.read()

        with open(os.path.join(html_dir, "race_header_master.txt"), "r") as f:
            infMast = f.read()

        with open(os.path.join(html_dir, "race_headerFIRST_master.txt"), "r") as f:
            infOneMast = f.read()

        last = ''
        last_id = 0
        first_insert = True
        for j, a in enumerate(dataMaster):
            if a[-1] == last:
                htmlMaster = htmlMaster.replace("[rinda]",
                                                trMast.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8],
                                                              a[9], a[10], a[11], a[12], a[13], a[14]) + "\n[rinda]")
                last = a[-1]
            else:
                if not first_insert:
                    htmlMaster = htmlMaster.replace("[rinda]", "")
                    last_prog = info[last_id-1][8]
                    htmlMaster = htmlMaster.replace("[prog_sys]", last_prog)
                    htmlMaster = htmlMaster.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"log_mast_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(htmlMaster)

                    with open(os.path.join(html_dir, "res_with_qual.html"), "r") as f:
                        htmlMaster = f.read()

                    htmlMaster = htmlMaster.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                   cDate=current_date,
                                                   cTime=current_time)

                    with open(os.path.join(html_dir, "tbody_res_with_qual_masters.txt"), "r") as f:
                        trMast = f.read()

                    with open(os.path.join(html_dir, "race_header_master.txt"), "r") as f:
                        infMast = f.read()

                    with open(os.path.join(html_dir, "race_headerFIRST_master.txt"), "r") as f:
                        infOneMast = f.read()

                    htmlMaster = htmlMaster.replace("[rinda]", trMast.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                             a[6], a[7], a[8], a[9], a[10],
                                                                             a[11], a[12], a[13], a[14]) + "\n[rinda]")
                    htmlMaster = htmlMaster.replace("[header]", infOneMast.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                else:
                    htmlMaster = htmlMaster.replace("[header]", infOneMast.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                    htmlMaster = htmlMaster.replace("[rinda]", trMast.format(a[0],
                                                                             a[1],
                                                                             a[2],
                                                                             a[3],
                                                                             a[4],
                                                                             a[5],
                                                                             a[6],
                                                                             a[7],
                                                                             a[8], a[9], a[10], a[11], a[12], a[13], a[14]) + "\n[rinda]")
                    first_insert = False

                last = a[-1]
                last_id += 1
        htmlMaster = htmlMaster.replace("[rinda]", "")
        last_prog = info[last_id-1][8]
        htmlMaster = htmlMaster.replace("[prog_sys]", last_prog)
        htmlMaster = htmlMaster.replace("[prog_sys]", "")

        with open(os.path.join(html_dir, f"log_mast_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(htmlMaster)

        # -----------------------------------------------------------------------------------------------
        with open(os.path.join(html_dir, "start_lists_master.txt"), "r") as f:
            tr = f.read()

        with open(os.path.join(html_dir, "start_header_master.txt"), "r") as f:
            inf = f.read()

        with open(os.path.join(html_dir, "start_headerFIRST_master.txt"), "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert_start = True

        for j, a in enumerate(start_data_master):
            if a[-1] == last:
                start_htmlMaster = start_htmlMaster.replace("[st_rinda]",
                                                            tr.format(a[0], a[1], a[2], a[3], a[4], a[5]) + "\n[st_rinda]")
                last = a[-1]

            else:
                if not first_insert_start:
                    start_htmlMaster = start_htmlMaster.replace("[st_rinda]", "")
                    last_prog = info[last_id - 1][8]
                    start_htmlMaster = start_htmlMaster.replace("[prog_sys]", last_prog)
                    start_htmlMaster = start_htmlMaster.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"start_log_master_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(start_htmlMaster)

                    with open(os.path.join(html_dir, "start_lists.html"), "r") as f:
                        start_htmlMaster = f.read()

                    start_htmlMaster = start_htmlMaster.format(compName=self.comp_name.text,
                                                               compDates=self.comp_date.text,
                                                               cDate=current_date, cTime=current_time)

                    with open(os.path.join(html_dir, "start_lists_master.txt"), "r") as f:
                        tr = f.read()

                    with open(os.path.join(html_dir, "start_header_master.txt"), "r") as f:
                        inf = f.read()

                    with open(os.path.join(html_dir, "start_headerFIRST_master.txt"), "r") as f:
                        infOne = f.read()

                    start_htmlMaster = start_htmlMaster.replace("[st_rinda]",
                                                                tr.format(a[0], a[1], a[2], a[3],
                                                                          a[4], a[5]) + "\n[st_rinda]")
                    start_htmlMaster = start_htmlMaster.replace("[header]",
                                                                infOne.format(info[last_id][7], info[last_id][2],
                                                                              info[last_id][1],
                                                                              info[last_id][4], info[last_id][3],
                                                                              info[last_id][0],
                                                                              info[last_id][5], info[last_id][6]))
                else:
                    start_htmlMaster = start_htmlMaster.replace("[st_rinda]", tr.format(a[0], a[1], a[2], a[3],
                                                                                        a[4], a[5]) + "\n[st_rinda]")
                    start_htmlMaster = start_htmlMaster.replace("[header]",
                                                                infOne.format(info[last_id][7], info[last_id][2],
                                                                              info[last_id][1],
                                                                              info[last_id][4], info[last_id][3],
                                                                              info[last_id][0],
                                                                              info[last_id][5], info[last_id][6]))
                    first_insert_start = False

                last = a[-1]
                last_id += 1

        start_htmlMaster = start_htmlMaster.replace("[st_rinda]", "")
        last_prog = info[last_id - 1][8]
        start_htmlMaster = start_htmlMaster.replace("[prog_sys]", last_prog)
        start_htmlMaster = start_htmlMaster.replace("[prog_sys]", "")
        with open(os.path.join(html_dir, f"start_log_master_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(start_htmlMaster)
        # Start lists (ordinary)
        with open(os.path.join(html_dir, "start_lists.txt"), "r") as f:
            tr = f.read()

        with open(os.path.join(html_dir, "start_header.txt"), "r") as f:
            inf = f.read()

        with open(os.path.join(html_dir, "start_headerFIRST.txt"), "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert_start = True

        for j, a in enumerate(start_data):
            if a[-1] == last:
                start_html = start_html.replace("[st_rinda]", tr.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                last = a[-1]
            else:
                if not first_insert_start:
                    start_html = start_html.replace("[st_rinda]", "")
                    last_prog = info[last_id - 1][8]
                    start_html = start_html.replace("[prog_sys]", last_prog)
                    start_html = start_html.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"start_log_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(start_html)

                    with open(os.path.join(html_dir, "start_lists.html"), "r") as f:
                        start_html = f.read()

                    start_html = start_html.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                   cDate=current_date,
                                                   cTime=current_time)

                    with open(os.path.join(html_dir, "start_lists.txt"), "r") as f:
                        tr = f.read()

                    with open(os.path.join(html_dir, "start_header.txt"), "r") as f:
                        inf = f.read()

                    with open(os.path.join(html_dir, "start_headerFIRST.txt"), "r") as f:
                        infOne = f.read()

                    start_html = start_html.replace("[st_rinda]",
                                                    tr.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                    start_html = start_html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2],
                                                                              info[last_id][1],
                                                                              info[last_id][4], info[last_id][3],
                                                                              info[last_id][0],
                                                                              info[last_id][5], info[last_id][6]))
                else:
                    start_html = start_html.replace("[st_rinda]",
                                                    tr.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                    start_html = start_html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2],
                                                                              info[last_id][1],
                                                                              info[last_id][4], info[last_id][3],
                                                                              info[last_id][0],
                                                                              info[last_id][5], info[last_id][6]))
                    first_insert_start = False

                last = a[-1]
                last_id += 1

        start_html = start_html.replace("[st_rinda]", "")
        last_prog = info[last_id - 1][8]
        start_html = start_html.replace("[prog_sys]", last_prog)
        start_html = start_html.replace("[prog_sys]", "")
        with open(os.path.join(html_dir, f"start_log_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(start_html)

#------------------------------------------------------------------------------

        with open(os.path.join(html_dir, "entry_lists.txt"), "r") as f:
            tr = f.read()

        with open(os.path.join(html_dir, "entry_lists_header.txt"), "r") as f:
            infOne = f.read()

        last = ''
        last_id = 0
        first_insert_start = True

        for j, a in enumerate(entry_data):
            if a[-1] == last:
                entry_html = entry_html.replace("[entry_rinda]",
                                                tr.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                last = a[-1]
            else:
                if not first_insert_start:
                    entry_html = entry_html.replace("[entry_rinda]", "")

                    with open(os.path.join(html_dir, f"entry_by_events_log_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(entry_html)

                    with open(os.path.join(html_dir, "entry_lists.html"), "r") as f:
                        entry_html = f.read()

                    entry_html = entry_html.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                   cDate=current_date,
                                                   cTime=current_time)

                    with open(os.path.join(html_dir, "entry_lists.txt"), "r") as f:
                        tr = f.read()

                    with open(os.path.join(html_dir, "entry_lists_header.txt"), "r") as f:
                        infOne = f.read()

                    entry_html = entry_html.replace("[entry_rinda]",
                                                    tr.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                    entry_html = entry_html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2],
                                                                                    info[last_id][1]))
                else:
                    entry_html = entry_html.replace("[entry_rinda]",
                                                    tr.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                    entry_html = entry_html.replace("[header]", infOne.format(info[last_id][7], info[last_id][2],
                                                                                    info[last_id][1]))
                    first_insert_start = False

                last = a[-1]
                last_id += 1

        entry_html = entry_html.replace("[entry_rinda]", "")
        with open(os.path.join(html_dir, f"entry_by_events_log_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(entry_html)

# ------------------------------------------------------------------------------
        with open(os.path.join(html_dir, "tbody_res_no_qual.txt"), "r") as f:
            trQ = f.read()

        with open(os.path.join(html_dir, "race_header_noq.txt"), "r") as f:
            infQ = f.read()

        with open(os.path.join(html_dir, "race_headerFIRST_noq.txt"), "r") as f:
            infOneQ = f.read()

        last = ''
        last_id = 0
        first_insert = True

        for j, a in enumerate(dataQ):
            if a[-1] == last:
                htmlQ = htmlQ.replace("[rinda_noq]",
                                      trQ.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                             a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                last = a[-1]
            else:
                if not first_insert:
                    htmlQ = htmlQ.replace("[rinda_noq]", "")
                    htmlQ = htmlQ.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"log_noq_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(htmlQ)

                    with open(os.path.join(html_dir, "res_no_qual.html"), "r") as f:
                        htmlQ = f.read()

                    htmlQ = htmlQ.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                         cDate=current_date,
                                         cTime=current_time)

                    with open(os.path.join(html_dir, "tbody_res_no_qual.txt"), "r") as f:
                        trQ = f.read()

                    with open(os.path.join(html_dir, "race_header_noq.txt"), "r") as f:
                        infQ = f.read()

                    with open(os.path.join(html_dir, "race_headerFIRST_noq.txt"), "r") as f:
                        infOneQ = f.read()

                    htmlQ = htmlQ.replace("[rinda_noq]", trQ.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                             a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                    htmlQ = htmlQ.replace("[header]",
                                          infOneQ.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                         info[last_id][4], info[last_id][3], info[last_id][0],
                                                         info[last_id][5], info[last_id][6]))
                else:
                    htmlQ = htmlQ.replace("[header]",
                                          infOneQ.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                         info[last_id][4], info[last_id][3], info[last_id][0],
                                                         info[last_id][5], info[last_id][6]))
                    htmlQ = htmlQ.replace("[rinda_noq]", trQ.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                             a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                    first_insert = False

                last = a[-1]
                last_id += 1

        htmlQ = htmlQ.replace("[rinda_noq]", "")
        htmlQ = htmlQ.replace("[prog_sys]", "")

        with open(os.path.join(html_dir, f"log_noq_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(htmlQ)
# ------------------------------------------------------------------------------
        with open(os.path.join(html_dir, "atlase.txt"), "r") as f:
            trQ = f.read()

        with open(os.path.join(html_dir, "atlase_header.txt"), "r") as f:
            infOneQ = f.read()

        last = ''
        last_id = 0
        first_insert = True

        for j, a in enumerate(atlase):
            if a[-1] == last:
                atlase_html = atlase_html.replace("[rinda_noq]",
                                      trQ.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8],
                                                 a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17]) + "\n[rinda_noq]")
                last = a[-1]
            else:
                if not first_insert:
                    atlase_html = atlase_html.replace("[rinda_noq]", "")
                    atlase_html = atlase_html.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"atlase_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(atlase_html)

                    with open(os.path.join(html_dir, "res_no_qual.html"), "r") as f:
                        atlase_html = f.read()

                    atlase_html = atlase_html.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                         cDate=current_date,
                                         cTime=current_time)

                    with open(os.path.join(html_dir, "atlase.txt"), "r") as f:
                        trQ = f.read()

                    with open(os.path.join(html_dir, "atlase_header.txt"), "r") as f:
                        infOneQ = f.read()

                    atlase_html = atlase_html.replace("[rinda_noq]", trQ.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                    a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17]) + "\n[rinda_noq]")
                    atlase_html = atlase_html.replace("[header]",
                                          infOneQ.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                         info[last_id][4], info[last_id][3], info[last_id][0],
                                                         info[last_id][5], info[last_id][6]))
                else:
                    atlase_html = atlase_html.replace("[header]",
                                          infOneQ.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                         info[last_id][4], info[last_id][3], info[last_id][0],
                                                         info[last_id][5], info[last_id][6]))
                    atlase_html = atlase_html.replace("[rinda_noq]", trQ.format(a[0], a[1], a[2], a[3],
                                                                    a[4], a[5], a[6], a[7], a[8],
                                                                    a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17]) + "\n[rinda_noq]")
                    first_insert = False

                last = a[-1]
                last_id += 1

        atlase_html = atlase_html.replace("[rinda_noq]", "")
        atlase_html = atlase_html.replace("[prog_sys]", "")

        with open(os.path.join(html_dir, f"atlase_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(atlase_html)

# --------------------------------------------------------------------------------------

        with open(os.path.join(html_dir, "tbody_res_no_qual_masters.txt"), "r") as f:
            trMasterQ = f.read()

        with open(os.path.join(html_dir, "race_header_noq_master.txt"), "r") as f:
            infQMast = f.read()

        with open(os.path.join(html_dir, "race_headerFIRST_noq_master.txt"), "r") as f:
            infOneQMast = f.read()

        last = ''
        last_id = 0
        first_insert = True

        for j, a in enumerate(dataMasterQ):
            if a[-1] == last:
                htmlMasterQ = htmlMasterQ.replace("[rinda_noq]",
                                                  trMasterQ.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8],
                                                                   a[9], a[10], a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                last = a[-1]
            else:
                if not first_insert:
                    htmlMasterQ = htmlMasterQ.replace("[rinda_noq]", "")
                    htmlMasterQ = htmlMasterQ.replace("[prog_sys]", "")

                    with open(os.path.join(html_dir, f"log_noq_master_{last}.html"), "w", encoding='utf-8') as ft:
                        ft.write(htmlMasterQ)

                    with open(os.path.join(html_dir, "res_no_qual.html"), "r") as f:
                        htmlMasterQ = f.read()

                    htmlMasterQ = htmlMasterQ.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                     cDate=current_date,
                                                     cTime=current_time)

                    with open(os.path.join(html_dir, "tbody_res_no_qual_masters.txt"), "r") as f:
                        trMasterQ = f.read()

                    with open(os.path.join(html_dir, "race_header_noq_master.txt"), "r") as f:
                        infQMast = f.read()

                    with open(os.path.join(html_dir, "race_headerFIRST_noq_master.txt"), "r") as f:
                        infOneQMast = f.read()

                    htmlMasterQ = htmlMasterQ.replace("[rinda_noq]",
                                                      trMasterQ.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                       a[6], a[7], a[8], a[9], a[10],
                                                                       a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                    htmlMasterQ = htmlMasterQ.replace("[header]", infOneQMast.format(info[last_id][7], info[last_id][2],
                                                                                     info[last_id][1],
                                                                                     info[last_id][4], info[last_id][3],
                                                                                     info[last_id][0],
                                                                                     info[last_id][5],
                                                                                     info[last_id][6]))
                else:
                    htmlMasterQ = htmlMasterQ.replace("[header]", infOneQMast.format(info[last_id][7], info[last_id][2],
                                                                                     info[last_id][1],
                                                                                     info[last_id][4], info[last_id][3],
                                                                                     info[last_id][0],
                                                                                     info[last_id][5],
                                                                                     info[last_id][6]))
                    htmlMasterQ = htmlMasterQ.replace("[rinda_noq]", trMasterQ.format(a[0], a[1], a[2], a[3], a[4],
                                                                                      a[5], a[6], a[7], a[8], a[9],
                                                                                      a[10], a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                    first_insert = False

                last = a[-1]
                last_id += 1

        htmlMasterQ = htmlMasterQ.replace("[rinda_noq]", "")
        htmlMasterQ = htmlMasterQ.replace("[prog_sys]", "")

        with open(os.path.join(html_dir, f"log_noq_master_{last}.html"), "w", encoding='utf-8') as ft:
            ft.write(htmlMasterQ)

# --------------------------------------------------------------------------------------
        with open(os.path.join(html_dir, "start_lists_short.txt"), "r") as f:
            stListSh = f.read()

        last = ''
        last_id = 0
        first_insert = True

        for j, a in enumerate(infoShort):
            if a[-1] == last:
                start_short = start_short.replace("[short_rinda]",
                                                  stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                   a[8],
                                                                   a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                last = a[-1]
            else:
                if not first_insert:
                    start_short = start_short.replace("[short_rinda]",
                                                      stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                   a[8],
                                                                   a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                else:
                    start_short = start_short.replace("[short_rinda]", stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                   a[8],
                                                                   a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                    first_insert = False

                last = a[-1]
                last_id += 1

        start_short = start_short.replace("[short_rinda]", "")
        with open(os.path.join(html_dir, f"start_short_log.html"), "w", encoding='utf-8') as ft:
            ft.write(start_short)

        end = datetime.datetime.now()
        print("End: ", end)
        print("Total: ", (end - start))


class MyApp(App):
    def build(self):
        return RaceApp()


if __name__ == '__main__':
    MyApp().run()