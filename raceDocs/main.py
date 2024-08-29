import os
import json
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
from kivy.utils import platform

from kivy.uix.filechooser import FileChooserListView
from win32file import GetFileAttributesExW, FILE_ATTRIBUTE_HIDDEN
import pywintypes

def get_documents_path():
    if platform == 'win':
        return os.path.join(os.path.expanduser('~'), 'Documents')
    elif platform == 'linux' or platform == 'macosx':
        return os.path.join(os.path.expanduser('~'), 'Documents')
    else:
        return os.path.expanduser('~')  # На случай, если платформа не Windows, Linux или macOS

def get_default_folder():
    documents_path = get_documents_path()
    default_folder = os.path.join(documents_path, 'rowSysDocs')
    if not os.path.exists(default_folder):
        os.makedirs(default_folder)
    return default_folder
class CustomFileChooser(FileChooserListView):
    def is_hidden(self, fn):
        try:
            return GetFileAttributesExW(fn)[0] & FILE_ATTRIBUTE_HIDDEN
        except pywintypes.error as e:
            # Игнорируем ошибку, возвращаем False для системных файлов
            return False

class DriveButton(Button, ButtonBehavior):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


class EventDetailPopup(Popup):
    def __init__(self, event_data, **kwargs):
        super().__init__(**kwargs)
        self.title = f"{event_data.get('EventNum', '')} - {event_data.get('Event', '')}"
        self.size_hint = (0.8, 0.8)
        self.auto_dismiss = True

        layout = BoxLayout(orientation='vertical')
        table_layout = GridLayout(cols=2, spacing=10, size_hint_y=None)
        table_layout.bind(minimum_height=table_layout.setter('height'))

        # Adding rows for each event detail
        for key, value in event_data.items():
            table_layout.add_widget(Label(text=key, bold=True))
            table_layout.add_widget(Label(text=str(value)))

        scroll_view = ScrollView(size_hint=(1, None), size=(Window.width, Window.height * 0.7))
        scroll_view.add_widget(table_layout)
        layout.add_widget(scroll_view)

        self.content = layout

class RaceApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.shift_down = False
        self.last_active_checkbox = None

        # File where presets will be saved
        self.data_file = 'competition_presets.json'

        # Top Section: Update File Button
        self.file_chooser_btn = Button(text='Update file', size_hint_y=None, height=40)
        self.file_chooser_btn.bind(on_press=self.choose_file)
        self.add_widget(self.file_chooser_btn)

        # Middle Section: ScrollView for Events
        self.event_layout = GridLayout(cols=1, size_hint_y=None)
        self.event_layout.bind(minimum_height=self.event_layout.setter('height'))
        self.scroll_view = ScrollView(size_hint=(1, None), size=(Window.width, Window.height * 0.5))
        self.scroll_view.add_widget(self.event_layout)
        self.add_widget(self.scroll_view)

        # Below ScrollView: Competition Name & Competition Date
        self.comp_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)
        self.comp_name = TextInput(hint_text='Competition Name', multiline=False, size_hint_x=0.5)
        self.comp_date = TextInput(hint_text='Competition Date', multiline=False, size_hint_x=0.5)
        self.comp_layout.add_widget(self.comp_name)
        self.comp_layout.add_widget(self.comp_date)
        self.add_widget(self.comp_layout)

        # Section: Save and Load Buttons in one row
        self.save_load_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40)

        self.save_btn = Button(text='Save Data', size_hint_x=0.5)
        self.save_btn.bind(on_press=self.save_competition_data)
        self.save_load_layout.add_widget(self.save_btn)

        self.load_btn = Button(text='Load Data', size_hint_x=0.5)
        self.load_btn.bind(on_press=self.load_competition_data)
        self.save_load_layout.add_widget(self.load_btn)

        self.add_widget(self.save_load_layout)

        # Bottom Section: Print Files Button
        self.print_btn = Button(text='Print Files', size_hint_y=None, height=40)
        self.print_btn.bind(on_press=self.show_export_popup)
        self.add_widget(self.print_btn)

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
        if key == 304:  # Left Shift key code
            self.shift_down = False

    def save_competition_data(self, instance):
        # Prompt user for a preset name
        content = BoxLayout(orientation='vertical', padding=10)
        content.add_widget(Label(text='Enter preset name:', size_hint_y=None, height=40))
        preset_name_input = TextInput(multiline=False)
        content.add_widget(preset_name_input)
        save_btn = Button(text='Save', size_hint_y=None, height=40)
        content.add_widget(save_btn)

        popup = Popup(title='Save Preset', content=content, size_hint=(0.6, 0.4))

        def on_save(btn_instance):
            preset_name = preset_name_input.text.strip()
            if preset_name:
                data = {
                    'competition_name': self.comp_name.text,
                    'competition_date': self.comp_date.text
                }
                try:
                    if os.path.exists(self.data_file):
                        with open(self.data_file, 'r') as f:
                            presets = json.load(f)
                    else:
                        presets = {}

                    presets[preset_name] = data

                    with open(self.data_file, 'w') as f:
                        json.dump(presets, f)

                    popup.dismiss()
                    success_popup = Popup(title='Success', content=Label(text='Preset saved successfully!'),
                                          size_hint=(0.6, 0.4))
                    success_popup.open()
                except Exception as e:
                    error_popup = Popup(title='Error', content=Label(text=f'Error saving data: {str(e)}'),
                                        size_hint=(0.6, 0.4))
                    error_popup.open()

        save_btn.bind(on_press=on_save)
        popup.open()

    def load_competition_data(self, instance):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r') as f:
                    presets = json.load(f)

                # Main layout for the popup
                content = BoxLayout(orientation='vertical', spacing=10, padding=(10, 20))

                # Spinner for selecting a preset
                content.add_widget(Label(text='Select a preset:', size_hint_y=None, height=30))
                preset_spinner = Spinner(
                    text='Select Preset',
                    values=list(presets.keys()),
                    size_hint_y=None,
                    height=40
                )
                content.add_widget(preset_spinner)

                # Label to preview the selected preset's details
                preset_details_label = Label(text='', size_hint_y=None, height=60, valign='top')
                content.add_widget(preset_details_label)

                # Horizontal layout for the buttons
                button_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50, spacing=20)
                load_btn = Button(text='Load', size_hint_x=0.5)
                cancel_btn = Button(text='Cancel', size_hint_x=0.5)

                button_layout.add_widget(load_btn)
                button_layout.add_widget(cancel_btn)

                content.add_widget(button_layout)

                popup = Popup(title='Load Preset', content=content, size_hint=(0.6, 0.5))

                def on_spinner_select(spinner, text):
                    if text in presets:
                        data = presets[text]
                        comp_name = data.get('competition_name', '')
                        comp_date = data.get('competition_date', '')
                        preset_details_label.text = f"Competition Name: {comp_name}\nCompetition Date: {comp_date}"
                    else:
                        preset_details_label.text = ''

                preset_spinner.bind(text=on_spinner_select)

                def on_load(btn_instance):
                    preset_name = preset_spinner.text
                    if preset_name in presets:
                        data = presets[preset_name]
                        self.comp_name.text = data.get('competition_name', '')
                        self.comp_date.text = data.get('competition_date', '')
                        popup.dismiss()
                        success_popup = Popup(title='Success', content=Label(text='Preset loaded successfully!'),
                                              size_hint=(0.6, 0.4))
                        success_popup.open()

                def on_cancel(btn_instance):
                    popup.dismiss()

                load_btn.bind(on_press=on_load)
                cancel_btn.bind(on_press=on_cancel)

                popup.open()
            except Exception as e:
                error_popup = Popup(title='Error', content=Label(text=f'Error loading data: {str(e)}'),
                                    size_hint=(0.6, 0.4))
                error_popup.open()
        else:
            error_popup = Popup(title='Error', content=Label(text='No saved presets found.'),
                                size_hint=(0.6, 0.4))
            error_popup.open()

    def choose_file(self, instance):
        file_path = "C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv"
        try:
            self.df = pd.read_csv(file_path, skip_blank_lines=True, na_filter=True)
            self.df = self.df[self.df["Event"].notna()]
            self.df = self.df.reset_index()
            self.file_chooser_btn.text = f'Update file'
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
                label.index = idx  # Store index in label
                label.bind(on_touch_down=self.on_label_touch_down)  # Bind touch event
                box.add_widget(label)
                self.event_layout.add_widget(box)
                self.event_checkboxes.append((checkbox, idx))

    def on_label_touch_down(self, label, touch):
        if label.collide_point(*touch.pos):
            self.show_event_details(label.index)

    def show_event_details(self, index):
        try:
            # Загрузка данных из `results.csv`
            results_df = pd.read_csv('C:\\Users\\Admin\\Documents\\raceDocs\\results.csv')
            print("Results DataFrame:\n", results_df)
            print("Unique Events in Results DataFrame:\n", results_df["Event"].unique())

            # Убедитесь, что `results_df` содержит нужные данные
            event_row = self.df.loc[index]
            event_num = event_row.get('EventNum', '')
            event = event_row.get('Event', '')
            title = f"{event_num} - {event}"

            # Извлечение значения после третьего пробела
            event_parts = event.split(' ')
            if len(event_parts) >= 3:
                event_to_search = ' '.join(event_parts[1:]).strip()  # Значение после второго пробела
            else:
                event_to_search = event.strip()

            results_df['Event'] = results_df['Event'].str.strip()
            print("Event Cleaned for Filtering:", event_to_search)

            # Фильтрация данных
            filtered_results_df = results_df[results_df['Event'] == event_to_search]

            # Заменяем NaN на пробелы
            filtered_results_df = filtered_results_df.fillna(' ')

            # Применяем split по точке в столбце `Place` и берем первый элемент
            if 'Place' in filtered_results_df.columns:
                filtered_results_df['Place'] = filtered_results_df['Place'].apply(lambda x: str(x).split('.')[0])

            # Оставляем только нужные столбцы
            columns_to_keep = ['Place', 'Crew', 'Bow', 'Stroke', '500m', '1000m', '1500m', 'RawTime', 'PenaltyCode',
                               'AdjTime', 'Delta', 'Rank', 'Qual']
            filtered_results_df = filtered_results_df[columns_to_keep]

            # Отладка: вывод отфильтрованных данных
            print("Filtered Results DataFrame:\n", filtered_results_df)

            # Создание содержимого для Popup
            content = BoxLayout(orientation='vertical', padding=10)

            # Добавляем заголовок события
            content.add_widget(Label(text=title, font_size='20sp', size_hint_y=None, height=40))

            if filtered_results_df.empty:
                content.add_widget(Label(text='No results found for this event.', size_hint_y=None, height=40))
            else:
                # Создание таблицы
                table_layout = GridLayout(cols=len(filtered_results_df.columns), spacing=10, size_hint_y=None)
                table_layout.bind(minimum_height=table_layout.setter('height'))

                # Добавление заголовков таблицы
                for column in filtered_results_df.columns:
                    header_label = Label(text=column, bold=True, size_hint_y=None, height=40)
                    table_layout.add_widget(header_label)

                # Добавление строк таблицы
                for _, row in filtered_results_df.iterrows():
                    for value in row:
                        cell_label = Label(text=str(value), size_hint_y=None, height=30)
                        table_layout.add_widget(cell_label)

                # Создание ScrollView и добавление таблицы в него
                scroll_view = ScrollView(size_hint=(1, 1))
                scroll_view.add_widget(table_layout)
                content.add_widget(scroll_view)

            # Открытие Popup с таблицей
            popup = Popup(title='Event Results', content=content, size_hint=(0.9, 0.9))
            popup.open()
        except Exception as e:
            # Обработка ошибок
            popup = Popup(title='Error', content=Label(text=f'Error displaying event details: {str(e)}'),
                          size_hint=(0.6, 0.4))
            popup.open()

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
        content = BoxLayout(orientation='horizontal')

        # Left side layout
        left_layout = BoxLayout(orientation='vertical', spacing=10, padding=[10, 20, 10, 20], size_hint=(0.7, 1))

        # File name input
        file_name_input = TextInput(hint_text='Enter file name', multiline=False, size_hint_y=None, height=dp(44))
        left_layout.add_widget(file_name_input)

        # File type spinner
        file_spinner = Spinner(
            text='Select File',
            values=('Results', 'Results (no qual)', 'Master results', 'Master results (no qual)', 'Startlists',
                    'Master startlists', 'Short startlists', 'Entry list by event', 'Atlase'),
            size_hint_y=None,
            height=dp(44)
        )
        left_layout.add_widget(file_spinner)

        # Checkbox for deleting intermediate files
        delete_intermediate_cb = CheckBox(size_hint_y=None, height=dp(44))
        delete_intermediate_label = Label(text='Delete intermediate files', size_hint_y=None, height=dp(44))
        delete_intermediate_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))
        delete_intermediate_layout.add_widget(delete_intermediate_cb)
        delete_intermediate_layout.add_widget(delete_intermediate_label)
        left_layout.add_widget(delete_intermediate_layout)

        # Checkbox for opening file in Chrome
        open_in_chrome_cb = CheckBox(size_hint_y=None, height=dp(44))
        open_in_chrome_label = Label(text='Open file in Chrome after printing', size_hint_y=None, height=dp(44))
        open_in_chrome_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))
        open_in_chrome_layout.add_widget(open_in_chrome_cb)
        open_in_chrome_layout.add_widget(open_in_chrome_label)
        left_layout.add_widget(open_in_chrome_layout)

        # Select Drive button
        dropdown = DropDown()
        for drive in ['C:', 'D:', 'E:', 'F:', 'G:']:
            btn = Button(text=drive, size_hint_y=None, height=dp(44))
            btn.bind(on_release=lambda btn: self.select_drive(btn.text, file_chooser))
            dropdown.add_widget(btn)
        drive_button = Button(text='Select Drive', size_hint_y=None, height=dp(44))
        drive_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(drive_button, 'text', x))
        left_layout.add_widget(drive_button)

        # Reset to default path button
        reset_path_button = Button(text='Reset to Default Path', size_hint_y=None, height=dp(44))
        reset_path_button.bind(on_press=lambda x: self.reset_to_default_path(file_chooser, get_default_folder()))
        left_layout.add_widget(reset_path_button)

        # Export button
        export_btn = Button(text='Export', size_hint_y=None, height=dp(44))
        export_btn.bind(on_press=lambda x: self.print_files(file_name_input.text, file_spinner.text, file_chooser.path,
                                                            delete_intermediate_cb.active, open_in_chrome_cb.active))
        left_layout.add_widget(export_btn)

        # File chooser on the right
        default_folder = get_default_folder()
        file_chooser = FileChooserListView(path=default_folder, size_hint=(0.3, 1))

        # Add widgets to content
        content.add_widget(left_layout)
        content.add_widget(file_chooser)

        # Open the popup
        popup = Popup(title='Export File', content=content, size_hint=(0.9, 0.9))
        popup.open()

    def reset_to_default_path(self, file_chooser, default_folder):
        file_chooser.path = default_folder

    def print_files(self, file_name, selected_file, folder_path, delete_intermediate, open_in_chrome):
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

                print(f"Input file path URL: {input_file_path_url}")  # Debugging URL

                cmd = [
                    chrome_path,
                    "--headless",
                    "--disable-gpu",
                    "--no-sandbox",  # Added --no-sandbox for potential fix
                    f"--print-to-pdf={output_file_path}",
                    input_file_path_url
                ]

                # Debug the command
                print(f"Running command: {' '.join(cmd)}")

                result = subprocess.run(cmd, capture_output=True, text=True)

                # Debugging outputs
                print("stdout:", result.stdout)
                print("stderr:", result.stderr)

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

            if delete_intermediate:
                for pdf in temp_pdf_files:
                    os.remove(pdf)
            else:
                print("Intermediate files retained.")

            if open_in_chrome:
                file_to_open = merged_output_path if os.path.exists(merged_output_path) else temp_pdf_files
                if isinstance(file_to_open, list):
                    for pdf in file_to_open:
                        subprocess.Popen([chrome_path, pdf])
                else:
                    subprocess.Popen([chrome_path, file_to_open])

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

        atlase.sort(key=lambda x: x[13] if x[13] is not None else -float('inf'), reverse=True)
        for index, row in enumerate(atlase, start=1):
            row[0] = str(index)

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