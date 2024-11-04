import os
import json
import subprocess
import platform
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
from kivy.uix.widget import Widget

from ftplib import FTP_TLS
import webbrowser
from kivy.uix.filechooser import FileChooserListView
from win32file import GetFileAttributesExW, FILE_ATTRIBUTE_HIDDEN
import pywintypes
import math
import win32print
import win32api

class FileChooserPopup(Popup):
    def __init__(self, is_ftp=False, **kwargs):
        super().__init__(**kwargs)
        self.is_ftp = is_ftp
        self.title = 'Select Files'
        self.size_hint = (0.8, 0.8)
        self.auto_dismiss = True

        self.selected_files = []

        self.default_path = os.path.join(
            os.path.expanduser("~"), 'Documents', 'rowSysDocs'
        )

        if not os.path.exists(self.default_path):
            self.default_path = os.path.expanduser("~Documents")

        layout = BoxLayout(orientation='vertical')

        self.file_chooser = FileChooserListView(
            multiselect=True,
            path=self.default_path,
            size_hint_y=0.9
        )
        layout.add_widget(self.file_chooser)

        select_button = Button(text='Select', size_hint_y=None, height=40)
        select_button.bind(on_press=self.on_file_selected)
        layout.add_widget(select_button)

        self.content = layout

    def on_file_selected(self, instance):
        self.selected_files = self.file_chooser.selection
        self.dismiss()

def get_documents_path():
    if platform == 'win':
        return os.path.join(os.path.expanduser('~'), 'Documents')
    elif platform == 'linux' or platform == 'macosx':
        return os.path.join(os.path.expanduser('~'), 'Documents')
    else:
        return os.path.expanduser('~')

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

        self.ftp_profiles_file = 'ftp_profiles.json'
        self.data_file = 'competition_data.json'
        self.current_profile = None
        self.ftp_connected = False
        self.current_path = ''
        self.documents_path = os.path.expanduser("~/Documents")
        self.current_local_path = os.path.join(self.documents_path, 'rowSysDocs')

        self.ftp_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)
        self.ftp_input_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)
        self.ftp_server_input = TextInput(hint_text='FTP Server Address', multiline=False, size_hint_x=0.33)
        self.ftp_username_input = TextInput(hint_text='FTP Username', multiline=False, size_hint_x=0.33)
        self.ftp_password_input = TextInput(hint_text='FTP Password', password=True, multiline=False, size_hint_x=0.33)
        self.ftp_input_layout.add_widget(self.ftp_server_input)
        self.ftp_input_layout.add_widget(self.ftp_username_input)
        self.ftp_input_layout.add_widget(self.ftp_password_input)
        self.ftp_layout.add_widget(self.ftp_input_layout)

        self.ftp_profiles_spinner = Spinner(
            text='Profiles',
            values=list(self.load_ftp_profiles().keys()),
            size_hint_x=0.2,
            height=40
        )
        self.ftp_profiles_spinner.bind(text=self.on_profile_selected)
        self.connect_disconnect_btn = Button(text='Connect', size_hint_x=0.2, height=40)
        self.connect_disconnect_btn.bind(on_press=self.connect_disconnect_ftp)

        self.ftp_layout.add_widget(self.ftp_profiles_spinner)
        self.ftp_layout.add_widget(self.connect_disconnect_btn)

        self.add_widget(self.ftp_layout)

        self.space_between_ftp_and_update = Label(size_hint_y=None, height=20)
        self.add_widget(self.space_between_ftp_and_update)

        self.main_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=500)
        self.local_file_list_layout = GridLayout(cols=1, size_hint_y=None)
        self.local_file_list_layout.bind(minimum_height=self.local_file_list_layout.setter('height'))
        self.local_file_scroll_view = ScrollView(size_hint=(0.5, 1))
        self.local_file_scroll_view.add_widget(self.local_file_list_layout)
        self.main_layout.add_widget(self.local_file_scroll_view)

        self.ftp_file_list_layout = GridLayout(cols=1, size_hint_y=None)
        self.ftp_file_list_layout.bind(minimum_height=self.ftp_file_list_layout.setter('height'))
        self.ftp_file_scroll_view = ScrollView(size_hint=(0.5, 1))
        self.ftp_file_scroll_view.add_widget(self.ftp_file_list_layout)
        self.main_layout.add_widget(self.ftp_file_scroll_view)

        self.add_widget(self.main_layout)

        button_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40)
        self.update_results_btn = Button(text='Update Results', size_hint_x=0.5, height=40)
        self.update_results_btn.bind(on_press=self.update_file)
        button_layout.add_widget(self.update_results_btn)

        self.update_heatsheet_btn = Button(text='Update Heatsheet', size_hint_x=0.5, height=40)
        self.update_heatsheet_btn.bind(on_press=self.update_heatsheet)
        button_layout.add_widget(self.update_heatsheet_btn)

        self.add_widget(button_layout)

        self.file_chooser_btn = Button(text='Select File', size_hint_y=None, height=40)
        self.file_chooser_btn.bind(on_press=self.choose_file)
        self.add_widget(self.file_chooser_btn)

        self.event_layout = GridLayout(cols=1, size_hint_y=None)
        self.event_layout.bind(minimum_height=self.event_layout.setter('height'))
        self.event_scroll_view = ScrollView(size_hint=(1, 0.2))
        self.event_scroll_view.add_widget(self.event_layout)
        self.add_widget(self.event_scroll_view)

        self.bottom_buttons_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)

        self.comp_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)
        self.comp_name = TextInput(hint_text='Competition Name', multiline=False, size_hint_x=0.5)
        self.comp_date = TextInput(hint_text='Competition Date', multiline=False, size_hint_x=0.5)
        self.regatta_code = TextInput(hint_text='Ragatta code', multiline=False, size_hint_x=0.5)
        self.comp_layout.add_widget(self.comp_name)
        self.comp_layout.add_widget(self.comp_date)
        self.comp_layout.add_widget(self.regatta_code)

        self.save_load_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50)
        self.save_btn = Button(text='Save Data', size_hint_x=0.5)
        self.save_btn.bind(on_press=self.save_competition_data)
        self.save_load_layout.add_widget(self.save_btn)

        self.load_btn = Button(text='Load Data', size_hint_x=0.5, height=50)
        self.load_btn.bind(on_press=self.load_competition_data)
        self.save_load_layout.add_widget(self.load_btn)

        self.bottom_buttons_layout.add_widget(self.comp_layout)
        self.bottom_buttons_layout.add_widget(self.save_load_layout)

        self.edit_heat = Button(text='Edit heatsheet', size_hint_x=0.5)
        self.edit_heat.bind(on_press=self.edit_heatsheet)
        self.save_load_layout.add_widget(self.edit_heat)

        self.print_btn = Button(text='Print Files', size_hint_x=0.5, height=50)
        self.print_btn.bind(on_press=self.show_export_popup)
        self.bottom_buttons_layout.add_widget(self.print_btn)

        self.add_widget(self.bottom_buttons_layout)

        self.file_path = ''
        self.df = None
        self.selected_events = []
        self.event_checkboxes = []

        Window.bind(on_key_down=self.shift_pressed)
        Window.bind(on_key_up=self.shift_unpressed)

        self.list_local_directory(self.current_local_path)

    def edit_heatsheet(self, instance):
        try:
            self.widget_map = {}
            heatsheet_df = pd.read_csv('C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv')
            title = 'Heatsheet Details'

            content = BoxLayout(orientation='vertical', padding=10)
            content.add_widget(Label(text=title, font_size='20sp', size_hint_y=None, height=40))

            if heatsheet_df.empty:
                content.add_widget(Label(text='No heatsheet data found.', size_hint_y=None, height=40))
            else:
                table_layout = GridLayout(cols=len(heatsheet_df.columns), spacing=10, size_hint_y=None)
                table_layout.bind(minimum_height=table_layout.setter('height'))

                for column in heatsheet_df.columns:
                    header_label = Label(text=column, bold=True, size_hint_y=None, height=40)
                    table_layout.add_widget(header_label)

                def on_focus(instance, value):
                    if not value:
                        row_idx = instance.row_idx
                        col_name = instance.col_name
                        self.widget_map[(row_idx, col_name)] = instance

                for idx, row in heatsheet_df.iterrows():
                    for col in heatsheet_df.columns:
                        cell_value = str(row[col])
                        editable_cell = TextInput(text=cell_value, multiline=False, size_hint_y=None, height=30)
                        editable_cell.row_idx = idx
                        editable_cell.col_name = col
                        editable_cell.bind(focus=on_focus)
                        table_layout.add_widget(editable_cell)
                        self.widget_map[(idx, col)] = editable_cell

                scroll_view = ScrollView(size_hint=(1, 1))
                scroll_view.add_widget(table_layout)
                content.add_widget(scroll_view)

            self.save_button = Button(text='Save Changes', size_hint_y=None, height=40)
            self.save_button.bind(on_press=self.save_heatsheet_changes)
            content.add_widget(self.save_button)

            self.popup = Popup(title='Heatsheet Details', content=content, size_hint=(0.9, 0.9))
            self.popup.open()

        except Exception as e:
            error_popup = Popup(title='Error', content=Label(text=f'Error displaying heatsheet details: {str(e)}'),
                                size_hint=(0.6, 0.4))
            error_popup.open()

    def save_heatsheet_changes(self, instance):
        try:

            heatsheet_df = pd.read_csv('C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv', encoding='utf-8')

            for (row_idx, col_name), widget in self.widget_map.items():
                heatsheet_df.at[row_idx, col_name] = widget.text

            heatsheet_df = heatsheet_df.fillna(' ')

            heatsheet_df.to_csv('C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv', index=False, encoding='utf-8')

            with open('C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv', 'r', encoding='utf-8') as file:
                file_content = file.read()

            file_content = file_content.replace('nan', ' ')

            with open('C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv', 'w', encoding='utf-8') as file:
                file.write(file_content)

            self.popup.dismiss()
        except Exception as e:
            error_popup = Popup(title='Error', content=Label(text=f'Error saving heatsheet changes: {str(e)}'),
                                size_hint=(0.6, 0.4))
            error_popup.open()

    def find_widget_at(self, row_idx, col_name):

        for widget in self.children:
            if isinstance(widget, ScrollView):
                for child in widget.children:
                    if isinstance(child, GridLayout):
                        for subwidget in child.children:
                            if isinstance(subwidget,
                                          TextInput) and subwidget.row_idx == row_idx and subwidget.col_name == col_name:
                                return subwidget
        return None

    def open_file_in_browser(self, file_path, is_local, popup):
        """Open the file in the default web browser."""
        try:
            if is_local:

                file_url = f'file://{file_path}'
            else:

                ftp_server = self.ftp_server_input.text
                file_url = f'http://{ftp_server}/{file_path.lstrip("/")}'

            webbrowser.open(file_url)
            popup.dismiss()
            self.show_popup('Success', f'Opened {file_path} in browser')
        except Exception as e:
            self.show_popup('Error', f'Error opening {file_path} in browser: {str(e)}')

    def upload_to_ftp(self, instance):
        """Handle uploading files from local directory to FTP server."""

        file_chooser = FileChooserPopup()
        file_chooser.bind(on_dismiss=self.upload_files)
        file_chooser.open()

    def upload_files(self, instance):
        """Upload selected files to FTP server."""
        file_paths = instance.selected_files
        if file_paths:
            for file_path in file_paths:
                try:
                    file_name = os.path.basename(file_path)
                    with open(file_path, 'rb') as file:
                        self.ftps.storbinary(f'STOR {file_name}', file)
                    print(f"File {file_name} uploaded successfully.")

                    self.show_popup('File Uploaded', f'File {file_name} uploaded successfully!')
                except Exception as e:
                    print(f"Error uploading file: {file_name}, {str(e)}")

                    self.show_popup('Error', f'Error uploading file: {file_name}, {str(e)}')

    def download_file(self, instance):
        """Download the selected file from the FTP server."""
        file_name = instance.selected_file
        if file_name:
            try:
                file_content = self.download_file_from_ftp(file_name)
                local_file_path = os.path.join(self.documents_path, file_name)
                with open(local_file_path, 'wb') as file:
                    file.write(file_content)
                print(f"File {file_name} downloaded successfully.")

                self.show_popup('File Downloaded', f'File {file_name} downloaded successfully!')
            except Exception as e:
                print(f"Error downloading file: {str(e)}")
                self.show_popup('Error', f'Error downloading file: {str(e)}')

    def download_file_from_ftp(self, file_name):
        """Download a file from FTP server."""
        from io import BytesIO
        if not hasattr(self, 'ftps'):
            raise AttributeError("FTP connection has not been established.")

        buffer = BytesIO()
        self.ftps.retrbinary(f"RETR {file_name}", buffer.write)
        buffer.seek(0)
        return buffer.getvalue()

    def show_popup(self, title, message):
        """Show a popup message."""
        content = BoxLayout(orientation='vertical', padding=20, spacing=10)
        message_label = Label(text=message, font_size='18sp', size_hint_y=None, height=40)
        content.add_widget(message_label)
        popup = Popup(title=title, content=content, size_hint=(0.8, 0.6), auto_dismiss=True)
        popup.open()

    def list_local_directory(self, path):
        """List files and directories in the local directory."""
        self.local_file_list_layout.clear_widgets()
        try:

            if os.path.abspath(path) != os.path.abspath(os.sep):
                btn = Button(text='..', size_hint_y=None, height=40)
                btn.bind(on_press=self.go_to_parent_local_directory)
                self.local_file_list_layout.add_widget(btn)

            files = os.listdir(path)
            for file in files:
                full_path = os.path.join(path, file)
                if os.path.isdir(full_path):
                    btn = Button(text=file, size_hint_y=None, height=40)
                    btn.bind(on_press=self.on_local_directory_selected)
                    self.local_file_list_layout.add_widget(btn)
                else:
                    btn = Button(text=file, size_hint_y=None, height=40)
                    btn.bind(on_press=self.on_local_file_selected)
                    self.local_file_list_layout.add_widget(btn)
        except Exception as e:
            print(f"Error retrieving local directory list: {str(e)}")

    def on_local_file_selected(self, instance):
        """Handle file selection from the local directory."""
        file_name = instance.text
        file_path = os.path.join(self.current_local_path, file_name)

        self.show_file_interaction_popup(file_name, file_path, is_local=True)

    def show_file_interaction_popup(self, file_name, file_path, is_local):
        """Show popup for file interaction: download, delete, rename, move, open in browser."""
        content = BoxLayout(orientation='vertical', padding=20, spacing=10)

        download_btn = Button(text='Download', size_hint_y=None, height=50)
        delete_btn = Button(text='Delete', size_hint_y=None, height=50)
        rename_btn = Button(text='Rename', size_hint_y=None, height=50)
        move_btn = Button(text='Move to Opposite Folder', size_hint_y=None, height=50)
        open_in_browser_btn = Button(text='Open in Browser', size_hint_y=None, height=50)
        show_printer_selection_btn = Button(text='Print', size_hint_y=None, height=50)

        content.add_widget(Label(text=f'File: {file_name}', size_hint_y=None, height=40))
        content.add_widget(download_btn)
        content.add_widget(delete_btn)
        content.add_widget(rename_btn)
        content.add_widget(move_btn)
        content.add_widget(open_in_browser_btn)
        content.add_widget(show_printer_selection_btn)

        popup = Popup(title='File Actions', content=content, size_hint=(0.8, 0.6), auto_dismiss=True)
        popup.open()

        download_btn.bind(on_press=lambda x: self.download_file(file_name, is_local, popup))
        delete_btn.bind(on_press=lambda x: self.delete_file(file_path, is_local, popup))
        rename_btn.bind(on_press=lambda x: self.rename_file(file_path, is_local, popup))
        move_btn.bind(on_press=lambda x: self.move_file(file_name, is_local, popup))
        open_in_browser_btn.bind(on_press=lambda x: self.open_file_in_browser(file_path, is_local, popup))
        show_printer_selection_btn.bind(on_press=lambda x: self.show_printer_selection(file_path))

    def get_printers(self):
        """Получить список доступных принтеров."""
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        printer_names = [printer[2] for printer in printers]
        return printer_names

    def show_printer_selection(self, file_path):
        """Показать всплывающее окно для выбора принтера и инициировать печать."""
        printers = self.get_printers()

        content = BoxLayout(orientation='vertical', padding=20, spacing=10)

        printer_label = Label(text='Select Printer:', size_hint_y=None, height=40)
        content.add_widget(printer_label)

        printer_dropdown = DropDown()
        for printer in printers:
            btn = Button(text=printer, size_hint_y=None, height=40)
            btn.bind(on_release=lambda btn: (printer_dropdown.select(btn.text), printer_dropdown.dismiss()))
            printer_dropdown.add_widget(btn)

        main_button = Button(text='Choose Printer', size_hint_y=None, height=50)
        main_button.bind(on_release=printer_dropdown.open)
        content.add_widget(main_button)

        def on_select(printer_name, *args):
            self.print_file(file_path, printer_name)

        printer_dropdown.bind(on_select=on_select)

        spacer = Widget(size_hint_y=1)
        content.add_widget(spacer)

        popup = Popup(title='Printer Selection', content=content, size_hint=(0.5, 0.5), auto_dismiss=True)
        popup.open()

    def print_file(self, file_path, printer_name):
        """Печать файла на выбранном принтере."""
        try:
            win32api.ShellExecute(0, "print", file_path, f'/d:"{printer_name}"', ".", 0)
            print("Print command sent successfully.")

            os.system("taskkill /im AcroRd32.exe /f")
        except Exception as e:
            print(f"Error printing file: {e}")

    def download_file(self, file_name, is_local, popup):
        """Download the selected file."""
        if is_local:

            self.show_popup('Download', f'File {file_name} downloaded from local storage.')
        else:

            self.download_file_from_ftp(file_name)
            self.show_popup('Download', f'File {file_name} downloaded from FTP.')

        popup.dismiss()

    def delete_file(self, file_path, is_local, popup):
        """Delete the selected file."""
        try:
            if is_local:
                os.remove(file_path)
            else:
                self.ftps.delete(file_path)

            self.show_popup('Deleted', f'File {file_path} deleted successfully.')
        except Exception as e:
            self.show_popup('Error', f'Error deleting file: {str(e)}')

        popup.dismiss()

    def rename_file(self, file_path, is_local, popup):
        """Rename the selected file."""
        content = BoxLayout(orientation='vertical', padding=20, spacing=10)
        new_name_input = TextInput(hint_text='Enter new name', multiline=False)
        rename_btn = Button(text='Rename', size_hint_y=None, height=50)
        content.add_widget(new_name_input)
        content.add_widget(rename_btn)

        rename_popup = Popup(title='Rename File', content=content, size_hint=(0.8, 0.6), auto_dismiss=True)
        rename_popup.open()

        rename_btn.bind(
            on_press=lambda x: self.perform_rename(file_path, new_name_input.text, is_local, rename_popup, popup))

    def perform_rename(self, file_path, new_name, is_local, rename_popup, original_popup):
        """Perform renaming of the file."""
        if not new_name:
            self.show_popup('Error', 'File name cannot be empty.')
            return

        new_path = os.path.join(os.path.dirname(file_path), new_name)

        try:
            if is_local:
                os.rename(file_path, new_path)
            else:

                self.ftps.rename(file_path, new_name)

            self.show_popup('Renamed', f'File renamed to {new_name}.')
            rename_popup.dismiss()
            original_popup.dismiss()
        except Exception as e:
            self.show_popup('Error', f'Error renaming file: {str(e)}')

    def move_file(self, file_name, is_local, popup):
        """Move the file between local and FTP folders."""
        try:
            if is_local:

                file_path = os.path.join(self.current_local_path, file_name)
                with open(file_path, 'rb') as file:
                    self.ftps.storbinary(f'STOR {file_name}', file)
                self.show_popup('Moved', f'File {file_name} moved to FTP.')
            else:

                file_content = self.download_file_from_ftp(file_name)
                local_file_path = os.path.join(self.current_local_path, file_name)
                with open(local_file_path, 'wb') as file:
                    file.write(file_content)
                self.show_popup('Moved', f'File {file_name} moved to local.')

            popup.dismiss()
        except Exception as e:
            self.show_popup('Error', f'Error moving file: {str(e)}')

    def on_local_directory_selected(self, instance):
        selected_dir = instance.text
        if self.current_local_path == self.documents_path:
            self.current_local_path = os.path.join(self.documents_path, selected_dir)
        else:
            self.current_local_path = os.path.join(self.current_local_path, selected_dir)
        self.list_local_directory(self.current_local_path)

    def go_to_parent_local_directory(self, instance):
        if self.current_local_path != self.documents_path:
            parent_path = os.path.dirname(self.current_local_path)
            self.current_local_path = parent_path
            self.list_local_directory(self.current_local_path)

    def connect_disconnect_ftp(self, instance):
        if not self.ftp_connected:
            self.connect_to_ftp()
        else:
            self.disconnect_from_ftp()

    def connect_to_ftp(self):
        server = self.ftp_server_input.text
        username = self.ftp_username_input.text
        password = self.ftp_password_input.text
        if server and username:
            try:
                self.ftps = FTP_TLS(server)
                self.ftps.login(user=username, passwd=password)
                self.ftps.prot_p()
                self.ftp_connected = True
                self.connect_disconnect_btn.text = 'Disconnect'
                self.save_ftp_profile()
                self.current_path = '/'
                self.update_file_list()
                print(f"Connected to FTPS server as {username}")
                popup = Popup(title='Connection Successful', content=Label(text=f'Connected as {username}'),
                              size_hint=(0.6, 0.4))
                popup.open()
            except Exception as e:
                popup = Popup(title='Error', content=Label(text=f'Error connecting to FTPS: {str(e)}'),
                              size_hint=(0.6, 0.4))
                popup.open()
                print(f"Error connecting to FTPS: {str(e)}")
        else:
            popup = Popup(title='Error', content=Label(text='Please enter server address and username'),
                          size_hint=(0.6, 0.4))
            popup.open()

    def update_file(self, instance):
        import requests
        import pandas as pd
        import os

        regatta_code = self.regatta_code.text.strip()

        if not regatta_code:
            error_popup = Popup(title='Error', content=Label(text='Regatta code is empty.'),
                                size_hint=(0.6, 0.4))
            error_popup.open()
            return

        results_url = f'https://www.crewtimer.com/results?asfile=true&regatta=r{regatta_code}'

        documents_folder = os.path.expanduser("~/Documents/raceDocs/")
        temp_results_path = 'temp_results.csv'
        results_path = os.path.join(documents_folder, 'results.csv')

        response = requests.get(results_url)
        if response.status_code == 200:
            with open(temp_results_path, 'wb') as file:
                file.write(response.content)
        else:
            error_popup = Popup(title='Error',
                                content=Label(text=f'Failed to download results. Status code: {response.status_code}'),
                                size_hint=(0.6, 0.4))
            error_popup.open()
            return

        res = pd.read_csv(temp_results_path)

        for col in ["500m", "1000m", "1500m"]:
            if col not in res.columns:
                res[col] = ""

        def format_time_to_hms(time_str):
            try:
                return pd.to_datetime(time_str, format='%H:%M:%S.%f').strftime('%H:%M:%S.%f')[:-4]
            except:
                return ""

        def format_time_to_ms(time_str):
            try:
                return pd.to_datetime(time_str, format='%M:%S.%f').strftime('%M:%S.%f')[:-4]
            except:
                return ""

        for col in ["Start", "Finish"]:
            if col in res.columns:
                res[col] = res[col].apply(format_time_to_hms)

        for col in ["RawTime", "AdjTime", "Delta", "500m", "1000m", "1500m"]:
            if col in res.columns:
                res[col] = res[col].apply(format_time_to_ms)

        res["Qual"] = ""
        res["Rank"] = ""
        res.to_csv(temp_results_path, index=False)

        if os.path.exists(results_path):
            f_res = pd.read_csv(results_path)
            res["Qual"] = f_res["Qual"]
            res["Rank"] = f_res["Rank"]
            res.to_csv(results_path, index=False)

        success_popup = Popup(title='Success',
                              content=Label(text='Results file successfully downloaded and updated.'),
                              size_hint=(0.6, 0.4))
        success_popup.open()

    def update_heatsheet(self, instance):
        import requests

        regatta_code = self.regatta_code.text.strip()

        if not regatta_code:

            error_popup = Popup(title='Error', content=Label(text='Regatta code is empty.'),
                                size_hint=(0.6, 0.4))
            error_popup.open()
            return

        heatsheet_url = f'https://www.crewtimer.com/heatsheet?asfile=true&regatta=r{regatta_code}'

        documents_folder = os.path.expanduser("~/Documents/raceDocs/")
        temp_heatsheet_path = 'temp_heatsheet.csv'
        heatsheet_path = os.path.join(documents_folder, 'heatsheet.csv')

        response = requests.get(heatsheet_url)
        if response.status_code == 200:
            with open(temp_heatsheet_path, 'wb') as file:
                file.write(response.content)
        else:
            error_popup = Popup(title='Error', content=Label(
                text=f'Failed to download heatsheet. Status code: {response.status_code}'),
                                size_hint=(0.6, 0.4))
            error_popup.open()
            return

        heat = pd.read_csv(temp_heatsheet_path)
        heat["Prog"] = ""
        heat = heat.astype({"Prog": "str"})

        index_abbrev = heat[heat.apply(lambda row: row.astype(str).str.contains("Abbrev").any(), axis=1)].index
        if len(index_abbrev) > 0:
            heat = heat[:index_abbrev[0]]

        heat.to_csv(temp_heatsheet_path, index=False)

        if os.path.exists(heatsheet_path):
            f_heat = pd.read_csv(heatsheet_path)
            heat["Prog"] = f_heat["Prog"].astype("str")
            heat.to_csv(heatsheet_path, index=False)

        success_popup = Popup(title='Success',
                              content=Label(text='Heatsheet file successfully downloaded and updated.'),
                              size_hint=(0.6, 0.4))
        success_popup.open()

    def choose_file(self, instance):
        file_path = "C:\\Users\\Admin\\Documents\\raceDocs\\heatsheet.csv"
        try:
            self.df = pd.read_csv(file_path, skip_blank_lines=True, na_filter=True)
            self.df = self.df[self.df["Event"].notna()]
            self.df = self.df.reset_index()
            self.file_chooser_btn.text = f'Select File'
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
                label.index = idx
                label.bind(on_touch_down=self.on_label_touch_down)
                box.add_widget(label)
                self.event_layout.add_widget(box)
                self.event_checkboxes.append((checkbox, idx))

    def on_label_touch_down(self, label, touch):
        if label.collide_point(*touch.pos):
            self.show_event_details(label.index)

    def go_to_parent_directory(self, instance):
        if self.current_path != '/':
            parent_path = '/'.join(self.current_path.rstrip('/').split('/')[:-1])
            if parent_path == '':
                parent_path = '/'
            self.current_path = parent_path
            self.update_file_list()

    def update_file_list(self, *args):
        self.ftp_file_list_layout.clear_widgets()
        try:
            self.ftps.cwd(self.current_path)
            files = self.ftps.nlst()

            if self.current_path != '/':
                btn = Button(text='..', size_hint_y=None, height=40)
                btn.bind(on_press=self.go_to_parent_directory)
                self.ftp_file_list_layout.add_widget(btn)

            for file in files:
                if self.is_directory(file):
                    btn = Button(text=file, size_hint_y=None, height=40)
                    btn.bind(on_press=self.on_directory_selected)
                    self.ftp_file_list_layout.add_widget(btn)
                else:
                    btn = Button(text=file, size_hint_y=None, height=40)
                    btn.bind(on_press=self.on_file_selected)
                    self.ftp_file_list_layout.add_widget(btn)

        except Exception as e:
            print(f"Error retrieving file list: {str(e)}")

    def on_file_selected(self, instance):
        file_name = instance.text
        file_path = self.file_path
        is_local = True

        content = BoxLayout(orientation='vertical', padding=20, spacing=10)

        file_label = Label(text=f'File: {file_name}', font_size='18sp', size_hint_y=None, height=40)
        content.add_widget(file_label)

        download_btn = Button(
            text='Download',
            size_hint_y=None,
            height=50,
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
            font_size='16sp'
        )
        download_btn.bind(on_press=lambda x: self.download_file_popup_action(file_name))
        content.add_widget(download_btn)

        delete_btn = Button(
            text='Delete',
            size_hint_y=None,
            height=50,
            background_color=(0.8, 0.2, 0.2, 1),
            color=(1, 1, 1, 1),
            font_size='16sp'
        )
        delete_btn.bind(on_press=lambda x: self.delete_file_popup_action(file_name))
        content.add_widget(delete_btn)

        rename_btn = Button(
            text='Rename',
            size_hint_y=None,
            height=50,
            background_color=(0.5, 0.5, 0.5, 1),
            color=(1, 1, 1, 1),
            font_size='16sp'
        )
        rename_btn.bind(on_press=lambda x: self.rename_file_popup_action(file_name))
        content.add_widget(rename_btn)

        move_btn = Button(
            text='Move',
            size_hint_y=None,
            height=50,
            background_color=(0.2, 0.8, 0.2, 1),
            color=(1, 1, 1, 1),
            font_size='16sp'
        )
        move_btn.bind(on_press=lambda x: self.move_file_popup_action(file_name))
        content.add_widget(move_btn)

        open_in_browser_btn = Button(
            text='Open in Browser',
            size_hint_y=None,
            height=50,
            background_color=(0.2, 0.6, 0.8, 1),
            color=(1, 1, 1, 1),
            font_size='16sp'
        )
        open_in_browser_btn.bind(on_press=lambda x: self.open_file_in_browser(file_path, is_local, popup))
        content.add_widget(open_in_browser_btn)

        popup = Popup(
            title='File Actions',
            content=content,
            size_hint=(0.8, 0.6),
            auto_dismiss=True
        )
        popup.open()

    def download_file_popup_action(self, file_name):
        """Download the file from the FTP server."""
        try:
            local_file_path = os.path.join(self.documents_path, file_name)
            with open(local_file_path, 'wb') as f:
                self.ftps.retrbinary(f'RETR {file_name}', f.write)
            self.show_popup('Success', f'{file_name} downloaded successfully!')
        except Exception as e:
            self.show_popup('Error', f'Error downloading {file_name}: {str(e)}')

    def delete_file_popup_action(self, file_name):
        """Delete the file from the FTP server."""
        try:
            self.ftps.delete(file_name)
            self.show_popup('Success', f'{file_name} deleted successfully!')
            self.update_file_list()
        except Exception as e:
            self.show_popup('Error', f'Error deleting {file_name}: {str(e)}')

    def rename_file_popup_action(self, file_name):
        """Open a dialog to rename the file."""
        rename_content = BoxLayout(orientation='vertical', padding=20, spacing=10)

        rename_input = TextInput(text=file_name, multiline=False, size_hint_y=None, height=40)
        rename_content.add_widget(rename_input)

        rename_btn = Button(text='Rename', size_hint_y=None, height=50)
        rename_btn.bind(on_press=lambda x: self.perform_file_rename(file_name, rename_input.text))
        rename_content.add_widget(rename_btn)

        rename_popup = Popup(
            title='Rename File',
            content=rename_content,
            size_hint=(0.6, 0.4),
            auto_dismiss=True
        )
        rename_popup.open()

    def perform_file_rename(self, old_name, new_name):
        """Rename the file on the FTP server."""
        try:
            self.ftps.rename(old_name, new_name)
            self.show_popup('Success', f'File renamed to {new_name}')
            self.update_file_list()
        except Exception as e:
            self.show_popup('Error', f'Error renaming file: {str(e)}')

    def move_file_popup_action(self, file_name):
        """Download the file from FTP and move it locally."""
        try:

            self.download_file_popup_action(file_name)
            self.show_popup('Success', f'{file_name} moved successfully!')
            self.update_file_list()
        except Exception as e:
            self.show_popup('Error', f'Error moving file: {str(e)}')

    def open_file(self, file_path):
        import platform
        import subprocess
        import os

        system = platform.system()
        if system == 'Windows':
            os.startfile(file_path)
        elif system == 'Darwin':
            subprocess.call(['open', file_path])
        elif system == 'Linux':
            subprocess.call(['xdg-open', file_path])
        else:
            print("Unsupported OS")

    def is_directory(self, filename):
        return '.' not in filename

    def on_directory_selected(self, instance):
        selected_dir = instance.text
        if self.current_path == '/':
            self.current_path = f'/{selected_dir}'
        else:
            self.current_path = f'{self.current_path}/{selected_dir}'

        self.update_file_list()

    def disconnect_from_ftp(self):
        if self.ftp_connected:
            self.ftps.quit()
            self.ftp_connected = False
            self.connect_disconnect_btn.text = 'Connect'
            print("Disconnected from FTPS server")

    def save_ftp_profiles(self, profiles):
        with open(self.ftp_profiles_file, 'w') as f:
            json.dump(profiles, f)

    def save_ftp_profile(self):
        server = self.ftp_server_input.text
        username = self.ftp_username_input.text
        profiles = self.load_ftp_profiles()
        profiles[username] = {'server': server, 'username': username}
        self.save_ftp_profiles(profiles)
        self.ftp_profiles_spinner.values = list(profiles.keys())
        self.ftp_profiles_spinner.text = username

    def load_ftp_profiles(self):
        try:
            with open(self.ftp_profiles_file, 'r') as f:
                profiles = json.load(f)
            return profiles
        except FileNotFoundError:
            return {}

    def on_profile_selected(self, spinner, text):
        profiles = self.load_ftp_profiles()
        if text in profiles:
            profile = profiles[text]
            self.ftp_server_input.text = profile.get('server', '')
            self.ftp_username_input.text = profile.get('username', '')

    def show_event_details(self, index):
        try:

            results_df = pd.read_csv('C:\\Users\\Admin\\Documents\\raceDocs\\results.csv')
            event_row = self.df.loc[index]
            event_num = event_row.get('EventNum', '')
            event = event_row.get('Event', '')
            title = f"{event_num} - {event}"

            event_parts = event.split(' ')
            if len(event_parts) >= 3:
                event_to_search = ' '.join(event_parts[1:]).strip()
            else:
                event_to_search = event.strip()

            results_df['Event'] = results_df['Event'].str.strip()

            filtered_results_df = results_df[results_df['Event'] == event_to_search].fillna(' ')

            if 'Place' in filtered_results_df.columns:
                filtered_results_df['Place'] = filtered_results_df['Place'].apply(lambda x: str(x).split('.')[0])

            columns_to_keep = ['Place', 'Crew', 'Bow', 'Stroke', '500m', '1000m', '1500m', 'RawTime', 'PenaltyCode',
                               'AdjTime', 'Delta', 'Rank', 'Qual']
            filtered_results_df = filtered_results_df[columns_to_keep]

            original_filtered_df = filtered_results_df.copy()

            changes = []

            content = BoxLayout(orientation='vertical', padding=10)
            content.add_widget(Label(text=title, font_size='20sp', size_hint_y=None, height=40))

            if filtered_results_df.empty:
                content.add_widget(Label(text='No results found for this event.', size_hint_y=None, height=40))
            else:
                table_layout = GridLayout(cols=len(filtered_results_df.columns), spacing=10, size_hint_y=None)
                table_layout.bind(minimum_height=table_layout.setter('height'))

                for column in filtered_results_df.columns:
                    header_label = Label(text=column, bold=True, size_hint_y=None, height=40)
                    table_layout.add_widget(header_label)

                def on_focus(instance, value):
                    if not value:
                        row_idx = instance.row_idx
                        col_name = instance.col_name

                        new_value = instance.text
                        original_value = original_filtered_df.at[row_idx, col_name]
                        if new_value != original_value:

                            changes.append((row_idx, col_name, new_value))
                            filtered_results_df.at[row_idx, col_name] = new_value

                for idx, row in filtered_results_df.iterrows():
                    for col in filtered_results_df.columns:
                        cell_value = str(row[col])
                        editable_cell = TextInput(text=cell_value, multiline=False, size_hint_y=None, height=30)
                        editable_cell.row_idx = idx
                        editable_cell.col_name = col
                        editable_cell.bind(focus=on_focus)
                        table_layout.add_widget(editable_cell)

                scroll_view = ScrollView(size_hint=(1, 1))
                scroll_view.add_widget(table_layout)
                content.add_widget(scroll_view)

            save_button = Button(text='Save Changes', size_hint_y=None, height=40)
            content.add_widget(save_button)

            def save_changes(instance):
                if changes:
                    for row_idx, col_name, new_value in changes:

                        results_df.loc[results_df.index == filtered_results_df.index[row_idx], col_name] = new_value

                    results_df.to_csv('C:\\Users\\Admin\\Documents\\raceDocs\\results.csv', index=False)

                popup.dismiss()

            save_button.bind(on_press=save_changes)

            popup = Popup(title='Event Results', content=content, size_hint=(0.9, 0.9))
            popup.open()

        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error displaying event details: {str(e)}'),
                          size_hint=(0.6, 0.4))
            popup.open()

    def shift_pressed(self, window, key, scancode, codepoint, modifier):
        if 'shift' in modifier:
            self.shift_down = True

    def shift_unpressed(self, window, key, scancode):
        if key == 304:
            self.shift_down = False

    def save_competition_data(self, instance):
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
                    'competition_date': self.comp_date.text,
                    'regatta_code': self.regatta_code.text
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

                    success_popup = Popup(title='Success', content=Label(text='Data saved successfully'),
                                          size_hint=(0.6, 0.4))
                    success_popup.open()
                except Exception as e:
                    popup.dismiss()
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

                content = BoxLayout(orientation='vertical', spacing=10, padding=(10, 20))

                content.add_widget(Label(text='Select a preset:', size_hint_y=None, height=30))
                preset_spinner = Spinner(
                    text='Select Preset',
                    values=list(presets.keys()),
                    size_hint_y=None,
                    height=40
                )
                content.add_widget(preset_spinner)

                preset_details_label = Label(text='', size_hint_y=None, height=60, valign='top')
                content.add_widget(preset_details_label)

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
                        regatta_code = data.get('regatta_code', '')
                        preset_details_label.text = (
                            f"Competition Name: {comp_name}\n"
                            f"Competition Date: {comp_date}\n"
                            f"Regatta Code: {regatta_code}"
                        )
                    else:
                        preset_details_label.text = ''

                preset_spinner.bind(text=on_spinner_select)

                def on_load(btn_instance):
                    preset_name = preset_spinner.text
                    if preset_name in presets:
                        data = presets[preset_name]
                        self.comp_name.text = data.get('competition_name', '')
                        self.comp_date.text = data.get('competition_date', '')
                        self.regatta_code.text = data.get('regatta_code', '')
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

    def display_events(self):
        self.event_layout.clear_widgets()
        if self.df is not None:
            for i, row in self.df.iterrows():
                event_name = row.get('Event Name', f'Event {i + 1}')
                checkbox = CheckBox()
                label = Label(text=event_name, size_hint_y=None, height=40)
                layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40)
                layout.add_widget(checkbox)
                layout.add_widget(label)
                self.event_layout.add_widget(layout)

                self.event_checkboxes.append(checkbox)

                checkbox.bind(active=lambda checkbox, value, idx=i: self.on_checkbox_active(value, idx, checkbox))

    def on_checkbox_active(self, is_active, index, checkbox):
        if is_active:
            if self.shift_down and self.last_active_checkbox is not None:
                last_index = self.event_checkboxes.index(self.last_active_checkbox)
                start = min(last_index, index)
                end = max(last_index, index) + 1
                for i in range(start, end):
                    if not self.event_checkboxes[i].active:
                        self.event_checkboxes[i].active = True
                        self.selected_events.append(self.df.iloc[i])
            else:
                self.selected_events.append(self.df.iloc[index])
            self.last_active_checkbox = checkbox
        else:
            self.selected_events.remove(self.df.iloc[index])

    def show_export_popup(self, instance):
        content = BoxLayout(orientation='horizontal')

        left_layout = BoxLayout(orientation='vertical', spacing=10, padding=[10, 20, 10, 20], size_hint=(0.7, 1))

        file_name_input = TextInput(hint_text='Enter file name', multiline=False, size_hint_y=None, height=dp(44))
        left_layout.add_widget(file_name_input)

        file_spinner = Spinner(
            text='Select File',
            values=('ENTRY LIST', 'MASTER RESULTS WITH NO QUAL', 'MASTER RESULTS WITH NO QUAL (START TIME)',
                    'MASTER RESULTS WITH QUAL', 'MASTER RESULTS WITH QUAL (START TIME)',
                    'MASTER STARTLIST', 'RESULTS WITH NO QUAL', 'RESULTS WITH NO QUAL (START TIME)',
                    'RESULTS WITH QUAL', 'RESULTS WITH QUAL (START TIME)',
                    'SHORT STARTLIST', 'STARTLIST', 'STARTLIST (START TIME)' 'TEST RACE'),
            size_hint_y=None,
            height=dp(44)
        )
        left_layout.add_widget(file_spinner)

        delete_intermediate_cb = CheckBox(size_hint_y=None, height=dp(44))
        delete_intermediate_label = Label(text='Delete intermediate files', size_hint_y=None, height=dp(44))
        delete_intermediate_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))
        delete_intermediate_layout.add_widget(delete_intermediate_cb)
        delete_intermediate_layout.add_widget(delete_intermediate_label)
        left_layout.add_widget(delete_intermediate_layout)

        open_in_chrome_cb = CheckBox(size_hint_y=None, height=dp(44))
        open_in_chrome_label = Label(text='Open file in Chrome after printing', size_hint_y=None, height=dp(44))
        open_in_chrome_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))
        open_in_chrome_layout.add_widget(open_in_chrome_cb)
        open_in_chrome_layout.add_widget(open_in_chrome_label)
        left_layout.add_widget(open_in_chrome_layout)

        combine_races_cb = CheckBox(size_hint_y=None, height=dp(44))
        combine_races_label = Label(text='Combine races', size_hint_y=None, height=dp(44))
        combine_races_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))
        combine_races_layout.add_widget(combine_races_cb)
        combine_races_layout.add_widget(combine_races_label)

        hint_label = Label(
            text='Race combination works only with: Results with qual, Results with qual (start time), Master results with qual, Master results with qual (start time)',
            size_hint_y=None,
            height=dp(44),
            color=(1, 0, 0, 1)
        )
        left_layout.add_widget(combine_races_layout)
        left_layout.add_widget(hint_label)

        dropdown = DropDown()
        for drive in ['C:', 'D:', 'E:', 'F:', 'G:']:
            btn = Button(text=drive, size_hint_y=None, height=dp(44))
            btn.bind(on_release=lambda btn: self.select_drive(btn.text, file_chooser))
            dropdown.add_widget(btn)

        drive_button = Button(text='Select Drive', size_hint_y=None, height=dp(44))
        drive_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(drive_button, 'text', x))
        left_layout.add_widget(drive_button)

        reset_path_button = Button(text='Reset to Default Path', size_hint_y=None, height=dp(44))
        reset_path_button.bind(on_press=lambda x: self.reset_to_default_path(file_chooser, get_default_folder()))
        left_layout.add_widget(reset_path_button)

        button_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(44))

        export_btn = Button(text='Export', size_hint_x=0.5, height=dp(44))
        export_and_print_btn = Button(text='Export and Print', size_hint_x=0.5, height=dp(44))

        button_layout.add_widget(export_btn)
        button_layout.add_widget(export_and_print_btn)

        left_layout.add_widget(button_layout)

        # Bind export button actions
        export_btn.bind(on_press=lambda x: self.export_files(
            file_name_input.text,
            file_spinner.text,
            file_chooser.path,
            delete_intermediate_cb.active,
            open_in_chrome_cb.active,
            combine_races_cb.active
        ))

        export_and_print_btn.bind(on_press=lambda x: self.export_and_print(
            file_name_input.text,
            file_spinner.text,
            file_chooser.path,
            delete_intermediate_cb.active,
            open_in_chrome_cb.active,
            combine_races_cb.active
        ))

        default_folder = get_default_folder()
        file_chooser = FileChooserListView(path=default_folder, size_hint=(0.3, 1))

        content.add_widget(left_layout)
        content.add_widget(file_chooser)

        popup = Popup(title='Export File', content=content, size_hint=(0.9, 0.9))
        popup.open()

    def export_and_print(self, file_name, selected_file, folder_path, delete_intermediate, open_in_chrome,
                         combine_race):
        try:
            # Generate the files first and get the merged output path
            merged_output_path = self.print_files(file_name, selected_file, folder_path, delete_intermediate,
                                                  open_in_chrome, combine_race)

            if os.path.exists(merged_output_path):
                self.show_printer_selection(merged_output_path)
            else:
                raise FileNotFoundError(f"The generated file {merged_output_path} does not exist.")
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()
            print(f"Error: {str(e)}")

    def export_files(self, file_name, selected_file, folder_path, delete_intermediate, open_in_chrome, combine_race):
        # Call the print_files function to generate the files without printing
        self.print_files(file_name, selected_file, folder_path, delete_intermediate, open_in_chrome, combine_race)

    def reset_to_default_path(self, file_chooser, default_folder):
        file_chooser.path = default_folder

    def print_files(self, file_name, selected_file, folder_path, delete_intermediate, open_in_chrome, combine_race):
        global selectedFile, combineRace
        selectedFile = selected_file
        combineRace = combine_race
        try:
            if not self.selected_events:
                raise ValueError("No events selected")

            if not folder_path:
                raise ValueError("No folder selected")

            selected_event_nums = [str(self.df.iloc[idx]['EventNum']) for idx in self.selected_events]
            selected_event_nums = sorted(selected_event_nums, key=int)
            print("Selected Event Numbers:", selected_event_nums)

            self.process_files(selected_event_nums)

            chrome_path = self.get_chrome_path()
            absolute_path = os.path.dirname(__file__)
            html_dir = os.path.join(absolute_path, 'html')

            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            home_dir = os.path.expanduser("~")
            race_docs_dir = os.path.join(home_dir, "Documents", "raceDocs")
            eventNumFILE = os.path.join(race_docs_dir, "compInfo/event_num.csv")
            cl = pd.read_csv(eventNumFILE, skip_blank_lines=True, na_filter=True)

            cat_list = {row["Category"]: row["EventNum"] for index, row in cl.iterrows()}

            temp_pdf_files = []
            html_short_startlists_log_handled = False

            event_loop_range = [selected_event_nums[0]] if combineRace else selected_event_nums

            for ev in event_loop_range:
                event_row = self.df[self.df['EventNum'] == int(ev)].iloc[0]
                event_name = event_row['Event'].split()[1]
                eta = event_row['Event'].split()[2] if len(event_row['Event'].split()) > 2 else event_row['EventNum']
                ev_num = cat_list.get(event_name)

                html_file = self.get_html_file(selected_file, ev)

                output_file_name = f"race-{ev}___event-{ev_num}___name-{event_name}___eta-{eta}.pdf"
                output_file_path = os.path.join(folder_path, output_file_name)
                temp_pdf_files.append(output_file_path)

                print(f"Printing {html_file} to {output_file_path}")

                input_file_path = os.path.join(html_dir, html_file)

                if not os.path.isfile(input_file_path):
                    raise FileNotFoundError(f"{input_file_path} does not exist")

                input_file_path_url = f"file:///{input_file_path.replace('\\', '/').replace(' ', '%20')}"
                print(f"Input file path URL: {input_file_path_url}")

                cmd = [
                    chrome_path,
                    "--headless",
                    "--disable-gpu",
                    "--no-sandbox",
                    f"--print-to-pdf={output_file_path}",
                    input_file_path_url
                ]

                print(f"Running command: {' '.join(cmd)}")
                result = subprocess.run(cmd, capture_output=True, text=True)

                print("stdout:", result.stdout)
                print("stderr:", result.stderr)

                if result.returncode != 0:
                    raise RuntimeError(f"Command failed with exit code {result.returncode}: {result.stderr}")

            # Merge PDFs
            merged_output_path = os.path.join(folder_path, f"{file_name}.pdf")
            merger = PdfMerger()

            for pdf in temp_pdf_files:
                merger.append(pdf)

            merger.write(merged_output_path)
            merger.close()

            # Delete intermediate files if required
            if delete_intermediate:
                for pdf in temp_pdf_files:
                    os.remove(pdf)
            else:
                print("Intermediate files retained.")
                os.remove(merged_output_path)

            if open_in_chrome:
                if os.path.exists(merged_output_path):
                    subprocess.Popen([chrome_path, merged_output_path])
                    print(f"Opened merged PDF: {merged_output_path}")
                elif temp_pdf_files:  # Check if there are any generated PDFs
                    for pdf_file in temp_pdf_files:
                        subprocess.Popen([chrome_path, pdf_file])  # Open each generated PDF
                        print(f"Opened generated PDF: {pdf_file}")
                else:
                    print("No generated PDF files found.")

            # Clean up HTML files
            html_files = glob.glob(os.path.join(html_dir, '*.html'))
            exceptions = {'body_r_noQual.html', 'body_r_withQual.html', 'main_s.html', 'main_entries.html',
                          'main_s_short.html', 'main_atlase.html'}
            for html_file_path in html_files:
                if os.path.basename(html_file_path) not in exceptions:
                    os.remove(html_file_path)

            popup = Popup(title='Print Complete', content=Label(text='Files printed and merged successfully!'),
                          size_hint=(0.6, 0.4))
            popup.open()

            return merged_output_path  # Return the merged PDF path
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error printing file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()
            print(f"Error: {str(e)}")
            return None  # Return None on error

    def get_html_file(self, selected_file, ev):
        if selected_file == 'RESULTS WITH QUAL':
            return f"log_{ev}.html"
        elif selected_file == 'RESULTS WITH NO QUAL':
            return f"log_noq_{ev}.html"
        elif selected_file == 'MASTER RESULTS WITH QUAL':
            return f"log_mast_{ev}.html"
        elif selected_file == 'MASTER RESULTS WITH NO QUAL':
            return f"log_noq_master_{ev}.html"
        elif selected_file == 'STARTLIST':
            return f"start_log_{ev}.html"
        elif selected_file == 'STARTLIST (START TIME)':
            return f"start_log_wthTime_{ev}.html"
        elif selected_file == 'MASTER STARTLIST':
            return f"start_log_master_{ev}.html"
        elif selected_file == 'ENTRY LIST':
            return f"entry_by_events_log_{ev}.html"
        elif selected_file == 'SHORT STARTLIST':
            return "html_short_startlists_log.html"
        elif selected_file == 'TEST RACE':
            return f"atlase_{ev}.html"
        elif selected_file == 'MASTER RESULTS WITH NO QUAL (START TIME)':
            return f"log_noq_master_wthStart_{ev}.html"
        elif selected_file == 'MASTER RESULTS WITH QUAL (START TIME)':
            return f"log_mast_wthStart_{ev}.html"
        elif selected_file == 'RESULTS WITH NO QUAL (START TIME)':
            return f"log_noq_wthStart_{ev}.html"
        elif selected_file == 'RESULTS WITH QUAL (START TIME)':
            return f"log_wthStart_{ev}.html"
        else:
            raise ValueError("Invalid file selection")

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

        start = datetime.datetime.now()
        print("Start: ", start)

        info = []
        infoShort = []
        res_withQual = []
        res_withQual_start = []
        res_noQual = []
        res_noQual_start = []
        atlase = []

        res_withQual_master = []
        res_withQual_master_start = []
        res_noQual_master = []
        res_noQual_master_start = []

        start_data = []
        start_data_wthStart = []
        start_data_master = []
        entry_data = []

        for i, en in enumerate(df["EventNum"]):
            if str(en) in selected_event_nums:

                def surnameCheck(input_string):
                    if not input_string or len(input_string) == 0:
                        return ' '

                    clean_string = ''.join(c for c in input_string if not c.isdigit() and c not in "())")

                    parts = clean_string.split()

                    uppercase_parts = [part for part in parts if part.isupper()]
                    if uppercase_parts:
                        surname = uppercase_parts[0]

                        return surname.capitalize()
                    return ' '

                def determine_indices(lane_split):

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
                    df["Prog"][i] if df["Prog"][i] and not math.isnan(df["Prog"][i]) else " "
                ])
                print(info)
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

                lane1_indices = determine_indices(lane1_split)
                lane2_indices = determine_indices(lane2_split)
                lane3_indices = determine_indices(lane3_split)
                lane4_indices = determine_indices(lane4_split)
                lane5_indices = determine_indices(lane5_split)
                lane6_indices = determine_indices(lane6_split)
                if selectedFile == "SHORT STARTLIST":
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
                        df["Prog"][i] if df["Prog"][i] and not math.isnan(df["Prog"][i]) else " ",
                        en
                    ])

        def combine_first_last(info):
            if len(info) < 2:
                return info

            first = info[0]
            last = info[-1]

            combined_result = [
                [
                    f"{first[0]}-{last[0]}",
                    *first[1:3],
                    first[3][0],
                    *first[4:6],
                    f"{first[6]}-{last[6]}",
                    *first[7:]
                ]
            ]

            return combined_result

        if (combineRace and (selectedFile == "RESULTS WITH QUAL" or selectedFile == "RESULTS WITH QUAL (START TIME)"
                or selectedFile == " MASTER RESULTS WITH QUAL" or selectedFile == "MASTER RESULTS WITH QUAL (START TIME)")):
            info = combine_first_last(info)

        def format_time(seconds):
            """Format time as mm:ss.00 if more than a minute, otherwise as ss.00."""
            if seconds is None:
                return ""
            minutes, sec = divmod(seconds, 60)
            if minutes > 0:
                return f"{int(minutes)}:{sec:05.2f}"
            return f"{sec:05.2f}"

        def time_to_seconds(time_str):
            """Convert a time string in mm:ss.00 or s.00 format to seconds."""
            if isinstance(time_str, str):
                time_str = time_str.strip()
                parts = time_str.split(':')
                if len(parts) == 2:
                    try:
                        minutes, seconds = map(float, parts)
                        return minutes * 60 + seconds
                    except ValueError:
                        return None
                elif len(parts) == 1:
                    try:
                        return float(parts[0])
                    except ValueError:
                        return None
            return None

        def rank_and_delta(time_column, df, index, event_num):
            """Calculate rank and delta from a given time column and row index."""
            current_time = time_to_seconds(df[time_column][index])
            if current_time is None:
                return " ", " "

            times = [time_to_seconds(df[time_column][i]) for i in df.index
                     if str(df["EventNum"][i]) == event_num and time_to_seconds(df[time_column][i]) is not None]

            if not times:
                return " ", " "

            sorted_times = sorted(set(times))
            print(f"Sorted times for event {event_num} in column {time_column}: {sorted_times}")

            rank = sum(t < current_time for t in sorted_times) + 1
            delta = current_time - min(sorted_times) if sorted_times else None
            if delta is not None and delta == 0:
                return f"({rank})", ""

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

            return f"{split_time_str}", " "

        time_columns = ["500m", "1000m", "1500m"]

        valid_indices = {col: [] for col in time_columns}
        for col in time_columns:
            valid_indices[col] = [i for i in fl.index if not pd.isna(fl[col][i])]

        columns_with_data = {col: len(valid_indices[col]) == len(fl.index) for col in time_columns}
        columns_with_data["Finish"] = True

        for j, en in enumerate(fl["EventNum"]):
            if str(en) in selected_event_nums:
                if fl["Crew"][j] != "Empty":
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

                        m1000split, _ = split_time(fl["500m"][j], fl["1000m"][j]) if columns_with_data["1000m"] else (
                        "", "")
                        m1500split, _ = split_time(fl["1000m"][j], fl["1500m"][j]) if columns_with_data["1500m"] else (
                        "", "")
                        mFinsplit, _ = split_time(fl["1500m"][j], fl["AdjTime"][j])

                        m500rank, m500delta = rank_and_delta("500m", fl, j, str(en))
                        m1000rank, m1000delta = rank_and_delta("1000m", fl, j, str(en)) if columns_with_data[
                            "1000m"] else ("", "")
                        m1500rank, m1500delta = rank_and_delta("1500m", fl, j, str(en)) if columns_with_data[
                            "1500m"] else ("", "")

                        if m500delta == "+ 00:00.00":
                            m500delta = ""
                        if m1000delta == "+ 00:00.00":
                            m1000delta = ""
                        if m1500delta == "+ 00:00.00":
                            m1500delta = ""
                        if mFinsplit == "+ 00:00.00":
                            mFinsplit = ""

                        if m500delta == "":
                            m500delta = m500split
                            m500split = ""
                        if m1000delta == "":
                            m1000delta = m1000split
                            m1000split = ""
                        if m1500delta == "":
                            m1500delta = m1500split
                            m1500split = ""
                        if mFinsplit == "":
                            mFinsplit = fl["Delta"][j]
                            fl["Delta"][j] = ""

                        bow_value = str(fl["Bow"][j])
                        bow_display = bow_value.split()[0] if " " in bow_value else bow_value

                        if selectedFile == "RESULTS WITH QUAL":
                            res_withQual.append([
                                str(fl["Place"][j]).split(sep=".")[0],
                                str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                format_time(time_to_seconds(fl["500m"][j])) if columns_with_data["500m"] else "",
                                m500rank, m500delta,
                                format_time(time_to_seconds(fl["1000m"][j])) if columns_with_data["1000m"] else "",
                                m1000rank, m1000delta, m1000split,
                                format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                m1500rank, m1500delta, m1500split,
                                fl["AdjTime"][j],
                                f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                mFinsplit, fl["Qual"][j], coach_list.get(fl["Stroke"][j]), en
                            ])

                        if selectedFile == "RESULTS WITH QUAL (START TIME)":
                            res_withQual_start.append([
                                str(fl["Place"][j]).split(sep=".")[0],
                                bow_display,
                                f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                " ",
                                " ", " ",
                                " ",
                                " ", " ", " ",
                                fl["Start"][j],
                                " ", " ", " ",
                                fl["AdjTime"][j],
                                f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                mFinsplit, fl["Qual"][j], coach_list.get(fl["Stroke"][j]), en
                            ])
                        if selectedFile == "RESULTS WITH NO QUAL":
                            res_noQual.append([str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                                            str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                            f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                            fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["500m"][j])) if columns_with_data["500m"] else "",
                                            m500rank, m500delta,
                                            format_time(time_to_seconds(fl["1000m"][j])) if columns_with_data["1000m"] else "",
                                            m1000rank, m1000delta, m1000split,
                                            format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                            m1500rank, m1500delta, m1500split,
                                            fl["AdjTime"][j],
                                            f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                            mFinsplit, coach_list.get(fl["Stroke"][j]), en])

                        if selectedFile == "RESULTS WITH NO QUAL (START TIME)":
                            res_noQual_start.append(
                                [str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                                 bow_display,
                                 f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                 fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                 " ",
                                 " ", " ",
                                 " ",
                                 " ", " ", " ",
                                 fl["Start"][j],
                                 " ", " ", " ",
                                 fl["AdjTime"][j],
                                 f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                 mFinsplit, coach_list.get(fl["Stroke"][j]), en])

                        if selectedFile == "TEST RACE":
                            atlase.append([str(fl["Place"][j]).split(sep=".")[0], bow_display,
                                           f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                           fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["Start"][j], fl["AdjTime"][j], fl["Qual"][j], modeltime, en])

                        if selectedFile == "MASTER RESULTS WITH QUAL":
                            res_withQual_master.append([str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                               f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                               fl["Crew"][j],
                                               fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                            m1500rank, m1500delta, fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                                               f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                               fl["Qual"][j], masterFun(str(fl["PenaltyCode"][j])), en])

                        if selectedFile == "MASTER RESULTS WITH QUAL (START TIME)":
                            res_withQual_master_start.append([str(fl["Place"][j]).split(sep=".")[0],
                                                                bow_display,
                                                                f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                                                fl["Crew"][j],
                                                                fl["Stroke"][j].replace("/", "<br>"),
                                                                fl["Start"][j],
                                                                " ", " ", fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                                 f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                                                fl["Qual"][j], masterFun(str(fl["PenaltyCode"][j])), en])

                        if selectedFile == "MASTER RESULTS WITH NO QUAL":
                            res_noQual_master.append([str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                                 str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                 f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                 fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), format_time(time_to_seconds(fl["1500m"][j])) if columns_with_data["1500m"] else "",
                                            m1500rank, m1500delta, fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                                               f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "", masterFun(str(fl["PenaltyCode"][j])), en])

                        if selectedFile == "MASTER RESULTS WITH NO QUAL (START TIME)":
                            res_noQual_master_start.append([str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep=".")[0]})",
                                 bow_display,
                                 f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                 fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                 fl["Start"][j],
                                 " ", " ", fl["RawTime"][j], mFinsplit, fl["AdjTime"][j],
                                 f'+ {format_time(time_to_seconds(fl["Delta"][j]))}' if fl["Delta"][j] else "",
                                 masterFun(str(fl["PenaltyCode"][j])), en])

                        if selectedFile == "STARTLIST":
                            start_data.append([str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                               f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                               fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), coach_list.get(fl["Stroke"][j]), en])

                        if selectedFile == "STARTLIST (START TIME)":
                            start_data_wthStart.append([str(fl["Bow"][j]).split(sep="[")[0],
                                               str(fl["Bow"][j]).split(sep="[")[1],
                                               fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), coach_list.get(fl["Stroke"][j]), en])
                        if selectedFile == "MASTER STARTLIST":
                            start_data_master.append([str(fl["Bow"][j]) if "[" in str(fl["Bow"][j]) else str(fl["Bow"][j]).split(sep=".")[0],
                                               f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                               fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"),
                                               masterFun(str(fl["PenaltyCode"][j])), coach_list.get(fl["Stroke"][j]), en])

                        if selectedFile == "ENTRY LIST":
                            entry_data.append([f'<img src="flags/{flag_list.get(str(fl["CrewAbbrev"][j]), "default_flag")}" style="max-width: 6mm">',
                                                  fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), en])

        print(res_withQual)

        if combineRace and selectedFile == "RESULTS WITH QUAL":
            first_en_value = res_withQual[0][21]
            for entry in res_withQual:
                entry[21] = first_en_value

            res_withQual.sort(key=lambda x: (not x[0], x[16] if x[16] is not None else '99:99.99'))

            place_counter = 1
            for entry in res_withQual:
                if entry[0]:
                    entry[0] = str(place_counter)
                    place_counter += 1
                else:
                    entry[0] = ""

            first_time = time_to_seconds(res_withQual[0][16]) if res_withQual[0][
                16] else None

            for entry in res_withQual:
                current_time = time_to_seconds(entry[16]) if entry[16] else None
                if current_time is not None and first_time is not None:
                    delta = current_time - first_time
                    entry[17] = f'+ {format_time(delta)}' if delta > 0 else ''
                else:
                    entry[17] = ""

            for entry in res_withQual:

                entry[5] = ""
                entry[6] = ""
                entry[7] = ""
                entry[8] = ""
                entry[9] = ""
                entry[10] = ""
                entry[11] = ""
                entry[12] = ""
                entry[13] = ""
                entry[14] = ""
                entry[15] = ""
                entry[18] = ""

        if combineRace and selectedFile == "RESULTS WITH QUAL (START TIME)":
            first_en_value = res_withQual_start[0][21]
            for entry in res_withQual_start:
                entry[21] = first_en_value

            res_withQual_start.sort(key=lambda x: (not x[0], x[16] if x[16] is not None else '99:99.99'))

            place_counter = 1
            for entry in res_withQual_start:
                if entry[0]:
                    entry[0] = str(place_counter)
                    place_counter += 1
                else:
                    entry[0] = ""

            first_time = time_to_seconds(res_withQual_start[0][16]) if res_withQual_start[0][
                16] else None

            for entry in res_withQual_start:
                current_time = time_to_seconds(entry[16]) if entry[16] else None
                if current_time is not None and first_time is not None:
                    delta = current_time - first_time
                    entry[17] = f'+ {format_time(delta)}' if delta > 0 else ''
                else:
                    entry[17] = ""

            for entry in res_withQual_start:
                entry[5] = ""
                entry[6] = ""
                entry[7] = ""
                entry[8] = ""
                entry[9] = ""
                entry[10] = ""
                entry[11] = ""
                entry[13] = ""
                entry[14] = ""
                entry[15] = ""
                entry[18] = ""

        if combineRace and selectedFile == "MASTER RESULTS WITH QUAL":
            first_en_value = res_withQual_master[0][14]
            for entry in res_withQual_master:
                entry[14] = first_en_value

            res_withQual_master.sort(key=lambda x: (not x[0], x[10] if x[10] is not None else '99:99.99'))

            place_counter = 1
            for entry in res_withQual_master:
                if entry[0]:
                    entry[0] = str(place_counter)
                    place_counter += 1
                else:
                    entry[0] = ""

            first_time = time_to_seconds(res_withQual_master[0][10]) if res_withQual_master[0][10] else None

            for entry in res_withQual_master:
                current_time = time_to_seconds(entry[10]) if entry[10] else None
                if current_time is not None and first_time is not None:
                    delta = current_time - first_time
                    entry[11] = f'+ {format_time(delta)}' if delta > 0 else ''
                else:
                    entry[11] = ""

            for entry in res_withQual_master:
                entry[5] = ""
                entry[6] = ""
                entry[7] = ""
                entry[9] = ""

        if combineRace and selectedFile == "MASTER RESULTS WITH QUAL (START TIME)":
            first_en_value = res_withQual_master_start[0][14]
            for entry in res_withQual_master_start:
                entry[14] = first_en_value

            res_withQual_master_start.sort(key=lambda x: (not x[0], x[10] if x[10] is not None else '99:99.99'))

            place_counter = 1
            for entry in res_withQual_master_start:
                if entry[0]:
                    entry[0] = str(place_counter)
                    place_counter += 1
                else:
                    entry[0] = ""

            first_time = time_to_seconds(res_withQual_master_start[0][10]) if res_withQual_master_start[0][10] else None

            for entry in res_withQual_master_start:
                current_time = time_to_seconds(entry[10]) if entry[10] else None
                if current_time is not None and first_time is not None:
                    delta = current_time - first_time
                    entry[11] = f'+ {format_time(delta)}' if delta > 0 else ''
                else:
                    entry[11] = ""

            for entry in res_withQual_master_start:
                entry[6] = ""
                entry[7] = ""
                entry[9] = ""

        atlase.sort(key=lambda x: x[8] if x[8] is not None else -float('inf'), reverse=True)
        for index, row in enumerate(atlase, start=1):
            row[0] = str(index)
        import re

        def extract_number(value):
            value_str = str(value)
            match = re.search(r'\d+', value_str)
            return int(
                match.group()) if match else 999999

        start_data.sort(key=lambda x: (extract_number(x[5]), extract_number(x[0])))
        start_data_wthStart.sort(key=lambda x: (extract_number(x[5]), extract_number(x[0])))
        start_data_master.sort(key=lambda x: (extract_number(x[6]), extract_number(x[0])))

        current_date = datetime.datetime.now().strftime('%Y-%m-%d')
        current_time = datetime.datetime.now().strftime('%H:%M:%S')

        html_dir = os.path.join(os.path.dirname(__file__), 'html')

        if selectedFile == "RESULTS WITH QUAL":
            with open("html/body_r_withQual.html", "r") as f:
                html_results_withQual = f.read()

            html_results_withQual = html_results_withQual.format(compName=self.comp_name.text,
                                                                 compDates=self.comp_date.text, cDate=current_date,
                                                                 cTime=current_time)

            with open(os.path.join(html_dir, "body_r_usualWithQual.txt"), "r") as f:
                bodyInfo = f.read()

            if combineRace:
                with open(os.path.join(html_dir, "header_r_usualWithQual_cmb.txt"), "r") as f:
                    headerInfo = f.read()
            else:
                with open(os.path.join(html_dir, "header_r_usualWithQual.txt"), "r") as f:
                    headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True
            for j, a in enumerate(res_withQual):
                if a[-1] == last:
                    html_results_withQual = html_results_withQual.replace("[rinda]",
                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_results_withQual = html_results_withQual.replace("[rinda]", "")
                        last_prog = info[last_id-1][8]
                        html_results_withQual = html_results_withQual.replace("[prog_sys]", last_prog)
                        html_results_withQual = html_results_withQual.replace("[prog_sys]", "")
                        with open(os.path.join(html_dir, f"log_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_results_withQual)

                        with open(os.path.join(html_dir, "body_r_withQual.html"), "r") as f:
                            html_results_withQual = f.read()

                        html_results_withQual = html_results_withQual.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                           cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_usualWithQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_usualWithQual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_results_withQual = html_results_withQual.replace("[rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                 a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                        html_results_withQual = html_results_withQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                      info[last_id][4], info[last_id][3], info[last_id][0],
                                                                      info[last_id][5], info[last_id][6]))
                    else:
                        html_results_withQual = html_results_withQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                                      info[last_id][4], info[last_id][3], info[last_id][0],
                                                                      info[last_id][5], info[last_id][6]))
                        html_results_withQual = html_results_withQual.replace("[rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                 a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_results_withQual = html_results_withQual.replace("[rinda]", "")
            last_prog = info[last_id-1][8]
            html_results_withQual = html_results_withQual.replace("[prog_sys]", last_prog)
            html_results_withQual = html_results_withQual.replace("[prog_sys]", "")
            with open(os.path.join(html_dir, f"log_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_results_withQual)

        if selectedFile == "RESULTS WITH QUAL (START TIME)":
            with open("html/body_r_withQual.html", "r") as f:
                html_results_withQual_wthStart = f.read()

            html_results_withQual_wthStart = html_results_withQual_wthStart.format(compName=self.comp_name.text,
                                                                                   compDates=self.comp_date.text,
                                                                                   cDate=current_date,
                                                                                   cTime=current_time)

            with open(os.path.join(html_dir, "body_r_usualWithQual.txt"), "r") as f:
                bodyInfo = f.read()

            if combineRace:
                with open(os.path.join(html_dir, "header_r_usualWithQual_wthStart_cmb.txt"), "r") as f:
                    headerInfo = f.read()
            else:
                with open(os.path.join(html_dir, "header_r_usualWithQual_wthStart.txt"), "r") as f:
                    headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True
            for j, a in enumerate(res_withQual_start):
                if a[-1] == last:
                    html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[rinda]",
                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10],
                                                        a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19],
                                                        a[20]) + "\n[rinda]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[rinda]", "")
                        last_prog = info[last_id - 1][8]
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[prog_sys]", last_prog)
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[prog_sys]", "")
                        with open(os.path.join(html_dir, f"log_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_results_withQual_wthStart)

                        with open(os.path.join(html_dir, "body_r_withQual.html"), "r") as f:
                            html_results_withQual_wthStart = f.read()

                        html_results_withQual_wthStart = html_results_withQual_wthStart.format(compName=self.comp_name.text, compDates=self.comp_date.text, cDate=current_date,
                                           cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_usualWithQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_usualWithQual_wthStart.txt"), "r") as f:
                            headerInfo = f.read()

                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                       a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13],
                                                                       a[14], a[15], a[16], a[17], a[18], a[19],
                                                                       a[20]) + "\n[rinda]")
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[header]",
                                            headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                              info[last_id][4], info[last_id][3], info[last_id][0],
                                                              info[last_id][5], info[last_id][6]))
                    else:
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[header]",
                                            headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                              info[last_id][4], info[last_id][3], info[last_id][0],
                                                              info[last_id][5], info[last_id][6]))
                        html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                       a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13],
                                                                       a[14], a[15], a[16], a[17], a[18], a[19],
                                                                       a[20]) + "\n[rinda]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[rinda]", "")
            last_prog = info[last_id - 1][8]
            html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[prog_sys]", last_prog)
            html_results_withQual_wthStart = html_results_withQual_wthStart.replace("[prog_sys]", "")
            with open(os.path.join(html_dir, f"log_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_results_withQual_wthStart)

        if selectedFile == "MASTER RESULTS WITH QUAL":
            with open("html/body_r_withQual.html", "r") as f:
                html_master_withQual = f.read()

            html_master_withQual = html_master_withQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                               cDate=current_date,
                                                               cTime=current_time)

            with open(os.path.join(html_dir, "body_r_masterWithQual.txt"), "r") as f:
                bodyInfo = f.read()
            if combineRace:
                with open(os.path.join(html_dir, "header_r_masterWithQual_cmb.txt"), "r") as f:
                    headerInfo = f.read()
            else:
                with open(os.path.join(html_dir, "header_r_masterWithQual.txt"), "r") as f:
                    headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True
            for j, a in enumerate(res_withQual_master):
                if a[-1] == last:
                    html_master_withQual = html_master_withQual.replace("[rinda]",
                                                    bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8],
                                                                  a[9], a[10], a[11], a[12], a[13], a[14]) + "\n[rinda]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_master_withQual = html_master_withQual.replace("[rinda]", "")
                        last_prog = info[last_id-1][8]
                        html_master_withQual = html_master_withQual.replace("[prog_sys]", last_prog)
                        html_master_withQual = html_master_withQual.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_mast_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_master_withQual)

                        with open(os.path.join(html_dir, "body_r_withQual.html"), "r") as f:
                            html_master_withQual = f.read()

                        html_master_withQual = html_master_withQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                       cDate=current_date,
                                                       cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_masterWithQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_masterWithQual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_master_withQual = html_master_withQual.replace("[rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                                 a[6], a[7], a[8], a[9], a[10],
                                                                                 a[11], a[12], a[13], a[14]) + "\n[rinda]")
                        html_master_withQual = html_master_withQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                      info[last_id][1],
                                                                                      info[last_id][4], info[last_id][3],
                                                                                      info[last_id][0],
                                                                                      info[last_id][5], info[last_id][6]))
                    else:
                        html_master_withQual = html_master_withQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                      info[last_id][1],
                                                                                      info[last_id][4], info[last_id][3],
                                                                                      info[last_id][0],
                                                                                      info[last_id][5], info[last_id][6]))
                        html_master_withQual = html_master_withQual.replace("[rinda]", bodyInfo.format(a[0],
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
            html_master_withQual = html_master_withQual.replace("[rinda]", "")
            last_prog = info[last_id-1][8]
            html_master_withQual = html_master_withQual.replace("[prog_sys]", last_prog)
            html_master_withQual = html_master_withQual.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_mast_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_master_withQual)

        if selectedFile == "MASTER RESULTS WITH QUAL (START TIME)":

            with open("html/body_r_withQual.html", "r") as f:
                html_master_withQual_wthStart = f.read()

            html_master_withQual_wthStart = html_master_withQual_wthStart.format(compName=self.comp_name.text,
                                                                                 compDates=self.comp_date.text,
                                                                                 cDate=current_date,
                                                                                 cTime=current_time)

            with open(os.path.join(html_dir, "body_r_masterWithQual.txt"), "r") as f:
                bodyInfo = f.read()

            if combineRace:
                with open(os.path.join(html_dir, "header_r_masterWithQual_wthStart_cmb.txt"), "r") as f:
                    headerInfo = f.read()
            else:
                with open(os.path.join(html_dir, "header_r_masterWithQual_wthStart.txt"), "r") as f:
                    headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True
            for j, a in enumerate(res_withQual_master_start):
                if a[-1] == last:
                    html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[rinda]",
                                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                        a[5], a[6], a[7], a[8],
                                                                                        a[9], a[10], a[11], a[12],
                                                                                        a[13], a[14]) + "\n[rinda]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[rinda]", "")
                        last_prog = info[last_id - 1][8]
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[prog_sys]", last_prog)
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_mast_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_master_withQual_wthStart)

                        with open(os.path.join(html_dir, "body_r_withQual.html"), "r") as f:
                            html_master_withQual_wthStart = f.read()

                        html_master_withQual_wthStart = html_master_withQual_wthStart.format(compName=self.comp_name.text,
                                                                           compDates=self.comp_date.text,
                                                                           cDate=current_date,
                                                                           cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_masterWithQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_masterWithQual_wthStart.txt"), "r") as f:
                            headerInfo = f.read()

                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[rinda]",
                                                                            bodyInfo.format(a[0], a[1], a[2], a[3],
                                                                                            a[4], a[5],
                                                                                            a[6], a[7], a[8], a[9],
                                                                                            a[10],
                                                                                            a[11], a[12], a[13],
                                                                                            a[14]) + "\n[rinda]")
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[header]",
                                                                            headerInfo.format(info[last_id][7],
                                                                                              info[last_id][2],
                                                                                              info[last_id][1],
                                                                                              info[last_id][4],
                                                                                              info[last_id][3],
                                                                                              info[last_id][0],
                                                                                              info[last_id][5],
                                                                                              info[last_id][6]))
                    else:
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[header]",
                                                                            headerInfo.format(info[last_id][7],
                                                                                              info[last_id][2],
                                                                                              info[last_id][1],
                                                                                              info[last_id][4],
                                                                                              info[last_id][3],
                                                                                              info[last_id][0],
                                                                                              info[last_id][5],
                                                                                              info[last_id][6]))
                        html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[rinda]", bodyInfo.format(a[0],
                                                                                                       a[1],
                                                                                                       a[2],
                                                                                                       a[3],
                                                                                                       a[4],
                                                                                                       a[5],
                                                                                                       a[6],
                                                                                                       a[7],
                                                                                                       a[8], a[9],
                                                                                                       a[10], a[11],
                                                                                                       a[12], a[13], a[
                                                                                                           14]) + "\n[rinda]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1
            html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[rinda]", "")
            last_prog = info[last_id - 1][8]
            html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[prog_sys]", last_prog)
            html_master_withQual_wthStart = html_master_withQual_wthStart.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_mast_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_master_withQual_wthStart)

        if selectedFile == "MASTER STARTLIST":

            with open("html/main_s.html", "r") as f:
                html_startlists_master = f.read()

            html_startlists_master = html_startlists_master.format(compName=self.comp_name.text,
                                                                   compDates=self.comp_date.text, cDate=current_date,
                                                                   cTime=current_time)

            with open(os.path.join(html_dir, "body_s_master.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_s_master.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert_start = True

            for j, a in enumerate(start_data_master):
                if a[-1] == last:
                    html_startlists_master = html_startlists_master.replace("[st_rinda]",
                                                                bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5]) + "\n[st_rinda]")
                    last = a[-1]

                else:
                    if not first_insert_start:
                        html_startlists_master = html_startlists_master.replace("[st_rinda]", "")
                        last_prog = info[last_id - 1][8]
                        html_startlists_master = html_startlists_master.replace("[prog_sys]", last_prog)
                        html_startlists_master = html_startlists_master.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"start_log_master_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_startlists_master)

                        with open(os.path.join(html_dir, "main_s.html"), "r") as f:
                            html_startlists_master = f.read()

                        html_startlists_master = html_startlists_master.format(compName=self.comp_name.text,
                                                                   compDates=self.comp_date.text,
                                                                   cDate=current_date, cTime=current_time)

                        with open(os.path.join(html_dir, "body_s_master.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_s_master.txt"), "r") as f:
                            headerInfo = f.read()

                        html_startlists_master = html_startlists_master.replace("[st_rinda]",
                                                                    bodyInfo.format(a[0], a[1], a[2], a[3],
                                                                              a[4], a[5]) + "\n[st_rinda]")
                        html_startlists_master = html_startlists_master.replace("[header]",
                                                                    headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                    else:
                        html_startlists_master = html_startlists_master.replace("[st_rinda]", bodyInfo.format(a[0], a[1], a[2], a[3],
                                                                                            a[4], a[5]) + "\n[st_rinda]")
                        html_startlists_master = html_startlists_master.replace("[header]",
                                                                    headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                        first_insert_start = False

                    last = a[-1]
                    last_id += 1

            html_startlists_master = html_startlists_master.replace("[st_rinda]", "")
            last_prog = info[last_id - 1][8]
            html_startlists_master = html_startlists_master.replace("[prog_sys]", last_prog)
            html_startlists_master = html_startlists_master.replace("[prog_sys]", "")
            with open(os.path.join(html_dir, f"start_log_master_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_startlists_master)

        if selectedFile == "STARTLIST":

            with open("html/main_s.html", "r") as f:
                html_startlists = f.read()

            html_startlists = html_startlists.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                     cDate=current_date,
                                                     cTime=current_time)

            with open(os.path.join(html_dir, "body_s.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_s_usual.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert_start = True

            for j, a in enumerate(start_data):
                if a[-1] == last:
                    html_startlists = html_startlists.replace("[st_rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                    last = a[-1]
                else:
                    if not first_insert_start:
                        html_startlists = html_startlists.replace("[st_rinda]", "")
                        last_prog = info[last_id - 1][8]
                        html_startlists = html_startlists.replace("[prog_sys]", last_prog)
                        html_startlists = html_startlists.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"start_log_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_startlists)

                        with open(os.path.join(html_dir, "main_s.html"), "r") as f:
                            html_startlists = f.read()

                        html_startlists = html_startlists.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                       cDate=current_date,
                                                       cTime=current_time)

                        with open(os.path.join(html_dir, "body_s.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_s_usual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_startlists = html_startlists.replace("[st_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                        html_startlists = html_startlists.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                    else:
                        html_startlists = html_startlists.replace("[st_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                        html_startlists = html_startlists.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                        first_insert_start = False

                    last = a[-1]
                    last_id += 1

            html_startlists = html_startlists.replace("[st_rinda]", "")
            last_prog = info[last_id - 1][8]
            html_startlists = html_startlists.replace("[prog_sys]", last_prog)
            html_startlists = html_startlists.replace("[prog_sys]", "")
            with open(os.path.join(html_dir, f"start_log_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_startlists)

        if selectedFile == "STARTLIST (START TIME)":

            with open("html/main_s.html", "r") as f:
                html_wthStart_startlists = f.read()

            html_wthStart_startlists = html_wthStart_startlists.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                     cDate=current_date,
                                                     cTime=current_time)

            with open(os.path.join(html_dir, "body_s_wthTime.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_s_usual_wthTime.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert_start = True

            for j, a in enumerate(start_data_wthStart):
                if a[-1] == last:
                    html_wthStart_startlists = html_wthStart_startlists.replace("[st_rinda]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                    last = a[-1]
                else:
                    if not first_insert_start:
                        html_wthStart_startlists = html_wthStart_startlists.replace("[st_rinda]", "")
                        last_prog = info[last_id - 1][8]
                        html_wthStart_startlists = html_wthStart_startlists.replace("[prog_sys]", last_prog)
                        html_wthStart_startlists = html_wthStart_startlists.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"start_log_wthTime_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_wthStart_startlists)

                        with open(os.path.join(html_dir, "main_s.html"), "r") as f:
                            html_wthStart_startlists = f.read()

                        html_wthStart_startlists = html_wthStart_startlists.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                       cDate=current_date,
                                                       cTime=current_time)

                        with open(os.path.join(html_dir, "body_s_wthStart.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_s_usual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_wthStart_startlists = html_wthStart_startlists.replace("[st_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                        html_wthStart_startlists = html_wthStart_startlists.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                    else:
                        html_wthStart_startlists = html_wthStart_startlists.replace("[st_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4]) + "\n[st_rinda]")
                        html_wthStart_startlists = html_wthStart_startlists.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                  info[last_id][1],
                                                                                  info[last_id][4], info[last_id][3],
                                                                                  info[last_id][0],
                                                                                  info[last_id][5], info[last_id][6]))
                        first_insert_start = False

                    last = a[-1]
                    last_id += 1

            html_wthStart_startlists = html_wthStart_startlists.replace("[st_rinda]", "")
            last_prog = info[last_id - 1][8]
            html_wthStart_startlists = html_wthStart_startlists.replace("[prog_sys]", last_prog)
            html_wthStart_startlists = html_wthStart_startlists.replace("[prog_sys]", "")
            with open(os.path.join(html_dir, f"start_log_wthTime_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_wthStart_startlists)

        if selectedFile == "ENTRY LIST":

            with open("html/main_entries.html", "r") as f:
                html_entries = f.read()

            html_entries = html_entries.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                               cDate=current_date,
                                               cTime=current_time)

            with open(os.path.join(html_dir, "body_entries.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_entries.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert_start = True

            for j, a in enumerate(entry_data):
                if a[-1] == last:
                    html_entries = html_entries.replace("[entry_rinda]",
                                                    bodyInfo.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                    last = a[-1]
                else:
                    if not first_insert_start:
                        html_entries = html_entries.replace("[entry_rinda]", "")

                        with open(os.path.join(html_dir, f"entry_by_events_log_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_entries)

                        with open(os.path.join(html_dir, "main_entries.html"), "r") as f:
                            html_entries = f.read()

                        html_entries = html_entries.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                       cDate=current_date,
                                                       cTime=current_time)

                        with open(os.path.join(html_dir, "body_entries.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_entries.txt"), "r") as f:
                            headerInfo = f.read()

                        html_entries = html_entries.replace("[entry_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                        html_entries = html_entries.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                        info[last_id][1]))
                    else:
                        html_entries = html_entries.replace("[entry_rinda]",
                                                        bodyInfo.format(a[0], a[1], a[2]) + "\n[entry_rinda]")
                        html_entries = html_entries.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                        info[last_id][1]))
                        first_insert_start = False

                    last = a[-1]
                    last_id += 1

            html_entries = html_entries.replace("[entry_rinda]", "")
            with open(os.path.join(html_dir, f"entry_by_events_log_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_entries)

        if selectedFile == "RESULTS WITH NO QUAL":

            with open("html/body_r_noQual.html", "r") as f:
                html_results_noQual = f.read()

            html_results_noQual = html_results_noQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                             cDate=current_date,
                                                             cTime=current_time)

            with open(os.path.join(html_dir, "body_r_usualNoQual.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_r_usualNoQual.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(res_noQual):
                if a[-1] == last:
                    html_results_noQual = html_results_noQual.replace("[rinda_noq]",
                                          bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                 a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_results_noQual = html_results_noQual.replace("[rinda_noq]", "")
                        html_results_noQual = html_results_noQual.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_noq_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_results_noQual)

                        with open(os.path.join(html_dir, "body_r_noQual.html"), "r") as f:
                            html_results_noQual = f.read()

                        html_results_noQual = html_results_noQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                             cDate=current_date,
                                             cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_usualNoQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_usualNoQual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_results_noQual = html_results_noQual.replace("[rinda_noq]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                 a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                        html_results_noQual = html_results_noQual.replace("[header]",
                                              headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                             info[last_id][4], info[last_id][3], info[last_id][0],
                                                             info[last_id][5], info[last_id][6]))
                    else:
                        html_results_noQual = html_results_noQual.replace("[header]",
                                              headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                             info[last_id][4], info[last_id][3], info[last_id][0],
                                                             info[last_id][5], info[last_id][6]))
                        html_results_noQual = html_results_noQual.replace("[rinda_noq]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                 a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20]) + "\n[rinda_noq]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_results_noQual = html_results_noQual.replace("[rinda_noq]", "")
            html_results_noQual = html_results_noQual.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_noq_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_results_noQual)

        if selectedFile == "RESULTS WITH NO QUAL (START TIME)":

            with open("html/body_r_noQual.html", "r") as f:
                html_results_noQual_wthStart = f.read()

            html_results_noQual_wthStart = html_results_noQual_wthStart.format(compName=self.comp_name.text,
                                                                               compDates=self.comp_date.text,
                                                                               cDate=current_date,
                                                                               cTime=current_time)

            with open(os.path.join(html_dir, "body_r_usualNoQual.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_r_usualNoQual_wthStart.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(res_noQual_start):
                if a[-1] == last:
                    html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[rinda_noq]",
                                                                      bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                      a[5],
                                                                                      a[6], a[7], a[8], a[9], a[10],
                                                                                      a[11], a[12], a[13], a[14], a[15],
                                                                                      a[16], a[17], a[18], a[19],
                                                                                      a[20]) + "\n[rinda_noq]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[rinda_noq]", "")
                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_noq_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_results_noQual_wthStart)

                        with open(os.path.join(html_dir, "body_r_noQual.html"), "r") as f:
                            html_results_noQual_wthStart = f.read()

                        html_results_noQual_wthStart = html_results_noQual_wthStart.format(compName=self.comp_name.text,
                                                                         compDates=self.comp_date.text,
                                                                         cDate=current_date,
                                                                         cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_usualNoQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_usualNoQual_wthStart.txt"), "r") as f:
                            headerInfo = f.read()

                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[rinda_noq]",
                                                                          bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                          a[5],
                                                                                          a[6], a[7], a[8], a[9], a[10],
                                                                                          a[11], a[12], a[13], a[14],
                                                                                          a[15], a[16], a[17], a[18],
                                                                                          a[19],
                                                                                          a[20]) + "\n[rinda_noq]")
                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[header]",
                                                                          headerInfo.format(info[last_id][7],
                                                                                            info[last_id][2],
                                                                                            info[last_id][1],
                                                                                            info[last_id][4],
                                                                                            info[last_id][3],
                                                                                            info[last_id][0],
                                                                                            info[last_id][5],
                                                                                            info[last_id][6]))
                    else:
                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[header]",
                                                                          headerInfo.format(info[last_id][7],
                                                                                            info[last_id][2],
                                                                                            info[last_id][1],
                                                                                            info[last_id][4],
                                                                                            info[last_id][3],
                                                                                            info[last_id][0],
                                                                                            info[last_id][5],
                                                                                            info[last_id][6]))
                        html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[rinda_noq]",
                                                                          bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                          a[5],
                                                                                          a[6], a[7], a[8], a[9], a[10],
                                                                                          a[11], a[12], a[13], a[14],
                                                                                          a[15], a[16], a[17], a[18],
                                                                                          a[19],
                                                                                          a[20]) + "\n[rinda_noq]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[rinda_noq]", "")
            html_results_noQual_wthStart = html_results_noQual_wthStart.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_noq_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_results_noQual_wthStart)

        if selectedFile == "TEST RACE":

            with open("html/main_atlase.html", "r") as f:
                html_atlase = f.read()

            html_atlase = html_atlase.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                             cDate=current_date, cTime=current_time)

            with open(os.path.join(html_dir, "body_atlase.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_atlase.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(atlase):
                if a[-1] == last:
                    html_atlase = html_atlase.replace("[rinda_noq]",
                                          bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]) + "\n[rinda_noq]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_atlase = html_atlase.replace("[rinda_noq]", "")
                        html_atlase = html_atlase.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"atlase_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_atlase)

                        with open(os.path.join(html_dir, "body_r_noQual.html"), "r") as f:
                            html_atlase = f.read()

                        html_atlase = html_atlase.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                             cDate=current_date,
                                             cTime=current_time)

                        with open(os.path.join(html_dir, "body_atlase.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_atlase.txt"), "r") as f:
                            headerInfo = f.read()

                        html_atlase = html_atlase.replace("[rinda_noq]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                        a[6], a[7], a[8]) + "\n[rinda_noq]")
                        html_atlase = html_atlase.replace("[header]",
                                              headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                             info[last_id][4], info[last_id][3], info[last_id][0],
                                                             info[last_id][5], info[last_id][6]))
                    else:
                        html_atlase = html_atlase.replace("[header]",
                                              headerInfo.format(info[last_id][7], info[last_id][2], info[last_id][1],
                                                             info[last_id][4], info[last_id][3], info[last_id][0],
                                                             info[last_id][5], info[last_id][6]))
                        html_atlase = html_atlase.replace("[rinda_noq]", bodyInfo.format(a[0], a[1], a[2], a[3],
                                                                        a[4], a[5], a[6], a[7], a[8]) + "\n[rinda_noq]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_atlase = html_atlase.replace("[rinda_noq]", "")
            html_atlase = html_atlase.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"atlase_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_atlase)

        if selectedFile == "MASTER RESULTS WITH NO QUAL":

            with open("html/body_r_noQual.html", "r") as f:
                html_master_noQual = f.read()

            html_master_noQual = html_master_noQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                           cDate=current_date,
                                                           cTime=current_time)

            with open(os.path.join(html_dir, "body_r_masterNoQual.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_r_masterNoQual.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(res_noQual_master):
                if a[-1] == last:
                    html_master_noQual = html_master_noQual.replace("[rinda_noq]",
                                                      bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8],
                                                                       a[9], a[10], a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_master_noQual = html_master_noQual.replace("[rinda_noq]", "")
                        html_master_noQual = html_master_noQual.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_noq_master_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_master_noQual)

                        with open(os.path.join(html_dir, "body_r_noQual.html"), "r") as f:
                            html_master_noQual = f.read()

                        html_master_noQual = html_master_noQual.format(compName=self.comp_name.text, compDates=self.comp_date.text,
                                                         cDate=current_date,
                                                         cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_masterNoQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_masterNoQual.txt"), "r") as f:
                            headerInfo = f.read()

                        html_master_noQual = html_master_noQual.replace("[rinda_noq]",
                                                          bodyInfo.format(a[0], a[1], a[2], a[3], a[4], a[5],
                                                                           a[6], a[7], a[8], a[9], a[10],
                                                                           a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                        html_master_noQual = html_master_noQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                         info[last_id][1],
                                                                                         info[last_id][4], info[last_id][3],
                                                                                         info[last_id][0],
                                                                                         info[last_id][5],
                                                                                         info[last_id][6]))
                    else:
                        html_master_noQual = html_master_noQual.replace("[header]", headerInfo.format(info[last_id][7], info[last_id][2],
                                                                                         info[last_id][1],
                                                                                         info[last_id][4], info[last_id][3],
                                                                                         info[last_id][0],
                                                                                         info[last_id][5],
                                                                                         info[last_id][6]))
                        html_master_noQual = html_master_noQual.replace("[rinda_noq]", bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                          a[5], a[6], a[7], a[8], a[9],
                                                                                          a[10], a[11], a[12], a[13], a[14]) + "\n[rinda_noq]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_master_noQual = html_master_noQual.replace("[rinda_noq]", "")
            html_master_noQual = html_master_noQual.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_noq_master_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_master_noQual)

        if selectedFile == "MASTER RESULTS WITH NO QUAL (START TIME)":

            with open("html/body_r_noQual.html", "r") as f:
                html_master_noQual_wthStart = f.read()

            html_master_noQual_wthStart = html_master_noQual_wthStart.format(compName=self.comp_name.text,
                                                                             compDates=self.comp_date.text,
                                                                             cDate=current_date,
                                                                             cTime=current_time)

            with open(os.path.join(html_dir, "body_r_masterNoQual.txt"), "r") as f:
                bodyInfo = f.read()

            with open(os.path.join(html_dir, "header_r_masterNoQual_wthStart.txt"), "r") as f:
                headerInfo = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(res_noQual_master_start):
                if a[-1] == last:
                    html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[rinda_noq]",
                                                                    bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                      a[5], a[6], a[7], a[8],
                                                                                      a[9], a[10], a[11], a[12], a[13],
                                                                                      a[14]) + "\n[rinda_noq]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[rinda_noq]", "")
                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[prog_sys]", "")

                        with open(os.path.join(html_dir, f"log_noq_master_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                            ft.write(html_master_noQual_wthStart)

                        with open(os.path.join(html_dir, "body_r_noQual.html"), "r") as f:
                            html_master_noQual_wthStart = f.read()

                        html_master_noQual_wthStart = html_master_noQual_wthStart.format(compName=self.comp_name.text,
                                                                       compDates=self.comp_date.text,
                                                                       cDate=current_date,
                                                                       cTime=current_time)

                        with open(os.path.join(html_dir, "body_r_masterNoQual.txt"), "r") as f:
                            bodyInfo = f.read()

                        with open(os.path.join(html_dir, "header_r_masterNoQual_wthStart.txt"), "r") as f:
                            headerInfo = f.read()

                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[rinda_noq]",
                                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                          a[5],
                                                                                          a[6], a[7], a[8], a[9], a[10],
                                                                                          a[11], a[12], a[13],
                                                                                          a[14]) + "\n[rinda_noq]")
                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[header]", headerInfo.format(info[last_id][7],
                                                                                                      info[last_id][2],
                                                                                                      info[last_id][1],
                                                                                                      info[last_id][4],
                                                                                                      info[last_id][3],
                                                                                                      info[last_id][0],
                                                                                                      info[last_id][5],
                                                                                                      info[last_id][6]))
                    else:
                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[header]", headerInfo.format(info[last_id][7],
                                                                                                      info[last_id][2],
                                                                                                      info[last_id][1],
                                                                                                      info[last_id][4],
                                                                                                      info[last_id][3],
                                                                                                      info[last_id][0],
                                                                                                      info[last_id][5],
                                                                                                      info[last_id][6]))
                        html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[rinda_noq]",
                                                                        bodyInfo.format(a[0], a[1], a[2], a[3], a[4],
                                                                                          a[5], a[6], a[7], a[8], a[9],
                                                                                          a[10], a[11], a[12], a[13],
                                                                                          a[14]) + "\n[rinda_noq]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[rinda_noq]", "")
            html_master_noQual_wthStart = html_master_noQual_wthStart.replace("[prog_sys]", "")

            with open(os.path.join(html_dir, f"log_noq_master_wthStart_{last}.html"), "w", encoding='utf-8') as ft:
                ft.write(html_master_noQual_wthStart)

        if selectedFile == "SHORT STARTLIST":

            with open("html/main_s_short.html", "r") as f:
                html_short_startlists = f.read()

            html_short_startlists = html_short_startlists.format(compName=self.comp_name.text,
                                                                 compDates=self.comp_date.text, cDate=current_date,
                                                                 cTime=current_time)

            with open(os.path.join(html_dir, "body_s_short.txt"), "r") as f:
                stListSh = f.read()

            last = ''
            last_id = 0
            first_insert = True

            for j, a in enumerate(infoShort):
                if a[-1] == last:
                    html_short_startlists = html_short_startlists.replace("[short_rinda]",
                                                      stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                       a[8],
                                                                       a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                    last = a[-1]
                else:
                    if not first_insert:
                        html_short_startlists = html_short_startlists.replace("[short_rinda]",
                                                          stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                       a[8],
                                                                       a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                    else:
                        html_short_startlists = html_short_startlists.replace("[short_rinda]", stListSh.format(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7],
                                                                       a[8],
                                                                       a[9], a[10], a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24]) + "\n[short_rinda]")
                        first_insert = False

                    last = a[-1]
                    last_id += 1

            html_short_startlists = html_short_startlists.replace("[short_rinda]", "")
            with open(os.path.join(html_dir, f"html_short_startlists_log.html"), "w", encoding='utf-8') as ft:
                ft.write(html_short_startlists)

        end = datetime.datetime.now()
        print("End: ", end)
        print("Total: ", (end - start))

class MyApp(App):
    def build(self):
        return RaceApp()

if __name__ == '__main__':
    MyApp().run()