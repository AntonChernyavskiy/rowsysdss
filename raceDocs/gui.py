import os
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from kivy.uix.checkbox import CheckBox
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window
from winreg import *

class RaceApp(BoxLayout):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.shift_down = False
        self.last_active_checkbox = None

        # Left side layout
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

        Window.bind(on_key_down=self._on_key_down)
        Window.bind(on_key_up=self._on_key_up)

    def _on_key_down(self, window, key, scancode, codepoint, modifier):
        if 'shift' in modifier:
            self.shift_down = True

    def _on_key_up(self, window, key, scancode):
        if key == 304:  # 304 is the key code for the Shift key
            self.shift_down = False

    def choose_file(self, instance):
        content = FileChooserIconView()
        content.bind(on_submit=self._file_selected)
        popup = Popup(title='Choose file', content=content, size_hint=(0.9, 0.9))
        popup.open()

    def _file_selected(self, filechooser, selection, touch):
        try:
            self.file_path = selection[0]
            self.df = pd.read_csv(self.file_path, skip_blank_lines=True, na_filter=True)
            self.df = self.df[self.df["Event"].notna()]
            self.df = self.df.reset_index()
            self.file_chooser_btn.text = f'File: {os.path.basename(self.file_path)}'
            self._populate_events()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error loading file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()

    def _populate_events(self):
        self.event_layout.clear_widgets()
        self.event_checkboxes.clear()
        if self.df is not None:
            for idx, row in self.df.iterrows():
                event_num = row.get('EventNum', '')
                event = row.get('Event', '')
                display_text = f"{event_num} - {event}" if event_num and event else str(event or event_num)
                box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30)
                checkbox = CheckBox()
                checkbox.bind(active=self._on_checkbox_active)
                box.add_widget(checkbox)
                label = Label(text=display_text, size_hint_x=0.8, halign='center', valign='middle')
                label.bind(size=label.setter('text_size'))
                box.add_widget(label)
                self.event_layout.add_widget(box)
                self.event_checkboxes.append((checkbox, idx))

    def _on_checkbox_active(self, checkbox, value):
        if self.shift_down and self.last_active_checkbox:
            current_index = next(idx for idx, (cb, _) in enumerate(self.event_checkboxes) if cb == checkbox)
            last_index = next(idx for idx, (cb, _) in enumerate(self.event_checkboxes) if cb == self.last_active_checkbox)
            start, end = sorted((current_index, last_index))
            for idx in range(start, end + 1):
                cb, _ = self.event_checkboxes[idx]
                cb.active = value
        else:
            if value:
                self.selected_events.append(checkbox)
            else:
                self.selected_events.remove(checkbox)
            self.last_active_checkbox = checkbox if value else None

    def show_export_popup(self, instance):
        content = BoxLayout(orientation='vertical')
        file_name_input = TextInput(hint_text='Enter file name', multiline=False)
        content.add_widget(file_name_input)
        export_btn = Button(text='Export')
        export_btn.bind(on_press=lambda x: self.print_files(file_name_input.text))
        content.add_widget(export_btn)
        popup = Popup(title='Export File', content=content, size_hint=(0.6, 0.4))
        popup.open()

    def print_files(self, file_name):
        try:
            aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
            aKey = OpenKey(aReg, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', 0, KEY_READ)
            chromePath = QueryValue(aKey, None)
            CloseKey(aKey)
            CloseKey(aReg)

            selected_events = [event for checkbox, event in self.event_checkboxes if checkbox.active]
            html = "<html><body><h1>Selected Events</h1><table border='1'>"
            for event in selected_events:
                html += f"<tr><td>{event}</td></tr>"
            html += "</table></body></html>"

            with open(f"{file_name}.html", "w") as f:
                f.write(html)

            absolute_path = os.path.dirname(__file__)
            os.system(f"\"{chromePath}\" --headless --disable-gpu --print-to-pdf-no-header --print-to-pdf={absolute_path}/{file_name}.pdf file:///{absolute_path}/{file_name}.html")

            popup = Popup(title='Print Complete', content=Label(text=f'Printed to {file_name}.pdf'), size_hint=(0.6, 0.4))
            popup.open()
        except Exception as e:
            popup = Popup(title='Error', content=Label(text=f'Error printing file: {str(e)}'), size_hint=(0.6, 0.4))
            popup.open()

class RaceAppMain(App):
    def build(self):
        return RaceApp()

if __name__ == '__main__':
    RaceAppMain().run()
