import kivy

kivy.require("1.9.1")
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout


class ButtonApp(App):
    def build(self):
        btn = Button(text="Competitions info",
                     font_size="16sp",
                     background_color=(1, 1, 1, 1),
                     color=(1, 1, 1, 1),
                     size=(32, 32),
                     size_hint=(.2, .1),
                     pos=(10, 150))
        btn.bind(on_press=self.comp_info)
        return btn

    def comp_info(self, instance):
        layout = BoxLayout(orientation='vertical', padding=10)

        self.comp_input = TextInput(hint_text="Competition name", multiline=False, size_hint=(1, 0.25))
        layout.add_widget(self.comp_input)

        self.date_input = TextInput(hint_text="Date", multiline=False, size_hint=(1, 0.25))
        layout.add_widget(self.date_input)

        confirm_button = Button(text="Confirm", size_hint=(1, 0.3))
        confirm_button.bind(on_release=self.confirm_data)
        layout.add_widget(confirm_button)

        popup = Popup(title='Competitions data input',
                      content=layout,
                      size_hint=(0.5, 0.5),
                      auto_dismiss=False)
        confirm_button.bind(on_press=popup.dismiss)
        popup.open()

    def confirm_data(self, instance):
        cn = self.comp_input.text
        cd = self.date_input.text
        print(f"Competition Name: {cn}")
        print(f"Date: {cd}")


root = ButtonApp()
root.run()
