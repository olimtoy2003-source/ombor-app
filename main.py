import os
import sys
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.clock import Clock

class OmborApp(App):
    def build(self):
        self.title = "Ombor - ZIP Platalar Baza"
        
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Qidiruv paneli
        search_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=50, spacing=10)
        
        self.search_input = TextInput(
            hint_text="Plata nomi yoki joylashuvini kiriting...",
            multiline=False,
            font_size='18sp',
            size_hint_x=0.75
        )
        self.search_input.bind(on_text_validate=self.search_data)
        
        search_button = Button(
            text="Qidirish",
            size_hint_x=0.25,
            background_color=(0.2, 0.6, 1, 1),
            font_size='16sp'
        )
        search_button.bind(on_press=self.search_data)
        
        search_layout.add_widget(self.search_input)
        search_layout.add_widget(search_button)
        main_layout.add_widget(search_layout)
        
        # ScrollView
        self.scroll_view = ScrollView(size_hint=(1, 1))
        self.results_grid = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        self.scroll_view.add_widget(self.results_grid)
        
        main_layout.add_widget(self.scroll_view)
        
        self.data_rows = []
        self.headers = []
        
        # Ilova to'liq ochilib olgandan so'ng 1.5 soniya o'tib Excel o'qiladi
        Clock.schedule_once(self.load_excel_data, 1.5)
        
        return main_layout

    def get_excel_path(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "2-TB ZIP platalar 31.07.2025.xlsx")

    def load_excel_data(self, dt):
        excel_file = self.get_excel_path()
        
        if not os.path.exists(excel_file):
            self.show_message(f"Fayl topilmadi:\n{excel_file}")
            return

        try:
            import openpyxl
            wb = openpyxl.load_workbook(excel_file, read_only=True, data_only=True)
            sheet = wb.active
            
            rows = []
            for row in sheet.iter_rows(values_only=True):
                rows.append(row)
                
            wb.close()

            if rows:
                self.headers = [str(h) if h is not None else "" for h in rows[0]]
                self.data_rows = rows[1:]
                self.show_message(f"Baza yuklandi! Jami platalar: {len(self.data_rows)} ta")
            else:
                self.show_message("Excel fayli bo'sh!")
        except Exception as e:
            self.show_message(f"Xatolik yuz berdi:\n{str(e)}")

    def show_message(self, text):
        self.results_grid.clear_widgets()
        lbl = Label(
            text=text,
            size_hint_y=None,
            height=150,
            font_size='16sp',
            halign='center',
            valign='middle'
        )
        lbl.bind(size=lbl.setter('text_size'))
        self.results_grid.add_widget(lbl)

    def search_data(self, instance):
        query = self.search_input.text.strip().lower()
        self.results_grid.clear_widgets()

        if not query:
            self.show_message("Qidirish uchun biror matn kiriting.")
            return

        if not self.data_rows:
            self.show_message("Baza tayyor emas yoki bo'sh.")
            return

        count = 0
        for row in self.data_rows:
            row_str = " ".join([str(val) for val in row if val is not None]).lower()
            if query in row_str:
                count += 1
                card_lines = []
                for idx, val in enumerate(row):
                    if val is not None and str(val).strip():
                        h_name = self.headers[idx] if idx < len(self.headers) else f"Ustun {idx+1}"
                        card_lines.append(f"[b]{h_name}:[/b] {val}")
                
                card_text = "\n".join(card_lines)
                
                card = Label(
                    text=card_text,
                    markup=True,
                    size_hint_y=None,
                    font_size='15sp',
                    padding=(12, 12)
                )
                card.bind(texture_size=lambda inst, val: setattr(inst, 'height', max(val[1] + 30, 80)))
                card.bind(size=card.setter('text_size'))
                self.results_grid.add_widget(card)

        if count == 0:
            self.show_message("Mos keluvchi plata topilmadi.")

if __name__ == '__main__':
    OmborApp().run()
