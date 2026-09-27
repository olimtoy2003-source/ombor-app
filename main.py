import os
import sys
import sqlite3
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout

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
        
        # ScrollView va Natijalar paneli
        self.scroll_view = ScrollView(size_hint=(1, 1))
        self.results_grid = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        self.scroll_view.add_widget(self.results_grid)
        
        main_layout.add_widget(self.scroll_view)
        
        self.check_db_status()
        
        return main_layout

    def get_db_path(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "baza.db")

    def check_db_status(self):
        db_file = self.get_db_path()
        if not os.path.exists(db_file):
            self.show_message(f"Baza fayli topilmadi:\n{db_file}")
        else:
            self.show_message("Ma'lumotlar bazasi tayyor. Qidirish uchun matn kiriting.")

    def show_message(self, text):
        self.results_grid.clear_widgets()
        lbl = Label(
            text=text,
            size_hint_y=None,
            height=120,
            font_size='16sp',
            halign='center',
            valign='middle'
        )
        lbl.bind(size=lbl.setter('text_size'))
        self.results_grid.add_widget(lbl)

    def search_data(self, instance):
        query = self.search_input.text.strip()
        self.results_grid.clear_widgets()

        if not query:
            self.show_message("Qidirish uchun biror matn kiriting.")
            return

        db_file = self.get_db_path()
        if not os.path.exists(db_file):
            self.show_message("Baza fayli (baza.db) topilmadi!")
            return

        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(platalar)")
            columns_info = cursor.fetchall()
            headers = [col[1] for col in columns_info]
            
            where_clauses = [f'"{col}" LIKE ?' for col in headers]
            sql_query = f"SELECT * FROM platalar WHERE {' OR '.join(where_clauses)}"
            
            search_param = f"%{query}%"
            params = [search_param] * len(headers)
            
            cursor.execute(sql_query, params)
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                self.show_message("Mos keluvchi plata topilmadi.")
                return

            for row in rows:
                card_lines = []
                for idx, val in enumerate(row):
                    if val is not None and str(val).strip():
                        h_name = headers[idx] if idx < len(headers) else f"Ustun {idx+1}"
                        card_lines.append(f"[b]{h_name}:[/b] {val}")
                
                card_text = "\n".join(card_lines)
                
                card = Label(
                    text=card_text,
                    markup=True,
                    size_hint_y=None,
                    font_size='15sp',
                    padding=(12, 12)
                )
                card.bind(texture_size=lambda inst, v: setattr(inst, 'height', max(v[1] + 30, 80)))
                card.bind(size=card.setter('text_size'))
                self.results_grid.add_widget(card)

        except Exception as e:
            self.show_message(f"Qidiruvda xatolik yuz berdi:\n{str(e)}")

if __name__ == '__main__':
    OmborApp().run()
