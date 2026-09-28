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
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp

# --- 1. Asosiy Menyu Ekrani ---
class MenuScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=[dp(20), dp(50), dp(20), dp(20)], 
            spacing=dp(15)
        )
        
        # Sarlavha
        title_box = BoxLayout(size_hint_y=None, height=dp(50))
        title = Label(
            text="Telekommunikatsiya ombori",
            font_size='22sp',
            bold=True,
            color=(0.07, 0.35, 0.36, 1)
        )
        title_box.add_widget(title)
        layout.add_widget(title_box)
        
        # Menyu tugmalari
        btn_platalar = Button(
            text="Platalar",
            size_hint_y=None,
            height=dp(55),
            background_normal='',
            background_color=(0.0, 0.47, 0.45, 1),
            font_size='18sp',
            bold=True
        )
        btn_platalar.bind(on_press=lambda x: self.open_search("Platalar"))
        layout.add_widget(btn_platalar)

        btn_sfp = Button(
            text="Sfp modullar",
            size_hint_y=None,
            height=dp(55),
            background_normal='',
            background_color=(0.0, 0.55, 0.55, 1),
            font_size='18sp',
            bold=True
        )
        btn_sfp.bind(on_press=lambda x: self.open_search("Sfp modullar"))
        layout.add_widget(btn_sfp)

        btn_qurilmalar = Button(
            text="Qurilmalar",
            size_hint_y=None,
            height=dp(55),
            background_normal='',
            background_color=(0.0, 0.65, 0.63, 1),
            font_size='18sp',
            bold=True
        )
        btn_qurilmalar.bind(on_press=lambda x: self.open_search("Qurilmalar"))
        layout.add_widget(btn_qurilmalar)

        # Bo'sh joy to'ldirgich
        layout.add_widget(BoxLayout())

        self.add_widget(layout)

    def open_search(self, category):
        search_screen = self.manager.get_screen('search')
        search_screen.set_category(category)
        self.manager.current = 'search'


# --- 2. Qidiruv va Ro'yxat Ekrani ---
class SearchScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        
        self.layout = BoxLayout(
            orientation='vertical', 
            padding=[dp(12), dp(40), dp(12), dp(12)], 
            spacing=dp(10)
        )
        
        # Ortga qaytish paneli
        top_bar = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(45), spacing=dp(10))
        
        btn_back = Button(
            text="< Ortga",
            size_hint_x=0.25,
            background_color=(0.4, 0.4, 0.4, 1)
        )
        btn_back.bind(on_press=self.go_back)
        
        self.title_label = Label(
            text="Qidiruv",
            font_size='18sp',
            bold=True,
            size_hint_x=0.75,
            halign='left',
            valign='middle'
        )
        self.title_label.bind(size=self.title_label.setter('text_size'))
        
        top_bar.add_widget(btn_back)
        top_bar.add_widget(self.title_label)
        self.layout.add_widget(top_bar)
        
        # Qidiruv oynasi
        search_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(50), spacing=dp(8))
        
        self.search_input = TextInput(
            hint_text="Qidiruv matnini kiriting...",
            multiline=False,
            font_size='16sp',
            size_hint_x=0.72,
            padding=[dp(10), dp(12), dp(10), dp(12)]
        )
        self.search_input.bind(on_text_validate=self.search_data)
        
        search_button = Button(
            text="Qidirish",
            size_hint_x=0.28,
            background_color=(0.12, 0.53, 0.9, 1),
            font_size='15sp',
            bold=True
        )
        search_button.bind(on_press=self.search_data)
        
        search_layout.add_widget(self.search_input)
        search_layout.add_widget(search_button)
        self.layout.add_widget(search_layout)
        
        # Natijalar ro'yxati
        self.scroll_view = ScrollView(size_hint=(1, 1))
        self.results_grid = GridLayout(cols=1, spacing=dp(12), size_hint_y=None)
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        self.scroll_view.add_widget(self.results_grid)
        
        self.layout.add_widget(self.scroll_view)
        self.add_widget(self.layout)

    def set_category(self, category_name):
        self.current_category = category_name
        self.title_label.text = f"Bo'lim: {category_name}"
        self.search_input.text = ""
        self.search_data(None)

    def go_back(self, instance):
        self.manager.current = 'menu'

    def get_db_path(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "baza.db")

    def show_message(self, text):
        self.results_grid.clear_widgets()
        lbl = Label(
            text=text,
            size_hint_y=None,
            height=dp(120),
            font_size='16sp',
            halign='center',
            valign='middle'
        )
        lbl.bind(size=lbl.setter('text_size'))
        self.results_grid.add_widget(lbl)

    def search_data(self, instance):
        query = self.search_input.text.strip()
        self.results_grid.clear_widgets()

        db_file = self.get_db_path()
        if not os.path.exists(db_file):
            self.show_message("Baza fayli (baza.db) topilmadi!")
            return

        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(platalar)")
            columns_info = cursor.fetchall()
            headers = [col[1] for col in columns_info if not str(col[1]).startswith("Unnamed")]
            
            if query:
                where_clauses = [f'"{col}" LIKE ?' for col in headers]
                sql_query = f"SELECT * FROM platalar WHERE ({' OR '.join(where_clauses)})"
                params = [f"%{query}%"] * len(headers)
            else:
                sql_query = "SELECT * FROM platalar LIMIT 50"
                params = []
            
            cursor.execute(sql_query, params)
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                self.show_message("Mos keluvchi ma'lumot topilmadi.")
                return

            for row in rows:
                card_lines = []
                for idx, col_name in enumerate(headers):
                    if idx < len(row):
                        val = row[idx]
                        if val is not None and str(val).strip() != "":
                            card_lines.append(f"[b]{col_name}:[/b] {val}")
                
                card_text = "\n".join(card_lines)
                
                card = Label(
                    text=card_text,
                    markup=True,
                    size_hint_y=None,
                    size_hint_x=1,
                    font_size='14sp',
                    padding=[dp(12), dp(12)],
                    halign='left',
                    valign='top'
                )
                card.bind(width=lambda inst, w: setattr(inst, 'text_size', (w - dp(24), None)))
                card.bind(texture_size=lambda inst, v: setattr(inst, 'height', max(v[1] + dp(24), dp(80))))
                
                self.results_grid.add_widget(card)

        except Exception as e:
            self.show_message(f"Qidiruvda xatolik yuz berdi:\n{str(e)}")


# --- 3. Asosiy Ilova ---
class OmborApp(App):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(MenuScreen(name='menu'))
        sm.add_widget(SearchScreen(name='search'))
        return sm

if __name__ == '__main__':
    OmborApp().run()
