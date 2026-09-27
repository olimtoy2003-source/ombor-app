import os
import sys
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window

class OmborApp(App):
    def build(self):
        self.title = "Ombor - ZIP Platalar Baza"
        
        # Asosiy ekran joylashuvi (Vertical Layout)
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Qidiruv paneli (TextInput + Button)
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
        
        # Natijalarni ko'rsatish uchun ScrollView
        self.scroll_view = ScrollView(size_hint=(1, 1))
        self.results_grid = GridLayout(cols=1, spacing=5, size_hint_y=None)
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        self.scroll_view.add_widget(self.results_grid)
        
        main_layout.add_widget(self.scroll_view)
        
        # Excel ma'lumotlarini yuklash
        self.load_excel_data()
        
        return main_layout

    def get_excel_path(self):
        """Android va Windows muhitida Excel fayl yo'lini to'g'ri aniqlash"""
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        return os.path.join(base_dir, "2-TB ZIP platalar 31.07.2025.xlsx")

    def load_excel_data(self):
        """Excel faylni o'qish va xatoliklarni ushlash"""
        excel_file = self.get_excel_path()
        
        if os.path.exists(excel_file):
            try:
                # Excel faylni barcha ustunlari bilan o'qish
                self.df = pd.read_excel(excel_file)
                # Barcha bo'sh kataklarni (NaN) bo'sh matn bilan almashtirish
                self.df = self.df.fillna("")
                self.show_message(f"Baza muvaffaqiyatli yuklandi! Jami yozuvlar: {len(self.df)}")
            except Exception as e:
                self.df = None
                self.show_message(f"Excel faylni o'qishda xatolik:\n{str(e)}")
        else:
            self.df = None
            self.show_message(f"Fayl topilmadi:\n{excel_file}\n\nFaylni GitHub'ga yuklaganingizga ishonch hosil qiling!")

    def show_message(self, text):
        """Ekran markazida xabar ko'rsatish"""
        self.results_grid.clear_widgets()
        lbl = Label(
            text=text,
            size_hint_y=None,
            height=100,
            font_size='16sp',
            halign='center',
            valign='middle'
        )
        lbl.bind(size=lbl.setter('text_size'))
        self.results_grid.add_widget(lbl)

    def search_data(self, instance):
        """Qidiruv so'rovi bo'yicha ma'lumotlarni filtrlash"""
        if self.df is None or self.df.empty:
            self.show_message("Ma'lumotlar bazasi yuklanmagan!")
            return

        query = self.search_input.text.strip().lower()
        self.results_grid.clear_widgets()

        if not query:
            self.show_message("Qidiruv uchun biror matn kiriting.")
            return

        # Barcha ustunlar bo'yicha mos keladigan satrlarni qidirish
        mask = self.df.apply(lambda row: row.astype(str).str.lower().str.contains(query).any(), axis=1)
        filtered_df = self.df[mask]

        if filtered_df.empty:
            self.show_message("Hech qanday ma'lumot topilmadi.")
            return

        # Topilgan natijalarni ekranga chiqarish
        for index, row in filtered_df.iterrows():
            row_text = []
            for col_name, val in row.items():
                if str(val).strip():
                    row_text.append(f"[b]{col_name}:[/b] {val}")
            
            card_text = "\n".join(row_text)
            
            card = Label(
                text=card_text,
                markup=True,
                size_hint_y=None,
                font_size='15sp',
                padding=(10, 10),
                color=(1, 1, 1, 1)
            )
            # Matn balandligiga qarab kartochka balandligini moslash
            card.bind(texture_size=lambda instance, value: setattr(instance, 'height', max(value[1] + 20, 60)))
            card.bind(size=card.setter('text_size'))
            
            self.results_grid.add_widget(card)

if __name__ == '__main__':
    OmborApp().run()
