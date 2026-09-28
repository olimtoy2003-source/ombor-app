from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle, Line, Rectangle, PushMatrix, PopMatrix, Rotate
from kivy.clock import Clock
import pandas as pd
import os
import subprocess
import platform
from datetime import datetime

# Kompyuterda telefon oynasi ko'rinishini hosil qilish
Window.size = (400, 750)
Window.resizable = False
Window.clearcolor = (0.95, 0.96, 0.98, 1)

EXCEL_PATH = "2-TB ZIP platalar 31.07.2025.xlsx"

def get_emoji_font():
    """Windows va boshqa Tizimlarda emojilar to'g'ri ko'rinishi uchun shrift yo'li"""
    if platform.system() == 'Windows':
        font_p = "C:\\Windows\\Fonts\\seguiemj.ttf"
        if os.path.exists(font_p):
            return font_p
    return None

EMOJI_FONT = get_emoji_font()

def load_excel_database():
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=0)
            df.columns = df.columns.str.strip()
            if "Tekshirilgan" not in df.columns:
                df["Tekshirilgan"] = ""
            if "Holati" not in df.columns:
                df["Holati"] = "Soz"
            
            if 'Plata' in df.columns:
                df['Plata'] = df['Plata'].astype(str).str.upper()
            if 'Zavod raqami' in df.columns:
                df['Zavod raqami'] = df['Zavod raqami'].astype(str).str.upper()
                
            return df
        except Exception as e:
            print("Xatolik:", e)
            return None
    return None

def save_to_excel_by_category(new_row_data):
    try:
        df = load_excel_database()
        if df is not None:
            if 'Plata' in new_row_data and isinstance(new_row_data['Plata'], str):
                new_row_data['Plata'] = new_row_data['Plata'].upper()
            if 'Zavod raqami' in new_row_data and isinstance(new_row_data['Zavod raqami'], str):
                new_row_data['Zavod raqami'] = new_row_data['Zavod raqami'].upper()

            chosen_qurilma = new_row_data.get('Qurilma', '')
            df['Normalized_Qurilma'] = df['Qurilma'].apply(normalize_qurilma_name)
            
            qurilma_mask = df['Normalized_Qurilma'].astype(str).str.strip() == chosen_qurilma.strip()
            
            if qurilma_mask.any():
                matching_indices = df[qurilma_mask].index
                insert_pos = matching_indices[-1] + 1
            else:
                insert_pos = len(df)

            df.drop(columns=['Normalized_Qurilma'], inplace=True, errors='ignore')

            df_before = df.iloc[:insert_pos].copy()
            df_after = df.iloc[insert_pos:].copy()
            
            new_df = pd.DataFrame([new_row_data])
            updated_df = pd.concat([df_before, new_df, df_after], ignore_index=True)
            
            if '№' in updated_df.columns:
                updated_df['№'] = range(1, len(updated_df) + 1)

            updated_df.to_excel(EXCEL_PATH, index=False)
            return True
    except Exception as e:
        print("Saqlashda xatolik:", e)
    return False

def open_file_externally(filename):
    try:
        if platform.system() == 'Windows':
            os.startfile(filename)
        elif platform.system() == 'Darwin':
            subprocess.call(('open', filename))
        else:
            subprocess.call(('xdg-open', filename))
    except Exception as e:
        print("Faylni ochishda xatolik:", e)

def normalize_qurilma_name(q_name):
    if not isinstance(q_name, str):
        return str(q_name)
    q_upper = q_name.upper().strip()
    
    if "OSN6800" in q_upper:
        return None
        
    if any(x in q_upper for x in ["OSN2500", "OSN3500", "OSN8800", "OSN9800"]):
        return "OSN2500/3500/8800/9800"
    if "SMS" in q_upper or "U-NODE" in q_upper:
        return "SMS / U-Node"
    return q_name.strip()

# Yuklanish uchun aylanma yumaloqcha (Spinner)
class LoadingSpinner(BoxLayout):
    def __init__(self, **kwargs):
        super(LoadingSpinner, self).__init__(**kwargs)
        self.size_hint = (None, None)
        self.size = (35, 35)
        self.angle = 0
        with self.canvas:
            Color(0.0, 0.45, 0.45, 1)
            self.p_matrix = PushMatrix()
            self.rot = Rotate(angle=0, origin=self.center)
            self.line = Line(circle=(0, 0, 14, 0, 270), width=3)
            self.pop_matrix = PopMatrix()
        self.bind(pos=self.update_canvas, size=self.update_canvas)
        Clock.schedule_interval(self.animate, 1 / 30)

    def update_canvas(self, *args):
        self.rot.origin = (self.center_x, self.center_y)

    def animate(self, dt):
        self.angle = (self.angle - 10) % 360
        self.rot.angle = self.angle

class CellLabel(Label):
    def __init__(self, **kwargs):
        super(CellLabel, self).__init__(**kwargs)
        self.color = (0.1, 0.1, 0.1, 1)
        self.font_size = 11
        self.halign = 'center'
        self.valign = 'middle'
        if EMOJI_FONT:
            self.font_name = EMOJI_FONT
        self.bind(size=self.update_canvas, pos=self.update_canvas)

    def update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(1, 1, 1, 1)
            Rectangle(pos=self.pos, size=self.size)
            Color(0.85, 0.88, 0.92, 1)
            Line(rectangle=(self.x, self.y, self.width, self.height), width=1)

class HeaderCell(Label):
    def __init__(self, **kwargs):
        super(HeaderCell, self).__init__(**kwargs)
        self.color = (0, 0.4, 0.4, 1)
        self.font_size = 12
        self.bold = True
        self.halign = 'center'
        self.valign = 'middle'
        if EMOJI_FONT:
            self.font_name = EMOJI_FONT
        self.bind(size=self.update_canvas, pos=self.update_canvas)

    def update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(0.9, 0.93, 0.96, 1)
            Rectangle(pos=self.pos, size=self.size)
            Color(0.75, 0.8, 0.85, 1)
            Line(rectangle=(self.x, self.y, self.width, self.height), width=1.5)

class SoliqButton(Button):
    def __init__(self, bg_color=(0.0, 0.45, 0.45, 1), **kwargs):
        super(SoliqButton, self).__init__(**kwargs)
        self.background_color = (0, 0, 0, 0)
        self.bg_color = bg_color
        self.color = (1, 1, 1, 1)
        self.font_size = 14
        self.bold = True
        if EMOJI_FONT:
            self.font_name = EMOJI_FONT
        self.bind(pos=self.update_canvas, size=self.update_canvas, text=self.adjust_font_size, size_hint=self.adjust_font_size)
        self.halign = 'center'
        self.valign = 'middle'
        self.bind(state=self.on_state_change)

    def adjust_font_size(self, *args):
        self.text_size = (self.width - 10, None)
        if len(self.text) > 25:
            self.font_size = 10
        elif len(self.text) > 15:
            self.font_size = 12
        else:
            self.font_size = 14

    def update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            if self.state == 'down':
                Color(self.bg_color[0]*0.8, self.bg_color[1]*0.8, self.bg_color[2]*0.8, 1)
            else:
                Color(*self.bg_color)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[10])

    def on_state_change(self, instance, value):
        self.update_canvas()

class QurilmaSelectButton(ToggleButton):
    def __init__(self, q_name, **kwargs):
        super(QurilmaSelectButton, self).__init__(**kwargs)
        self.q_name = q_name
        self.background_color = (0, 0, 0, 0)
        self.color = (1, 1, 1, 1)
        self.font_size = 13
        self.bold = True
        if EMOJI_FONT:
            self.font_name = EMOJI_FONT
        self.size_hint_y = None
        self.height = 36
        self.update_text()
        self.bind(pos=self.update_canvas, size=self.update_canvas, state=self.on_toggle_state)

    def on_toggle_state(self, instance, value):
        self.update_text()
        self.update_canvas()

    def update_text(self):
        if self.state == 'down':
            self.text = f"✅ {self.q_name}"
        else:
            self.text = f"◽ {self.q_name}"

    def update_canvas(self, *args):
        self.canvas.before.clear()
        with self.canvas.before:
            if self.state == 'down':
                Color(0.0, 0.55, 0.45, 1)
            else:
                Color(0.4, 0.5, 0.6, 1)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[6])

class AdaptiveLabel(Label):
    def __init__(self, **kwargs):
        super(AdaptiveLabel, self).__init__(**kwargs)
        if EMOJI_FONT:
            self.font_name = EMOJI_FONT
        self.bind(width=self.update_text_size)

    def update_text_size(self, *args):
        self.text_size = (self.width, None)

class TopBar(BoxLayout):
    def __init__(self, title_text, back_callback=None, **kwargs):
        super(TopBar, self).__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = 50
        self.padding = [10, 5, 10, 5]
        self.spacing = 10
        
        with self.canvas.before:
            Color(1, 1, 1, 1)
            self.rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self.update_rect, size=self.update_rect)

        if back_callback:
            btn_back = Button(
                text='‹', 
                font_size=28, 
                bold=True, 
                size_hint_x=None, 
                width=40,
                background_color=(0, 0, 0, 0),
                color=(0.0, 0.45, 0.45, 1)
            )
            btn_back.bind(on_press=back_callback)
            self.add_widget(btn_back)
        else:
            spacer = Label(size_hint_x=None, width=10)
            self.add_widget(spacer)

        lbl = AdaptiveLabel(
            text=title_text, 
            font_size=16, 
            bold=True, 
            color=(0.0, 0.45, 0.45, 1), 
            halign='center', 
            valign='middle'
        )
        lbl.bind(size=lbl.setter('text_size'))
        self.add_widget(lbl)
        
        spacer_right = Label(size_hint_x=None, width=40 if back_callback else 10)
        self.add_widget(spacer_right)

    def update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size

# --- 1. ASOSIY BOSH MENYU ---
class MainMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)

    def __init__(self, **kwargs):
        super(MainMenuScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=20, spacing=20)
        
        layout.add_widget(TopBar('Telekommunikatsiya ombori'))

        btn_platalar = SoliqButton(text='🧩 Platalar', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=65)
        btn_platalar.bind(on_press=lambda x: setattr(self.manager, 'current', 'platalar_menu'))
        layout.add_widget(btn_platalar)

        btn_sfp = SoliqButton(text='🔌 Sfp modullar', bg_color=(0.0, 0.55, 0.55, 1), size_hint_y=None, height=65)
        btn_sfp.bind(on_press=lambda x: setattr(self.manager, 'current', 'sfp_menu'))
        layout.add_widget(btn_sfp)

        btn_qurilmalar = SoliqButton(text='🖥 Qurilmalar', bg_color=(0.0, 0.65, 0.65, 1), size_hint_y=None, height=65)
        btn_qurilmalar.bind(on_press=lambda x: setattr(self.manager, 'current', 'qurilmalar_menu'))
        layout.add_widget(btn_qurilmalar)

        btn_excel = SoliqButton(text='📥 Excelga yuklab olish', bg_color=(0.2, 0.4, 0.6, 1), size_hint_y=None, height=65)
        btn_excel.bind(on_press=lambda x: setattr(self.manager, 'current', 'excel_menu'))
        layout.add_widget(btn_excel)

        layout.add_widget(Label(size_hint_y=1))
        self.add_widget(layout)

# --- EXCELGA YUKLAB OLISH ASOSIY MENYUSI ---
class ExcelMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.saved_file = None
        self.btn_open.opacity = 0
        self.btn_open.disabled = True
        self.msg_label.text = ""
        self.spinner.opacity = 0

    def __init__(self, **kwargs):
        super(ExcelMenuScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        
        layout.add_widget(TopBar('Excelga yuklab olish', back_callback=lambda x: setattr(self.manager, 'current', 'main_menu')))

        btn_platalar_exp = SoliqButton(text='🧩 Platalar bo\'yicha', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=55)
        btn_platalar_exp.bind(on_press=lambda x: setattr(self.manager, 'current', 'excel_platalar_menu'))
        layout.add_widget(btn_platalar_exp)

        btn_sfp_exp = SoliqButton(text='🔌 SFP modullar bo\'yicha', bg_color=(0.0, 0.55, 0.55, 1), size_hint_y=None, height=55)
        btn_sfp_exp.bind(on_press=lambda x: self.start_action(self.export_sfp))
        layout.add_widget(btn_sfp_exp)

        btn_qur_exp = SoliqButton(text='🖥 Qurilma bo\'yicha', bg_color=(0.0, 0.65, 0.65, 1), size_hint_y=None, height=55)
        btn_qur_exp.bind(on_press=lambda x: self.start_action(self.export_qurilma))
        layout.add_widget(btn_qur_exp)

        msg_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30, spacing=10)
        self.msg_label = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=30, color=(0, 0.6, 0.3, 1), halign='left')
        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0
        msg_box.add_widget(self.msg_label)
        msg_box.add_widget(self.spinner)
        layout.add_widget(msg_box)

        self.btn_open = SoliqButton(text='📂 Faylni ochish', bg_color=(0.0, 0.5, 0.3, 1), size_hint_y=None, height=45)
        self.btn_open.opacity = 0
        self.btn_open.disabled = True
        self.btn_open.bind(on_press=lambda x: open_file_externally(self.saved_file) if self.saved_file else None)
        layout.add_widget(self.btn_open)

        layout.add_widget(Label(size_hint_y=1))
        self.add_widget(layout)

    def start_action(self, func):
        self.spinner.opacity = 1
        self.msg_label.text = "Yuklanmoqda..."
        Clock.schedule_once(lambda dt: func(), 0.3)

    def export_sfp(self):
        df = load_excel_database()
        if df is not None:
            time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.saved_file = f"SFP_modullar_{time_str}.xlsx"
            df.to_excel(self.saved_file, index=False)
            self.msg_label.color = (0, 0.6, 0.3, 1)
            self.msg_label.text = f"Saqlandi: {self.saved_file}"
            self.btn_open.opacity = 1
            self.btn_open.disabled = False
        self.spinner.opacity = 0

    def export_qurilma(self):
        df = load_excel_database()
        if df is not None and 'Qurilma' in df.columns:
            time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.saved_file = f"Qurilmalar_{time_str}.xlsx"
            df.to_excel(self.saved_file, index=False)
            self.msg_label.color = (0, 0.6, 0.3, 1)
            self.msg_label.text = f"Saqlandi: {self.saved_file}"
            self.btn_open.opacity = 1
            self.btn_open.disabled = False
        self.spinner.opacity = 0

class ExcelPlatalarMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.saved_file = None
        self.btn_open.opacity = 0
        self.btn_open.disabled = True
        self.msg_label.text = ""
        self.spinner.opacity = 0

    def __init__(self, **kwargs):
        super(ExcelPlatalarMenuScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        
        layout.add_widget(TopBar('Platalar bo\'yicha eksport', back_callback=lambda x: setattr(self.manager, 'current', 'excel_menu')))

        btn_umumiy = SoliqButton(text='📦 Umumiy (Hamma ombor)', bg_color=(0.2, 0.4, 0.6, 1), size_hint_y=None, height=50)
        btn_umumiy.bind(on_press=lambda x: self.start_action('all'))
        layout.add_widget(btn_umumiy)

        btn_fergana = SoliqButton(text='🏢 Fargona', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=50)
        btn_fergana.bind(on_press=lambda x: self.start_action('FarAMTS ombor'))
        layout.add_widget(btn_fergana)

        btn_namangan = SoliqButton(text='🏢 Namangan', bg_color=(0.2, 0.5, 0.6, 1), size_hint_y=None, height=50)
        btn_namangan.bind(on_press=lambda x: self.start_action('NamAMTS ombor'))
        layout.add_widget(btn_namangan)

        btn_andijon = SoliqButton(text='🏢 Andijon', bg_color=(0.3, 0.6, 0.5, 1), size_hint_y=None, height=50)
        btn_andijon.bind(on_press=lambda x: self.start_action('AndAMTS ombor'))
        layout.add_widget(btn_andijon)

        msg_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30, spacing=10)
        self.msg_label = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=30, color=(0, 0.6, 0.3, 1), halign='left')
        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0
        msg_box.add_widget(self.msg_label)
        msg_box.add_widget(self.spinner)
        layout.add_widget(msg_box)

        self.btn_open = SoliqButton(text='📂 Faylni ochish', bg_color=(0.0, 0.5, 0.3, 1), size_hint_y=None, height=45)
        self.btn_open.opacity = 0
        self.btn_open.disabled = True
        self.btn_open.bind(on_press=lambda x: open_file_externally(self.saved_file) if self.saved_file else None)
        layout.add_widget(self.btn_open)

        layout.add_widget(Label(size_hint_y=1))
        self.add_widget(layout)

    def start_action(self, mode):
        self.spinner.opacity = 1
        self.msg_label.text = "Yuklanmoqda..."
        Clock.schedule_once(lambda dt: self.export_excel(mode), 0.3)

    def export_excel(self, mode):
        df = load_excel_database()
        if df is None:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Excel bazasi topilmadi!"
            self.spinner.opacity = 0
            return

        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        if mode == 'all':
            export_df = df
            self.saved_file = f"Platalar_Umumiy_{time_str}.xlsx"
        else:
            if "Joylashgan o'rni" in df.columns:
                export_df = df[df["Joylashgan o'rni"].astype(str).str.strip().str.lower() == mode.lower()]
            else:
                export_df = pd.DataFrame()
            self.saved_file = f"Plata_{mode.replace(' ', '_')}_{time_str}.xlsx"

        if export_df.empty:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Tanlangan hudud bo'yicha ma'lumot topilmadi!"
            self.btn_open.opacity = 0
            self.btn_open.disabled = True
            self.spinner.opacity = 0
            return

        try:
            export_df.to_excel(self.saved_file, index=False)
            self.msg_label.color = (0, 0.6, 0.3, 1)
            self.msg_label.text = f"Saqlandi: {self.saved_file}"
            self.btn_open.opacity = 1
            self.btn_open.disabled = False
        except Exception as e:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = f"Xatolik: {e}"
        self.spinner.opacity = 0

# --- 2. PLATALAR MENYUSI ---
class PlatalarMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)

    def __init__(self, **kwargs):
        super(PlatalarMenuScreen, self).__init__(**kwargs)
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))

        layout.add_widget(TopBar('Platalar bo\'limi', back_callback=lambda x: setattr(self.manager, 'current', 'main_menu')))

        btn_platalar = SoliqButton(text='📄 Platalar ro\'yxati', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=55)
        btn_platalar.bind(on_press=lambda x: setattr(self.manager, 'current', 'platalar_screen'))
        layout.add_widget(btn_platalar)

        btn_qoshish = SoliqButton(text='➕ Yangi plata qo\'shish', bg_color=(0.1, 0.6, 0.5, 1), size_hint_y=None, height=55)
        btn_qoshish.bind(on_press=lambda x: setattr(self.manager, 'current', 'add_screen'))
        layout.add_widget(btn_qoshish)

        btn_joy = SoliqButton(text='🏗 Stansiyaga o\'rnatish', bg_color=(0.85, 0.5, 0.1, 1), size_hint_y=None, height=55)
        btn_joy.bind(on_press=lambda x: setattr(self.manager, 'current', 'move_screen'))
        layout.add_widget(btn_joy)

        btn_inv = SoliqButton(text='📋 Invertarizatsiya', bg_color=(0.5, 0.3, 0.7, 1), size_hint_y=None, height=55)
        btn_inv.bind(on_press=lambda x: setattr(self.manager, 'current', 'inv_screen'))
        layout.add_widget(btn_inv)

        scroll.add_widget(layout)
        self.add_widget(scroll)

class SFPMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)

    def __init__(self, **kwargs):
        super(SFPMenuScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        layout.add_widget(TopBar('Sfp modullar bo\'limi', back_callback=lambda x: setattr(self.manager, 'current', 'main_menu')))
        layout.add_widget(AdaptiveLabel(text='Ushbu bo\'lim tez kunda ishga tushadi...', font_size=16, color=(0.3, 0.3, 0.3, 1), halign='center'))
        layout.add_widget(Label(size_hint_y=1))
        self.add_widget(layout)

class QurilmalarMenuScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)

    def __init__(self, **kwargs):
        super(QurilmalarMenuScreen, self).__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        layout.add_widget(TopBar('Qurilmalar bo\'limi', back_callback=lambda x: setattr(self.manager, 'current', 'main_menu')))
        layout.add_widget(AdaptiveLabel(text='Ushbu bo\'lim tez kunda ishga tushadi...', font_size=16, color=(0.3, 0.3, 0.3, 1), halign='center'))
        layout.add_widget(Label(size_hint_y=1))
        self.add_widget(layout)

class PlatalarScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.yukla_qurilmalar()

    def __init__(self, **kwargs):
        super(PlatalarScreen, self).__init__(**kwargs)
        self.layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        
        self.layout.add_widget(TopBar('Platalar ro\'yxati', back_callback=lambda x: setattr(self.manager, 'current', 'platalar_menu')))
        self.layout.add_widget(AdaptiveLabel(text='Qurilmani tanlang:', font_size=16, bold=True, size_hint_y=None, height=30, color=(0.2, 0.2, 0.2, 1), halign='left'))

        self.scroll = ScrollView()
        self.list_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=10)
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        self.scroll.add_widget(self.list_layout)
        self.layout.add_widget(self.scroll)
        
        self.add_widget(self.layout)

    def yukla_qurilmalar(self):
        self.list_layout.clear_widgets()
        df = load_excel_database()
        if df is not None and 'Qurilma' in df.columns:
            raw_qurilmalar = df['Qurilma'].dropna().unique()
            qurilmalar = sorted(list(set(normalize_qurilma_name(q) for q in raw_qurilmalar if normalize_qurilma_name(q) is not None)))
            for q in qurilmalar:
                btn = SoliqButton(text=f"🖥 {str(q)}", bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=50)
                btn.bind(on_press=lambda x, q_nomi=str(q): self.open_qurilma(q_nomi))
                self.list_layout.add_widget(btn)

    def open_qurilma(self, qurilma_nomi):
        detal_ekran = self.manager.get_screen('detal_screen')
        detal_ekran.set_qurilma(qurilma_nomi)
        self.manager.current = 'detal_screen'

class PlatalarDetalScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)

    def __init__(self, **kwargs):
        super(PlatalarDetalScreen, self).__init__(**kwargs)
        self.current_qurilma = ""
        self.layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        self.layout.add_widget(TopBar('Platalar jadvali', back_callback=lambda x: setattr(self.manager, 'current', 'platalar_screen')))
        
        self.lbl_sarlavha = AdaptiveLabel(text='Qurilma:', font_size=15, bold=True, size_hint_y=None, height=30, color=(0.2, 0.2, 0.2, 1), halign='left')
        self.layout.add_widget(self.lbl_sarlavha)

        search_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40, spacing=5)
        self.search_input = TextInput(hint_text='Plata nomi yoki zavod raqami oxirgi 4 ta raqami...', multiline=False)
        self.search_input.bind(text=self.filter_data)
        
        btn_clear = SoliqButton(text='❌', bg_color=(0.8, 0.3, 0.3, 1), size_hint_x=None, width=45, height=40)
        btn_clear.bind(on_press=lambda x: setattr(self.search_input, 'text', ''))

        search_layout.add_widget(self.search_input)
        search_layout.add_widget(btn_clear)
        self.layout.add_widget(search_layout)

        col_widths = [45, 130, 150, 110, 150, 100]
        self.total_width = sum(col_widths)

        self.header_layout = GridLayout(cols=6, size_hint_y=None, height=40, size_hint_x=None, spacing=1, col_force_default=True)
        self.header_layout.width = self.total_width
        for i, w in enumerate(col_widths):
            self.header_layout.cols_minimum[i] = w

        headers = ['№', 'Plata', 'Zavod raqami', 'Inv raqami', "Joyi", "Tekshirildi"]
        for i, h in enumerate(headers):
            self.header_layout.add_widget(HeaderCell(text=h, size_hint_x=None, width=col_widths[i], size_hint_y=None, height=40))
        
        self.header_scroll = ScrollView(size_hint_y=None, height=40, do_scroll_y=False)
        self.header_scroll.add_widget(self.header_layout)
        self.layout.add_widget(self.header_scroll)

        self.scroll = ScrollView()
        self.grid_layout = GridLayout(cols=6, size_hint_y=None, size_hint_x=None, spacing=1, col_force_default=True)
        self.grid_layout.width = self.total_width
        for i, w in enumerate(col_widths):
            self.grid_layout.cols_minimum[i] = w
        self.grid_layout.bind(minimum_height=self.grid_layout.setter('height'))
        
        self.scroll.add_widget(self.grid_layout)
        self.layout.add_widget(self.scroll)

        self.add_widget(self.layout)

    def set_qurilma(self, qurilma_nomi):
        self.current_qurilma = qurilma_nomi
        self.lbl_sarlavha.text = f"Qurilma: {qurilma_nomi}"
        self.search_input.text = ""
        self.load_table_data()

    def load_table_data(self, search_query=""):
        self.grid_layout.clear_widgets()
        df = load_excel_database()
        if df is not None:
            df['Normalized_Qurilma'] = df['Qurilma'].apply(normalize_qurilma_name)
            filtrlangan = df[df['Normalized_Qurilma'].astype(str).str.strip() == self.current_qurilma.strip()]
            
            if search_query:
                q = search_query.lower()
                filtrlangan = filtrlangan[
                    filtrlangan['Plata'].astype(str).str.lower().str.contains(q) |
                    filtrlangan['Zavod raqami'].astype(str).str.lower().str.endswith(q) |
                    filtrlangan['Inv raqami'].astype(str).str.lower().str.contains(q)
                ]

            if not filtrlangan.empty:
                col_widths = [45, 130, 150, 110, 150, 100]
                for index, row in filtrlangan.reset_index(drop=True).iterrows():
                    plata = str(row.get('Plata', '')).upper()
                    zavod = str(row.get('Zavod raqami', '')).upper()
                    inv = str(row.get('Inv raqami', ''))
                    joy = str(row.get("Joylashgan o'rni", ''))
                    tekshirildi = str(row.get("Tekshirilgan", ''))

                    plata = '' if plata == 'NAN' else plata
                    zavod = '' if zavod == 'NAN' else zavod
                    inv = '' if inv == 'nan' else inv
                    joy = '' if joy == 'nan' else joy
                    tekshirildi = '' if tekshirildi == 'nan' else tekshirildi

                    self.grid_layout.add_widget(CellLabel(text=str(index + 1), size_hint_y=None, height=40, size_hint_x=None, width=col_widths[0]))
                    self.grid_layout.add_widget(CellLabel(text=plata, size_hint_y=None, height=40, size_hint_x=None, width=col_widths[1]))
                    self.grid_layout.add_widget(CellLabel(text=zavod, size_hint_y=None, height=40, size_hint_x=None, width=col_widths[2]))
                    self.grid_layout.add_widget(CellLabel(text=inv, size_hint_y=None, height=40, size_hint_x=None, width=col_widths[3]))
                    self.grid_layout.add_widget(CellLabel(text=joy, size_hint_y=None, height=40, size_hint_x=None, width=col_widths[4]))
                    self.grid_layout.add_widget(CellLabel(text=tekshirildi, size_hint_y=None, height=40, size_hint_x=None, width=col_widths[5]))
            else:
                lbl = CellLabel(text='Ma\'lumot topilmadi', size_hint_y=None, height=40, size_hint_x=None, width=self.total_width)
                lbl.halign = 'center'
                lbl.font_size = 13
                self.grid_layout.add_widget(lbl)

    def filter_data(self, instance, value):
        self.load_table_data(value)

# --- YANGI PLATA QO'SHISH ---
class AddPlataScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.yukla_qurilmalar()
        self.chosen_joy = ""
        self.chosen_qurilma = ""
        self.chosen_holat = "Soz"
        self.is_valid_plata = True
        self.selected_joy_lbl.text = "Tanlangan joy: (Tanlanmagan)"
        self.selected_qurilma_lbl.text = "Tanlangan qurilma: (Tanlanmagan)"
        self.selected_holat_lbl.text = "Tanlangan holat: Soz"
        self.manual_joy_input.opacity = 0
        self.manual_joy_input.disabled = True
        self.manual_joy_input.text = ""
        self.zavod_input.text = ""
        self.inv_input.text = ""
        self.plata_input.text = ""
        self.plata_sug_layout.clear_widgets()
        self.spinner.opacity = 0
        self.msg_label.text = ""
        self.update_preview()

    def __init__(self, **kwargs):
        super(AddPlataScreen, self).__init__(**kwargs)
        
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=20, spacing=10, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))

        layout.add_widget(TopBar('Yangi plata qo\'shish', back_callback=lambda x: setattr(self.manager, 'current', 'platalar_menu')))

        layout.add_widget(AdaptiveLabel(text='Qurilmani tanlang:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        
        self.qurilma_scroll = ScrollView(size_hint_y=None, height=110)
        self.qurilma_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=8)
        self.qurilma_layout.bind(minimum_height=self.qurilma_layout.setter('height'))
        self.qurilma_scroll.add_widget(self.qurilma_layout)
        layout.add_widget(self.qurilma_scroll)

        self.selected_qurilma_lbl = AdaptiveLabel(text='Tanlangan qurilma: (Tanlanmagan)', font_size=14, size_hint_y=None, height=25, color=(0.0, 0.45, 0.45, 1), halign='left')
        layout.add_widget(self.selected_qurilma_lbl)

        layout.add_widget(AdaptiveLabel(text='Plata nomi:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.plata_input = TextInput(hint_text='Avval qurilma tanlang, keyin yozing...', multiline=False, size_hint_y=None, height=40)
        self.plata_input.bind(text=self.filter_platalar)
        layout.add_widget(self.plata_input)

        self.plata_sug_scroll = ScrollView(size_hint_y=None, height=80)
        self.plata_sug_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=3)
        self.plata_sug_layout.bind(minimum_height=self.plata_sug_layout.setter('height'))
        self.plata_sug_scroll.add_widget(self.plata_sug_layout)
        layout.add_widget(self.plata_sug_scroll)

        layout.add_widget(AdaptiveLabel(text='Zavod va inventar raqami:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.zavod_input = TextInput(hint_text='Zavod raqami', multiline=False, size_hint_y=None, height=40)
        self.zavod_input.bind(text=self.on_input_changed)
        
        self.inv_input = TextInput(hint_text='Inventar raqami', multiline=False, size_hint_y=None, height=40)
        self.inv_input.bind(text=self.on_input_changed)
        
        layout.add_widget(self.zavod_input)
        layout.add_widget(self.inv_input)

        layout.add_widget(AdaptiveLabel(text='Plata holatini tanlang:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        holat_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40, spacing=10)
        
        btn_soz = SoliqButton(text='✔️ Soz', bg_color=(0.1, 0.6, 0.4, 1))
        btn_soz.bind(on_press=lambda x: self.select_holat('Soz'))
        
        btn_nosoz = SoliqButton(text='⚠️ Nosoz', bg_color=(0.8, 0.3, 0.3, 1))
        btn_nosoz.bind(on_press=lambda x: self.select_holat('Nosoz'))

        holat_layout.add_widget(btn_soz)
        holat_layout.add_widget(btn_nosoz)
        layout.add_widget(holat_layout)

        self.selected_holat_lbl = AdaptiveLabel(text='Tanlangan holat: Soz', font_size=14, size_hint_y=None, height=25, color=(0.0, 0.45, 0.45, 1), halign='left')
        layout.add_widget(self.selected_holat_lbl)

        layout.add_widget(AdaptiveLabel(text='Joylashgan o\'rnini tanlang:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        
        joy_grid = GridLayout(cols=2, size_hint_y=None, height=90, spacing=8)
        
        btn_kok = SoliqButton(text='KokOmbor', bg_color=(0.0, 0.45, 0.45, 1), height=40)
        btn_kok.bind(on_press=lambda x: self.select_joy('KokOmbor'))
        
        btn_far = SoliqButton(text='FarAMTS ombor', bg_color=(0.0, 0.5, 0.5, 1), height=40)
        btn_far.bind(on_press=lambda x: self.select_joy('FarAMTS ombor'))
        
        btn_nam = SoliqButton(text='NamAMTS ombor', bg_color=(0.2, 0.5, 0.6, 1), height=40)
        btn_nam.bind(on_press=lambda x: self.select_joy('NamAMTS ombor'))
        
        btn_and = SoliqButton(text='AndAMTS ombor', bg_color=(0.3, 0.6, 0.5, 1), height=40)
        btn_and.bind(on_press=lambda x: self.select_joy('AndAMTS ombor'))

        joy_grid.add_widget(btn_kok)
        joy_grid.add_widget(btn_far)
        joy_grid.add_widget(btn_nam)
        joy_grid.add_widget(btn_and)
        layout.add_widget(joy_grid)

        btn_boshqa = SoliqButton(text='✏️ Boshqa joy', bg_color=(0.6, 0.4, 0.4, 1), size_hint_y=None, height=40)
        btn_boshqa.bind(on_press=lambda x: self.select_joy('Boshqa joy'))
        layout.add_widget(btn_boshqa)

        self.manual_joy_input = TextInput(hint_text='Boshqa joyni qo\'lda kiriting...', multiline=False, size_hint_y=None, height=0)
        self.manual_joy_input.bind(text=self.on_input_changed)
        layout.add_widget(self.manual_joy_input)

        self.selected_joy_lbl = AdaptiveLabel(text='Tanlangan joy: (Tanlanmagan)', font_size=14, size_hint_y=None, height=25, color=(0.0, 0.45, 0.45, 1), halign='left')
        layout.add_widget(self.selected_joy_lbl)

        layout.add_widget(AdaptiveLabel(text='Kiritilayotgan ma\'lumotlar ko\'rinishi:', font_size=14, bold=True, size_hint_y=None, height=25, color=(0.0, 0.45, 0.45, 1), halign='left'))
        
        self.preview_lbl = Label(
            text='', 
            font_size=13, 
            size_hint_y=None, 
            height=130, 
            color=(0.2, 0.2, 0.2, 1), 
            halign='left',
            valign='middle',
            markup=True
        )
        if EMOJI_FONT:
            self.preview_lbl.font_name = EMOJI_FONT
        self.preview_lbl.bind(size=self.preview_lbl.setter('text_size'))
        layout.add_widget(self.preview_lbl)

        msg_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30, spacing=10)
        self.msg_label = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=30, color=(0, 0.6, 0.3, 1), halign='left')
        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0
        msg_box.add_widget(self.msg_label)
        msg_box.add_widget(self.spinner)
        layout.add_widget(msg_box)

        btn_save = SoliqButton(text='💾 Saqlash', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=50)
        btn_save.bind(on_press=lambda x: self.start_save_action())
        layout.add_widget(btn_save)

        scroll.add_widget(layout)
        self.add_widget(scroll)

    def yukla_qurilmalar(self):
        self.qurilma_layout.clear_widgets()
        df = load_excel_database()
        if df is not None and 'Qurilma' in df.columns:
            raw_qurilmalar = df['Qurilma'].dropna().unique()
            qurilmalar = sorted(list(set(normalize_qurilma_name(q) for q in raw_qurilmalar if normalize_qurilma_name(q) is not None)))
            for q in qurilmalar:
                btn = SoliqButton(text=f"🖥 {str(q)}", bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=38)
                btn.bind(on_press=lambda x, val=str(q): self.select_qurilma(val))
                self.qurilma_layout.add_widget(btn)

    def select_qurilma(self, val):
        self.chosen_qurilma = val
        self.selected_qurilma_lbl.text = f"Tanlangan qurilma: {val}"
        self.plata_input.text = ""
        self.plata_sug_layout.clear_widgets()
        self.msg_label.text = ""
        self.is_valid_plata = True
        self.update_preview()

    def select_holat(self, val):
        self.chosen_holat = val
        self.selected_holat_lbl.text = f"Tanlangan holat: {val}"
        self.update_preview()

    def select_joy(self, val):
        if val == 'Boshqa joy':
            self.chosen_joy = ''
            self.manual_joy_input.height = 40
            self.manual_joy_input.opacity = 1
            self.manual_joy_input.disabled = False
            self.selected_joy_lbl.text = "Tanlangan joy: Boshqa joy (pastga yozing)"
        else:
            self.chosen_joy = val
            self.manual_joy_input.height = 0
            self.manual_joy_input.opacity = 0
            self.manual_joy_input.disabled = True
            self.manual_joy_input.text = ""
            self.selected_joy_lbl.text = f"Tanlangan joy: {val}"
        self.update_preview()

    def on_input_changed(self, instance, value):
        self.update_preview()

    def update_preview(self):
        q = self.chosen_qurilma if self.chosen_qurilma else "—"
        p = self.plata_input.text.strip().upper() if self.plata_input.text.strip() else "—"
        z = self.zavod_input.text.strip().upper() if self.zavod_input.text.strip() else "—"
        i = self.inv_input.text.strip() if self.inv_input.text.strip() else "—"
        h = self.chosen_holat
        j = self.manual_joy_input.text.strip() if self.manual_joy_input.opacity == 1 else (self.chosen_joy if self.chosen_joy else "—")
        
        info_str = (
            f"• Qurilma: [color=004545]{q}[/color]\n"
            f"• Plata: [color=004545]{p}[/color]\n"
            f"• Zavod raqami: [color=004545]{z}[/color]\n"
            f"• Inv №: [color=004545]{i}[/color]\n"
            f"• Holati: [color=004545]{h}[/color]\n"
            f"• Saqlanish joyi: [color=004545]{j}[/color]\n"
            f"• Tekshirildi: [color=aa8800]❌ Tekshirilmagan[/color]"
        )
        self.preview_lbl.text = info_str

    def filter_platalar(self, instance, value):
        self.plata_sug_layout.clear_widgets()
        self.update_preview()
        
        if not self.chosen_qurilma:
            lbl_err = AdaptiveLabel(
                text="Avval qurilmani tanlang!", 
                font_size=13, 
                size_hint_y=None, 
                height=30, 
                color=(0.8, 0.2, 0.2, 1), 
                halign='left'
            )
            self.plata_sug_layout.add_widget(lbl_err)
            self.is_valid_plata = False
            return

        val = value.strip().lower()
        if not val:
            self.is_valid_plata = True
            return

        df = load_excel_database()
        if df is not None and 'Qurilma' in df.columns and 'Plata' in df.columns:
            df['Normalized_Qurilma'] = df['Qurilma'].apply(normalize_qurilma_name)
            qurilma_df = df[df['Normalized_Qurilma'].astype(str).str.strip() == self.chosen_qurilma.strip()]
            all_platas = qurilma_df['Plata'].dropna().astype(str).unique()
            
            matched = [p for p in all_platas if val in p.lower()]
            
            if matched:
                self.is_valid_plata = True
                for p in matched[:10]:
                    btn = SoliqButton(text=p.upper(), bg_color=(0.3, 0.5, 0.5, 1), size_hint_y=None, height=35)
                    btn.bind(on_press=lambda x, val=p: self.select_plata(val))
                    self.plata_sug_layout.add_widget(btn)
            else:
                lbl_err = AdaptiveLabel(
                    text="Bunday plata bu qurilmaga tegishli emas!", 
                    font_size=13, 
                    bold=True,
                    size_hint_y=None, 
                    height=30, 
                    color=(0.85, 0.2, 0.2, 1), 
                    halign='left'
                )
                self.plata_sug_layout.add_widget(lbl_err)
                self.is_valid_plata = False

    def select_plata(self, val):
        self.plata_input.text = val.upper()
        self.plata_sug_layout.clear_widgets()
        self.msg_label.text = ""
        self.is_valid_plata = True
        self.update_preview()

    def set_prefilled_data(self, zavod_or_name):
        self.zavod_input.text = zavod_or_name.upper()
        self.update_preview()

    def start_save_action(self):
        q = self.chosen_qurilma
        p = self.plata_input.text.strip()
        z = self.zavod_input.text.strip()
        j = self.manual_joy_input.text.strip() if self.manual_joy_input.opacity == 1 else self.chosen_joy

        if not q or not p or not z or not j:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Barcha maydonlarni to'ldiring va joyni tanlang!"
            return

        if not self.is_valid_plata:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Plata nomi noto'g'ri kiritilgan!"
            return

        self.spinner.opacity = 1
        self.msg_label.color = (0.0, 0.45, 0.45, 1)
        self.msg_label.text = "Bazaga yuklanmoqda..."
        Clock.schedule_once(lambda dt: self.save_data(), 0.3)

    def save_data(self):
        q = self.chosen_qurilma
        p = self.plata_input.text.strip().upper()
        z = self.zavod_input.text.strip().upper()
        i = self.inv_input.text.strip()
        h = self.chosen_holat
        j = self.manual_joy_input.text.strip() if self.manual_joy_input.opacity == 1 else self.chosen_joy

        new_row = {
            'Qurilma': q,
            'Plata': p,
            'Zavod raqami': z,
            'Inv raqami': i,
            "Joylashgan o'rni": j,
            "Tekshirilgan": "",
            "Holati": h
        }

        if save_to_excel_by_category(new_row):
            self.msg_label.color = (0, 0.6, 0.3, 1)
            self.msg_label.text = "Muvaffaqiyatli qo'shildi va saqlandi!"
            self.zavod_input.text = ""
            self.inv_input.text = ""
            self.plata_input.text = ""
            self.manual_joy_input.text = ""
            self.chosen_joy = ""
            self.chosen_holat = "Soz"
            self.selected_qurilma_lbl.text = "Tanlangan qurilma: (Tanlanmagan)"
            self.selected_joy_lbl.text = "Tanlangan joy: (Tanlanmagan)"
            self.selected_holat_lbl.text = "Tanlangan holat: Soz"
            self.update_preview()
        else:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Xatolik yuz berdi!"
        self.spinner.opacity = 0

# --- STANSIYAGA O'RNATISH EKRANI ---
class MovePlataScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.selected_row = None
        self.selected_zavod = ""
        self.chosen_yangi_joy = ""
        self.selected_plata_lbl.text = "Tanlangan plata: (Tanlanmagan)"
        self.zavod_input.text = ""
        self.manual_yangi_joy.opacity = 0
        self.manual_yangi_joy.disabled = True
        self.manual_yangi_joy.text = ""
        self.sug_layout.clear_widgets()
        self.msg_topilmadi.text = ""
        self.msg_label.text = ""
        self.spinner.opacity = 0

    def __init__(self, **kwargs):
        super(MovePlataScreen, self).__init__(**kwargs)
        
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=15, spacing=8, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))

        layout.add_widget(TopBar('Stansiyaga o\'rnatish', back_callback=lambda x: setattr(self.manager, 'current', 'platalar_menu')))

        layout.add_widget(AdaptiveLabel(text='Zavod raqami (oxirgi 4 ta raqami):', font_size=13, size_hint_y=None, height=22, color=(0.3, 0.3, 0.3, 1), halign='left'))
        
        zavod_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=40, spacing=5)
        self.zavod_input = TextInput(hint_text='Masalan: 1919', multiline=False, font_size=14)
        self.zavod_input.bind(text=self.filter_zavodlar)
        
        btn_clear_z = SoliqButton(text='❌', bg_color=(0.8, 0.3, 0.3, 1), size_hint_x=None, width=40, height=40)
        btn_clear_z.bind(on_press=self.clear_zavod_input)

        zavod_box.add_widget(self.zavod_input)
        zavod_box.add_widget(btn_clear_z)
        layout.add_widget(zavod_box)

        self.msg_topilmadi = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=20, color=(0.8, 0.3, 0.3, 1), halign='left')
        layout.add_widget(self.msg_topilmadi)

        self.sug_scroll = ScrollView(size_hint_y=None, height=100)
        self.sug_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=3)
        self.sug_layout.bind(minimum_height=self.sug_layout.setter('height'))
        self.sug_scroll.add_widget(self.sug_layout)
        layout.add_widget(self.sug_scroll)

        self.selected_plata_lbl = Label(
            text='Tanlangan plata: (Tanlanmagan)', 
            font_size=14, 
            size_hint_y=None, 
            height=145, 
            color=(0.15, 0.15, 0.15, 1), 
            halign='left',
            valign='middle',
            markup=True
        )
        if EMOJI_FONT:
            self.selected_plata_lbl.font_name = EMOJI_FONT
        self.selected_plata_lbl.bind(size=self.selected_plata_lbl.setter('text_size'))
        layout.add_widget(self.selected_plata_lbl)

        layout.add_widget(AdaptiveLabel(text='Qaysi stansiyaga o\'rnatishni tanlang:', font_size=13, size_hint_y=None, height=22, color=(0.3, 0.3, 0.3, 1), halign='left'))
        
        joy_grid = GridLayout(cols=2, size_hint_y=None, height=85, spacing=8)
        
        btn_kok = SoliqButton(text='KokOmbor', bg_color=(0.0, 0.45, 0.45, 1), height=38)
        btn_kok.bind(on_press=lambda x: self.select_yangi_joy('KokOmbor'))
        
        btn_far = SoliqButton(text='FarAMTS ombor', bg_color=(0.0, 0.5, 0.5, 1), height=38)
        btn_far.bind(on_press=lambda x: self.select_yangi_joy('FarAMTS ombor'))
        
        btn_nam = SoliqButton(text='NamAMTS ombor', bg_color=(0.2, 0.5, 0.6, 1), height=38)
        btn_nam.bind(on_press=lambda x: self.select_yangi_joy('NamAMTS ombor'))
        
        btn_and = SoliqButton(text='AndAMTS ombor', bg_color=(0.3, 0.6, 0.5, 1), height=38)
        btn_and.bind(on_press=lambda x: self.select_yangi_joy('AndAMTS ombor'))

        joy_grid.add_widget(btn_kok)
        joy_grid.add_widget(btn_far)
        joy_grid.add_widget(btn_nam)
        joy_grid.add_widget(btn_and)
        layout.add_widget(joy_grid)

        self.manual_yangi_joy = TextInput(hint_text='Boshqa stansiya nomini qo\'lda kiriting...', multiline=False, size_hint_y=None, height=0)
        self.manual_yangi_joy.bind(text=self.on_manual_joy_text)
        layout.add_widget(self.manual_yangi_joy)

        msg_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=25, spacing=10)
        self.msg_label = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=25, color=(0, 0.6, 0.3, 1), halign='left')
        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0
        msg_box.add_widget(self.msg_label)
        msg_box.add_widget(self.spinner)
        layout.add_widget(msg_box)

        btn_update = SoliqButton(text='✏️ Boshqa stansiya (joy)', bg_color=(0.6, 0.4, 0.4, 1), size_hint_y=None, height=45)
        btn_update.bind(on_press=lambda x: self.start_update_action())
        layout.add_widget(btn_update)

        scroll.add_widget(layout)
        self.add_widget(scroll)

    def clear_zavod_input(self, instance):
        self.zavod_input.text = ""
        self.sug_layout.clear_widgets()
        self.selected_zavod = ""
        self.selected_row = None
        self.selected_plata_lbl.text = "Tanlangan plata: (Tanlanmagan)"
        self.msg_topilmadi.text = ""

    def filter_zavodlar(self, instance, value):
        self.sug_layout.clear_widgets()
        self.msg_topilmadi.text = ""
        search_val = value.strip().lower()
        if not search_val:
            return
        df = load_excel_database()
        if df is not None and 'Zavod raqami' in df.columns:
            matched = df[df['Zavod raqami'].astype(str).str.lower().str.endswith(search_val)]
            
            if not matched.empty:
                for idx, row in matched.head(10).iterrows():
                    plata = str(row.get('Plata', '')).upper()
                    zavod = str(row.get('Zavod raqami', '')).upper()
                    
                    zavod_colored = zavod
                    if search_val in zavod.lower():
                        search_upper = search_val.upper()
                        zavod_colored = zavod.replace(search_upper, f"[color=004545]{search_upper}[/color]")
                    
                    btn_text = f"{plata} | Zavod №: {zavod_colored}"
                    btn = SoliqButton(text=btn_text, bg_color=(0.3, 0.5, 0.5, 1), size_hint_y=None, height=38)
                    btn.markup = True
                    btn.bind(on_press=lambda x, z_num=zavod, r=row: self.select_zavod(z_num, r))
                    self.sug_layout.add_widget(btn)
            else:
                self.msg_topilmadi.text = "Bazada topilmadi"

    def select_zavod(self, zavod_num, row):
        self.selected_zavod = zavod_num.upper()
        self.selected_row = row
        self.chosen_yangi_joy = ""
        self.manual_yangi_joy.height = 0
        self.manual_yangi_joy.opacity = 0
        self.manual_yangi_joy.disabled = True
        self.manual_yangi_joy.text = ""
        self.update_info_display()
        self.sug_layout.clear_widgets()
        self.msg_topilmadi.text = ""

    def select_yangi_joy(self, val):
        self.chosen_yangi_joy = val
        self.manual_yangi_joy.height = 0
        self.manual_yangi_joy.opacity = 0
        self.manual_yangi_joy.disabled = True
        self.manual_yangi_joy.text = ""
        self.update_info_display()

    def on_manual_joy_text(self, instance, value):
        self.chosen_yangi_joy = value.strip()
        self.update_info_display()

    def update_info_display(self):
        if self.selected_row is not None:
            row = self.selected_row
            plata = str(row.get('Plata', '')).upper()
            zavod = str(row.get('Zavod raqami', '')).upper()
            inv = str(row.get('Inv raqami', ''))
            hozirgi_joy = str(row.get("Joylashgan o'rni", ''))
            yangi_j = f"[color=008800]{self.chosen_yangi_joy}[/color]" if self.chosen_yangi_joy else "—"

            info_str = (
                f"[b]Tanlangan plata ma'lumoti:[/b]\n"
                f"• Plata nomi: [color=004545]{plata}[/color]\n"
                f"• Zavod №: [color=004545]{zavod}[/color]\n"
                f"• Inv №: {inv}\n"
                f"• Hozirgi joyi: {hozirgi_joy}\n"
                f"• Yangi stansiya: {yangi_j}"
            )
            self.selected_plata_lbl.text = info_str

    def start_update_action(self):
        if not self.selected_zavod:
            self.msg_label.color = (0.8, 0.2, 0.2, 1)
            self.msg_label.text = "Iltimos, avval ro'yxatdan platani tanlang!"
            return

        yangi_joy = self.manual_yangi_joy.text.strip() if self.manual_yangi_joy.opacity == 1 else self.chosen_yangi_joy
        if not yangi_joy:
            if self.manual_yangi_joy.opacity == 0:
                self.manual_yangi_joy.height = 38
                self.manual_yangi_joy.opacity = 1
                self.manual_yangi_joy.disabled = False
                self.msg_label.color = (0.8, 0.5, 0.1, 1)
                self.msg_label.text = "QO'LDA STANSIYA NOMINI KIRITING VA YANA BOSING!"
                return
            else:
                self.msg_label.color = (0.8, 0.2, 0.2, 1)
                self.msg_label.text = "ILTIMOS, STANSIYA NOMINI YAZING!"
                return

        self.spinner.opacity = 1
        self.msg_label.color = (0.0, 0.45, 0.45, 1)
        self.msg_label.text = "Stansiyaga o'rnatilmoqda..."
        Clock.schedule_once(lambda dt: self.update_location(), 0.3)

    def update_location(self):
        yangi_joy = self.manual_yangi_joy.text.strip() if self.manual_yangi_joy.opacity == 1 else self.chosen_yangi_joy
        df = load_excel_database()
        if df is not None:
            mask = df['Zavod raqami'].astype(str).str.strip() == str(self.selected_zavod).strip()
            if mask.any():
                df.loc[mask, "Joylashgan o'rni"] = yangi_joy
                df.to_excel(EXCEL_PATH, index=False)
                self.msg_label.color = (0, 0.6, 0.3, 1)
                self.msg_label.text = f"Muvaffaqiyatli! Zavod № [{self.selected_zavod}] stansiyaga o'rnatildi."
                
                row = df[mask].iloc[0]
                plata = str(row.get('Plata', '')).upper()
                zavod = str(row.get('Zavod raqami', '')).upper()
                inv = str(row.get('Inv raqami', ''))
                
                info_str = (
                    f"[b]O'rnartildi (Saqlandi ✅):[/b]\n"
                    f"• Plata nomi: {plata}\n"
                    f"• Zavod №: {zavod}\n"
                    f"• Inv №: {inv}\n"
                    f"• Yangi stansiya: [color=008800]{yangi_joy}[/color]"
                )
                self.selected_plata_lbl.text = info_str

                self.zavod_input.text = ""
                self.manual_yangi_joy.text = ""
                self.selected_zavod = ""
                self.chosen_yangi_joy = ""
                self.selected_row = None
            else:
                self.msg_label.color = (0.8, 0.2, 0.2, 1)
                self.msg_label.text = "Bazada topilmadi"
        self.spinner.opacity = 0

# --- PLATANI TAHRIRLASH EKRANI ---
class EditPlataScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.spinner.opacity = 0
        self.msg_label.text = ""

    def set_data(self, row_data):
        self.row_data = row_data
        self.zavod_original = str(row_data.get('Zavod raqami', '')).upper()
        
        self.qurilma_input.text = str(row_data.get('Qurilma', ''))
        self.plata_input.text = str(row_data.get('Plata', '')).upper()
        self.zavod_input.text = str(row_data.get('Zavod raqami', '')).upper()
        self.inv_input.text = str(row_data.get('Inv raqami', ''))
        self.holat_input.text = str(row_data.get('Holati', 'Soz'))
        self.joy_input.text = str(row_data.get("Joylashgan o'rni", ''))

    def __init__(self, **kwargs):
        super(EditPlataScreen, self).__init__(**kwargs)
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=20, spacing=12, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))

        layout.add_widget(TopBar('Platani tahrirlash', back_callback=lambda x: setattr(self.manager, 'current', 'inv_screen')))

        layout.add_widget(AdaptiveLabel(text='Qurilma nomi:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.qurilma_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.qurilma_input)

        layout.add_widget(AdaptiveLabel(text='Plata nomi:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.plata_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.plata_input)

        layout.add_widget(AdaptiveLabel(text='Zavod raqami:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.zavod_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.zavod_input)

        layout.add_widget(AdaptiveLabel(text='Inventar raqami:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.inv_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.inv_input)

        layout.add_widget(AdaptiveLabel(text='Holati (Soz / Nosoz):', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.holat_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.holat_input)

        layout.add_widget(AdaptiveLabel(text='Saqlanish joyi:', font_size=14, size_hint_y=None, height=25, color=(0.3, 0.3, 0.3, 1), halign='left'))
        self.joy_input = TextInput(multiline=False, size_hint_y=None, height=40)
        layout.add_widget(self.joy_input)

        msg_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=30, spacing=10)
        self.msg_label = AdaptiveLabel(text='', font_size=14, size_hint_y=None, height=30, color=(0, 0.6, 0.3, 1), halign='left')
        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0
        msg_box.add_widget(self.msg_label)
        msg_box.add_widget(self.spinner)
        layout.add_widget(msg_box)

        btn_save = SoliqButton(text='💾 O\'zgarishlarni saqlash', bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=50)
        btn_save.bind(on_press=lambda x: self.start_save())
        layout.add_widget(btn_save)

        scroll.add_widget(layout)
        self.add_widget(scroll)

    def start_save(self):
        self.spinner.opacity = 1
        self.msg_label.color = (0.0, 0.45, 0.45, 1)
        self.msg_label.text = "Saqlanmoqda..."
        Clock.schedule_once(lambda dt: self.save_changes(), 0.3)

    def save_changes(self):
        df = load_excel_database()
        if df is not None:
            mask = df['Zavod raqami'].astype(str).str.strip() == str(self.zavod_original).strip()
            if mask.any():
                df.loc[mask, 'Qurilma'] = self.qurilma_input.text.strip()
                df.loc[mask, 'Plata'] = self.plata_input.text.strip().upper()
                df.loc[mask, 'Zavod raqami'] = self.zavod_input.text.strip().upper()
                df.loc[mask, 'Inv raqami'] = self.inv_input.text.strip()
                df.loc[mask, 'Holati'] = self.holat_input.text.strip()
                df.loc[mask, "Joylashgan o'rni"] = self.joy_input.text.strip()
                
                df.to_excel(EXCEL_PATH, index=False)
                self.msg_label.color = (0, 0.6, 0.3, 1)
                self.msg_label.text = "Muvaffaqiyatli saqlandi!"
                Clock.schedule_once(lambda dt: setattr(self.manager, 'current', 'inv_screen'), 1.0)
            else:
                self.msg_label.color = (0.8, 0.2, 0.2, 1)
                self.msg_label.text = "Bazada topilmadi!"
        self.spinner.opacity = 0

# --- 3. INVENTARIZATSIYA EKRANI ---
class InventoryScreen(Screen):
    def on_enter(self):
        Window.clearcolor = (0.95, 0.96, 0.98, 1)
        self.load_qurilmalar_buttons()
        self.result_layout.clear_widgets()
        self.msg_topilmadi.text = ""

    def __init__(self, **kwargs):
        super(InventoryScreen, self).__init__(**kwargs)
        self.selected_qurilmalar = set()
        self.qur_buttons = []

        layout = BoxLayout(orientation='vertical', padding=15, spacing=10)
        
        layout.add_widget(TopBar('Invertarizatsiya', back_callback=lambda x: setattr(self.manager, 'current', 'platalar_menu')))
        
        # Qidiruv paneli
        top_search_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=42, spacing=8)
        self.query_input = TextInput(hint_text='Plata nomi yoki zavod raqami...', multiline=False, font_size=13)
        self.query_input.bind(text=self.trigger_search)
        
        btn_clear_inv = SoliqButton(text='❌', bg_color=(0.8, 0.3, 0.3, 1), size_hint_x=None, width=40, height=42)
        btn_clear_inv.bind(on_press=self.clear_search)

        self.spinner = LoadingSpinner()
        self.spinner.opacity = 0

        top_search_layout.add_widget(self.query_input)
        top_search_layout.add_widget(self.spinner)
        top_search_layout.add_widget(btn_clear_inv)
        layout.add_widget(top_search_layout)

        # QURILMALARNI SELEKT QILISH PANELI
        layout.add_widget(AdaptiveLabel(text="Qurilmani tanlang:", font_size=13, bold=True, size_hint_y=None, height=20, color=(0.2, 0.2, 0.2, 1), halign='left'))

        self.qur_scroll = ScrollView(size_hint_y=None, height=110)
        self.qur_grid = GridLayout(cols=2, size_hint_y=None, spacing=6)
        self.qur_grid.bind(minimum_height=self.qur_grid.setter('height'))
        self.qur_scroll.add_widget(self.qur_grid)
        layout.add_widget(self.qur_scroll)

        self.msg_topilmadi = AdaptiveLabel(text='', font_size=13, size_hint_y=None, height=22, color=(0.8, 0.3, 0.3, 1), halign='left')
        layout.add_widget(self.msg_topilmadi)

        self.result_scroll = ScrollView()
        self.result_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=10)
        self.result_layout.bind(minimum_height=self.result_layout.setter('height'))
        self.result_scroll.add_widget(self.result_layout)
        layout.add_widget(self.result_scroll)

        self.add_widget(layout)

    def load_qurilmalar_buttons(self):
        """Bazada bor barcha qurilmalarga interaktiv Tugmalar yasash"""
        self.qur_grid.clear_widgets()
        self.selected_qurilmalar.clear()
        self.qur_buttons.clear()

        df = load_excel_database()
        if df is not None and 'Qurilma' in df.columns:
            df['Normalized_Qurilma'] = df['Qurilma'].apply(normalize_qurilma_name)
            raw_qurilmalar = df['Normalized_Qurilma'].dropna().unique()
            qurilmalar = sorted([str(q) for q in raw_qurilmalar if q is not None])

            for q in qurilmalar:
                btn = QurilmaSelectButton(q_name=q)
                btn.bind(on_press=lambda instance, name=q: self.toggle_qurilma_selection(instance, name))
                self.qur_grid.add_widget(btn)
                self.qur_buttons.append(btn)

    def toggle_qurilma_selection(self, instance, q_name):
        """Tugma bosilganda o'sha qurilmaga ko'ra platalarni avtomatik saralab beradi"""
        self.query_input.text = ""
        if instance.state == 'down':
            self.selected_qurilmalar.add(q_name)
        else:
            self.selected_qurilmalar.discard(q_name)

        self.filter_by_selected_qurilmalar()

    def filter_by_selected_qurilmalar(self):
        """Tanlangan qurilmalar bo'yicha platalarni ko'rsatish"""
        self.spinner.opacity = 1
        self.result_layout.clear_widgets()
        self.msg_topilmadi.text = ""

        if not self.selected_qurilmalar:
            self.spinner.opacity = 0
            return

        df = load_excel_database()
        if df is not None:
            df['Normalized_Qurilma'] = df['Qurilma'].apply(normalize_qurilma_name)
            filtrlangan = df[df['Normalized_Qurilma'].astype(str).str.strip().isin(self.selected_qurilmalar)]

            if not filtrlangan.empty:
                self.result_layout.add_widget(AdaptiveLabel(text=f"Tanlangan qurilmalar platalari (Jami: {len(filtrlangan)} ta):", font_size=14, bold=True, color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=30, halign='left'))
                self.render_items_list(filtrlangan)
            else:
                self.msg_topilmadi.text = "Tanlangan qurilmalarda platalar topilmadi!"
        self.spinner.opacity = 0

    def clear_search(self, instance):
        self.query_input.text = ""
        for btn in self.qur_buttons:
            btn.state = 'normal'
            btn.update_text()
            btn.update_canvas()
        self.selected_qurilmalar.clear()
        self.result_layout.clear_widgets()
        self.msg_topilmadi.text = ""

    def trigger_search(self, instance, value):
        val = value.strip().lower()
        if not val:
            if self.selected_qurilmalar:
                self.filter_by_selected_qurilmalar()
            else:
                self.result_layout.clear_widgets()
                self.msg_topilmadi.text = ""
            return

        self.spinner.opacity = 1
        self.result_layout.clear_widgets()
        self.msg_topilmadi.text = ""
        Clock.unschedule(self.perform_search)
        Clock.schedule_once(lambda dt: self.perform_search(val), 0.25)

    def perform_search(self, val):
        df = load_excel_database()
        if df is not None:
            if any(char.isalpha() for char in val):
                matched = df[df['Plata'].astype(str).str.lower().str.contains(val)]
            else:
                matched = df[df['Zavod raqami'].astype(str).str.lower().str.endswith(val)]

            if not matched.empty:
                self.result_layout.add_widget(AdaptiveLabel(text=f"Topildi! Jami mos keluvchi: {len(matched)} ta", font_size=15, bold=True, color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=35, halign='left'))
                self.render_items_list(matched, val)
            else:
                self.msg_topilmadi.text = "Bazada topilmadi"
                btn_add = SoliqButton(text="➕ Uni bazaga qo'shish", bg_color=(0.0, 0.45, 0.45, 1), size_hint_y=None, height=50)
                btn_add.bind(on_press=lambda x: self.go_to_add(val))
                self.result_layout.add_widget(btn_add)
        
        self.spinner.opacity = 0

    def render_items_list(self, df_data, val=""):
        """Elementlarni ketma-ket ekranga chiqarish"""
        for idx, row in df_data.iterrows():
            qurilma = str(row.get('Qurilma', ''))
            plata = str(row.get('Plata', '')).upper()
            zavod = str(row.get('Zavod raqami', '')).upper()
            inv = str(row.get('Inv raqami', ''))
            joy = str(row.get("Joylashgan o'rni", ''))
            holat = str(row.get('Holati', 'Soz'))
            tekshirildi = str(row.get("Tekshirilgan", ''))

            p_disp = plata
            z_disp = zavod
            i_disp = inv
            if val:
                if val in plata.lower():
                    p_disp = plata.upper().replace(val.upper(), f"[color=004545]{val.upper()}[/color]")
                if val in zavod.lower():
                    z_disp = zavod.upper().replace(val.upper(), f"[color=004545]{val.upper()}[/color]")
                if val in inv.lower():
                    i_disp = inv.replace(val, f"[color=004545]{val}[/color]")

            is_checked = (tekshirildi == "✅ Tekshirilgan")
            status_text = "[color=008800]✅ Tekshirilgan[/color]" if is_checked else "[color=aa8800]❌ Tekshirilmagan[/color]"
            
            info_text = (
                f"• Plata nomi: {p_disp} (Qurilma: {qurilma})\n"
                f"• Zavod raqami: {z_disp} | Inv raqami: {i_disp}\n"
                f"• Holati: {holat} | Saqlanish joyi: {joy}\n"
                f"• Tekshirildi: {status_text}"
            )
            
            box_height = 130 if is_checked else 170
            box = BoxLayout(orientation='vertical', size_hint_y=None, height=box_height, spacing=5)
            
            lbl = Label(
                text=info_text, 
                font_size=14, 
                size_hint_y=None, 
                height=95, 
                color=(0.2, 0.2, 0.2, 1),
                halign='left',
                valign='middle',
                markup=True
            )
            if EMOJI_FONT:
                lbl.font_name = EMOJI_FONT
            lbl.bind(size=lbl.setter('text_size'))
            box.add_widget(lbl)

            if not is_checked:
                btn_mark = SoliqButton(text="✅ Tekshirilda deb belgilash", bg_color=(0.0, 0.5, 0.4, 1), size_hint_y=None, height=35)
                btn_mark.bind(on_press=lambda x, z_num=zavod: self.mark_as_checked(z_num))
                box.add_widget(btn_mark)

            btn_row = BoxLayout(orientation='horizontal', size_hint_y=None, height=35, spacing=8)
            btn_edit = SoliqButton(text="✏️ Tahrirlash", bg_color=(0.2, 0.4, 0.6, 1))
            btn_edit.bind(on_press=lambda x, r_data=row: self.go_to_edit(r_data))
            
            btn_del = SoliqButton(text="🗑 O'chirish", bg_color=(0.8, 0.3, 0.3, 1))
            btn_del.bind(on_press=lambda x, z_num=zavod: self.delete_plata(z_num))

            btn_row.add_widget(btn_edit)
            btn_row.add_widget(btn_del)
            box.add_widget(btn_row)

            self.result_layout.add_widget(box)
            
            sep = AdaptiveLabel(text="----------------------------------------", font_size=14, size_hint_y=None, height=20, color=(0.7, 0.7, 0.7, 1), halign='left')
            self.result_layout.add_widget(sep)

    def mark_as_checked(self, zavod_num):
        df = load_excel_database()
        if df is not None:
            mask = df['Zavod raqami'].astype(str).str.strip() == str(zavod_num).strip()
            if mask.any():
                df.loc[mask, "Tekshirilgan"] = "✅ Tekshirilgan"
                df.to_excel(EXCEL_PATH, index=False)
                self.msg_topilmadi.color = (0, 0.6, 0.3, 1)
                self.msg_topilmadi.text = f"Zavod № [{zavod_num}] tekshirildi deb belgilandi!"
                
                if self.query_input.text.strip():
                    self.perform_search(self.query_input.text.strip().lower())
                elif self.selected_qurilmalar:
                    self.filter_by_selected_qurilmalar()

    def delete_plata(self, zavod_num):
        df = load_excel_database()
        if df is not None:
            df = df[df['Zavod raqami'].astype(str).str.strip() != str(zavod_num).strip()]
            df.to_excel(EXCEL_PATH, index=False)
            self.msg_topilmadi.color = (0.8, 0.3, 0.3, 1)
            self.msg_topilmadi.text = f"Zavod № [{zavod_num}] bazadan o'chirildi!"
            
            if self.query_input.text.strip():
                self.perform_search(self.query_input.text.strip().lower())
            elif self.selected_qurilmalar:
                self.filter_by_selected_qurilmalar()

    def go_to_edit(self, row_data):
        edit_screen = self.manager.get_screen('edit_screen')
        edit_screen.set_data(row_data)
        self.manager.current = 'edit_screen'

    def go_to_add(self, query_val):
        add_screen = self.manager.get_screen('add_screen')
        add_screen.set_prefilled_data(query_val)
        add_screen.plata_input.text = query_val.upper()
        self.manager.current = 'add_screen'

class OmborApp(App):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(MainMenuScreen(name='main_menu'))
        sm.add_widget(ExcelMenuScreen(name='excel_menu'))
        sm.add_widget(ExcelPlatalarMenuScreen(name='excel_platalar_menu'))
        sm.add_widget(PlatalarMenuScreen(name='platalar_menu'))
        sm.add_widget(SFPMenuScreen(name='sfp_menu'))
        sm.add_widget(QurilmalarMenuScreen(name='qurilmalar_menu'))
        
        sm.add_widget(PlatalarScreen(name='platalar_screen'))
        sm.add_widget(PlatalarDetalScreen(name='detal_screen'))
        sm.add_widget(AddPlataScreen(name='add_screen'))
        sm.add_widget(MovePlataScreen(name='move_screen'))
        sm.add_widget(InventoryScreen(name='inv_screen'))
        sm.add_widget(EditPlataScreen(name='edit_screen'))
        return sm

if __name__ == '__main__':
    OmborApp().run()
