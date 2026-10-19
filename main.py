# -*- coding: utf-8 -*-
import math
import os
import sys
from datetime import datetime, timedelta
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDFlatButton, MDRaisedButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.label import MDLabel
from kivymd.uix.screen import MDScreen
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.segmentedcontrol import (
    MDSegmentedControl,
    MDSegmentedControlItem,
)
from kivymd.uix.textfield import MDTextField
from kivymd.uix.toolbar import MDTopAppBar
from kivy.uix.image import Image


class GasCalculatorApp(MDApp):

    EQUIPMENT_DATA = {
        "stove_gvp": {
            "name": "Плита газова за наявності ЦГВ",
            "unit": "м3 / люд.-місяць",
            "rate": 9.8,
            "type": "people",
        },
        "stove_no_gvp": {
            "name": "Плита газова без ЦГВ та газового водонагрівача",
            "unit": "м3 / люд.-місяць",
            "rate": 18.3,
            "type": "people",
        },
        "stove_and_heater": {
            "name": "Плита газова та водонагрівач",
            "unit": "м3 / люд.-місяць",
            "rate": 23.6,
            "type": "people",
        },
        "heating_stove": {
            "name": "Опалювально-варильна піч (кухонне вогнище)",
            "unit": "Згідно Додатка 15 Кодексу ГРМ",
            "type": "hybrid_stove",
        },
        "individual_heating": {
            "name": "Індивідуальне опалення (крім печей)",
            "unit": "м3 / м2 на місяць",
            "rate_heating": 11.0,
            "type": "area",
        },
        "hole_diameter_calc": {
            "name": "Розрахувати за формулою діаметру отвору",
            "unit": "м3 (за формулою отвору)",
            "type": "hole_diameter",
        },
    }

    def build(self):
        self.theme_cls.primary_palette = "Teal"

        # --- Змінні стану ---
        self.consumer_type = "household"
        self.account_number = ""
        self.eic_code = ""
        self.selected_violation_code = None
        self.selected_violation_text = None
        self.is_hidden_violation = False
        self.dialog = None

        # Дати для побутового споживача
        self.date_from = None
        self.date_to = None

        # Дати та час для промислового споживача
        self.ind_datetime_from = None
        self.ind_datetime_to = None

        self.total_days = 0
        self.total_hours = 0.0
        self.heating_days = 0
        self.heating_hours = 0.0
        self.offseason_days = 0
        self.offseason_hours = 0.0

        self.registered_persons = 0
        self.heated_area = 0.0
        self.hole_diameter = 0.0
        self.gas_pressure = 0.0
        self.specific_volume = 0.0
        self.selected_equipment = set()

        # Змінні для промислового обладнання та методу витоку
        self.industrial_calculation_mode = "equipment"  # 'equipment' або 'hole'
        self.industrial_equipment_count = 0
        self.industrial_equipment_items = []

        self.last_total_volume = 0.0
        self.last_details_text = ""

        # Словник порушень для побутового споживача
        self.household_violations_dict = {
            "п. 1.1": "1.1 Наявність несанкціонованого газопроводу",
            "п. 1.2": "1.2 Несанкціоноване відновлення газоспоживання",
            "п. 1.3": "1.3 Несанкціоноване втручання в роботу ЗВТ",
            "п. 1.4": "1.4 Несанкціоноване підключення приладів (облік за нормами)",
            "п. 1.5": "1.5 Несанкціоноване підключення (перевищення діапазону ЗВТ - 70% від норми)",
            "п. 1.6": "1.6 Використання газу без/після розірвання договору",
        }

        # Словник порушень для промислового споживача
        self.industrial_violations_dict = {
            "п. 1.1": "1.1 Наявність несанкціонованого газопроводу",
            "п. 1.2": "1.2 Несанкціоноване відновлення газоспоживання",
            "п. 1.3": "1.3 Несанкціоноване втручання в роботу ЗВТ",
            "п. 1.4": "1.4 Несанкціоноване підключення приладів",
            "п. 1.5": "1.5 Перевищення діапазону обчислення / вимірювання ЗВТ",
            "п. 1.6": "1.6 Використання газу без/після розірвання договору",
            "Розділ XI. Глава 2. п. 3.1": "п. 3.1 Пошкодження ЗВТ (лічильника газу) або робота комерційного ВОГ чи його складових в позаштатному режимі...",
            "Розділ XI. Глава 2. п. 3.2": "п. 3.2 Фіксація коректором чи обчислювачем об'єму газу даних щодо збоїв в роботі комерційного ВОГ чи його складових...",
            "Розділ XI. Глава 2. п. 3.3": "п. 3.3 Виявлення помилок в конфігурації даних обчислювача об'єму газу, що були введені при його первинному програмуванні...",
            "Розділ XI. Глава 2. п. 3.4": "п. 3.4 Перевищення діапазону обчислення вузла обліку тощо",
        }

        screen = MDScreen()
        scroll_main = MDScrollView()

        self.main_layout = MDBoxLayout(
            orientation="vertical",
            spacing="12dp",
            padding=["16dp", "0dp", "16dp", "20dp"],
            size_hint_y=None,
        )
        self.main_layout.bind(minimum_height=self.main_layout.setter("height"))

        # Визначення шляху до логотипу (з урахуванням PyInstaller)
        if hasattr(sys, '_MEIPASS'):
            logo_path = os.path.join(sys._MEIPASS, "logo.png")
        else:
            logo_path = "logo.png"

        # Верхня панель зі збільшеним заголовком та логотипом
        toolbar = MDTopAppBar(
            title="",
            pos_hint={"top": 1},
            elevation=2,
            height="64dp",
        )
        
        logo_box = MDBoxLayout(
            orientation="horizontal",
            spacing="12dp",
            size_hint_x=0.95,
            pos_hint={"center_y": 0.5},
        )
        
        if os.path.exists(logo_path):
            logo_img = Image(
                source=logo_path,
                size_hint=(None, None),
                size=("44dp", "44dp"),
                allow_stretch=True,
                keep_ratio=True
            )
            logo_box.add_widget(logo_img)

        # Збільшений шрифт назви (40dp / H6)
        title_label = MDLabel(
            text="Газорозподільні мережі України",
            font_style="H6",
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            size_hint_y=None,
            height="40dp",
            pos_hint={"center_y": 0.5}
        )
        logo_box.add_widget(title_label)
        toolbar.add_widget(logo_box)

        step1_label = MDLabel(
            text="Крок 1: Оберіть категорію споживача",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )

        self.consumer_segment = MDSegmentedControl(
            size_hint_x=1, pos_hint={"center_x": 0.5}
        )
        item_household = MDSegmentedControlItem(text="Побутовий (Населення)")
        item_industrial = MDSegmentedControlItem(text="Промисловий (Непобутовий)")
        self.consumer_segment.add_widget(item_household)
        self.consumer_segment.add_widget(item_industrial)
        self.consumer_segment.bind(on_active=self.on_consumer_type_change)

        self.forms_container = MDBoxLayout(orientation="vertical", size_hint_y=None)
        self.forms_container.bind(minimum_height=self.forms_container.setter("height"))

        self.setup_household_form()
        self.setup_industrial_form()

        self.forms_container.add_widget(self.household_form)

        self.main_layout.add_widget(toolbar)
        self.main_layout.add_widget(step1_label)
        self.main_layout.add_widget(self.consumer_segment)
        self.main_layout.add_widget(self.forms_container)

        scroll_main.add_widget(self.main_layout)
        screen.add_widget(scroll_main)
        return screen

    # ==================== ФОРМА ПОБУТОВОГО СПОЖИВАЧА ====================
    def setup_household_form(self):
        self.household_form = MDBoxLayout(
            orientation="vertical", spacing="10dp", size_hint_y=None
        )
        self.household_form.bind(minimum_height=self.household_form.setter("height"))

        step1_1_label = MDLabel(
            text="Ідентифікатор споживача",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.household_form.add_widget(step1_1_label)

        self.account_field = MDTextField(
            hint_text="Особовий рахунок",
            mode="rectangle",
            size_hint_y=None,
            height="50dp",
        )
        self.account_field.bind(text=self.on_account_change)
        self.household_form.add_widget(self.account_field)

        step2_label = MDLabel(
            text="Крок 2: Оберіть вид порушення Кодексу ГРМ",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.household_form.add_widget(step2_label)

        self.select_violation_btn = MDRaisedButton(
            text="ОБРАТИ ВИД ПОРУШЕННЯ",
            pos_hint={"center_x": 0.5},
            size_hint_x=1,
            height="48dp",
        )
        self.select_violation_btn.bind(on_release=self.show_violation_dialog)
        self.household_form.add_widget(self.select_violation_btn)

        self.violation_field = MDTextField(
            hint_text="Обраний вид порушення",
            text="",
            readonly=True,
            mode="rectangle",
            size_hint_y=None,
            height="50dp",
        )
        self.household_form.add_widget(self.violation_field)

        self.hidden_option_layout = MDBoxLayout(
            orientation="horizontal",
            spacing="10dp",
            size_hint_y=None,
            height="40dp",
            opacity=0,
        )

        hidden_label = MDLabel(text="Ознака прихованості:", font_style="Body1", size_hint_x=0.7)
        self.hidden_checkbox = MDCheckbox(
            size_hint=(None, None), size=("36dp", "36dp"), pos_hint={"center_y": 0.5}
        )
        self.hidden_checkbox.bind(active=self.on_hidden_checkbox_change)

        self.hidden_checkbox_label = MDLabel(
            text="Неприховане",
            font_style="Body2",
            theme_text_color="Secondary",
            size_hint_x=0.3,
        )

        self.hidden_option_layout.add_widget(hidden_label)
        self.hidden_option_layout.add_widget(self.hidden_checkbox)
        self.hidden_option_layout.add_widget(self.hidden_checkbox_label)
        self.household_form.add_widget(self.hidden_option_layout)

        step3_label = MDLabel(
            text="Крок 3: Період нарахування (Введіть дати)",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.household_form.add_widget(step3_label)

        dates_layout = MDBoxLayout(
            orientation="horizontal", spacing="10dp", size_hint_y=None, height="55dp"
        )

        self.date_from_field = MDTextField(
            hint_text="З (ДД.ММ.РРРР)", mode="rectangle", size_hint_x=0.5
        )
        self.date_from_field.bind(text=self.on_household_date_from_change)

        self.date_to_field = MDTextField(
            hint_text="По (ДД.ММ.РРРР)", mode="rectangle", size_hint_x=0.5
        )
        self.date_to_field.bind(text=self.on_household_date_to_change)

        dates_layout.add_widget(self.date_from_field)
        dates_layout.add_widget(self.date_to_field)
        self.household_form.add_widget(dates_layout)

        self.days_label = MDLabel(
            text="Всього днів: 0 (0 год.)",
            font_style="Subtitle2",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.period_details_label = MDLabel(
            text="• Опалювальний період: 0 днів (0 год.)\n• Неопалювальний період: 0 днів (0 год.)",
            font_style="Caption",
            theme_text_color="Secondary",
            size_hint_y=None,
            height="36dp",
        )
        self.household_form.add_widget(self.days_label)
        self.household_form.add_widget(self.period_details_label)

        step4_label = MDLabel(
            text="Крок 4: Характеристики об'єкта",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.household_form.add_widget(step4_label)

        details_layout = MDBoxLayout(
            orientation="horizontal", spacing="10dp", size_hint_y=None, height="55dp"
        )

        self.persons_field = MDTextField(
            hint_text="Зареєстровані особи", mode="rectangle", input_filter="int", size_hint_x=0.5
        )
        self.persons_field.bind(text=self.on_persons_change)

        self.area_field = MDTextField(
            hint_text="Опалювальна площа, м2", mode="rectangle", input_filter="float", size_hint_x=0.5
        )
        self.area_field.bind(text=self.on_area_change)

        details_layout.add_widget(self.persons_field)
        details_layout.add_widget(self.area_field)
        self.household_form.add_widget(details_layout)

        step5_label = MDLabel(
            text="Крок 5: Вибір газового обладнання / типу розрахунку",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.household_form.add_widget(step5_label)

        self.equipment_checkboxes_layout = MDBoxLayout(
            orientation="vertical", spacing="6dp", size_hint_y=None
        )
        self.equipment_checkboxes_layout.bind(
            minimum_height=self.equipment_checkboxes_layout.setter("height")
        )
        self.create_equipment_checkboxes()
        self.household_form.add_widget(self.equipment_checkboxes_layout)

        self.hole_formula_layout = MDBoxLayout(
            orientation="vertical", spacing="10dp", size_hint_y=None, opacity=0, disabled=True
        )
        self.hole_formula_layout.bind(minimum_height=self.hole_formula_layout.setter("height"))

        hole_title = MDLabel(
            text="Параметри для розрахунку за формулою отвору (k = 1.4):",
            font_style="Subtitle2",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.hole_formula_layout.add_widget(hole_title)

        self.hole_diameter_field = MDTextField(
            hint_text="Діаметр отвору газопроводу d, м", mode="rectangle", input_filter="float", size_hint_y=None, height="55dp"
        )
        self.hole_diameter_field.bind(text=lambda inst, val: self.on_float_param_change("hole_diameter", val))

        self.gas_pressure_field = MDTextField(
            hint_text="Тиск газу в газопроводі P, кг/м2", mode="rectangle", input_filter="float", size_hint_y=None, height="55dp"
        )
        self.gas_pressure_field.bind(text=lambda inst, val: self.on_float_param_change("gas_pressure", val))

        self.specific_volume_field = MDTextField(
            hint_text="Питомий обсяг газу v, м3/кг", mode="rectangle", input_filter="float", size_hint_y=None, height="55dp"
        )
        self.specific_volume_field.bind(text=lambda inst, val: self.on_float_param_change("specific_volume", val))

        self.hole_formula_layout.add_widget(self.hole_diameter_field)
        self.hole_formula_layout.add_widget(self.gas_pressure_field)
        self.hole_formula_layout.add_widget(self.specific_volume_field)

        self.household_form.add_widget(self.hole_formula_layout)

        # Панель кнопок (Розрахувати та Скинути)
        household_btn_layout = MDBoxLayout(
            orientation="horizontal", spacing="10dp", size_hint_y=None, height="50dp"
        )
        self.calculate_btn = MDRaisedButton(
            text="РОЗРАХУВАТИ",
            size_hint_x=0.6,
            height="50dp",
            md_bg_color=self.theme_cls.primary_color,
        )
        self.calculate_btn.bind(on_release=self.on_calculate_click)

        self.reset_household_btn = MDFlatButton(
            text="СКИНУТИ",
            size_hint_x=0.4,
            height="50dp",
            theme_text_color="Custom",
            text_color=self.theme_cls.primary_color,
        )
        self.reset_household_btn.bind(on_release=lambda x: self.reset_form("household"))

        household_btn_layout.add_widget(self.calculate_btn)
        household_btn_layout.add_widget(self.reset_household_btn)
        self.household_form.add_widget(household_btn_layout)

    # ==================== ФОРМА НЕПОБУТОВОГО СПОЖИВАЧА ====================
    def setup_industrial_form(self):
        self.industrial_form = MDBoxLayout(
            orientation="vertical", spacing="10dp", size_hint_y=None
        )
        self.industrial_form.bind(minimum_height=self.industrial_form.setter("height"))

        step2_ind_label = MDLabel(
            text="Крок 2: Вкажіть EIC-код точки обліку",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.industrial_form.add_widget(step2_ind_label)

        self.ind_eic_field = MDTextField(
            hint_text="EIC-код точки обліку (16 символів)",
            mode="rectangle",
            max_text_length=16,
            size_hint_y=None,
            height="50dp",
        )
        self.ind_eic_field.bind(text=lambda i, v: setattr(self, "eic_code", v.strip().upper()))
        self.industrial_form.add_widget(self.ind_eic_field)

        step3_ind_label = MDLabel(
            text="Крок 3: Оберіть вид порушення Кодексу ГРМ",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.industrial_form.add_widget(step3_ind_label)

        self.ind_select_violation_btn = MDRaisedButton(
            text="ОБРАТИ ВИД ПОРУШЕННЯ (ПРОМИСЛОВІСТЬ)",
            pos_hint={"center_x": 0.5},
            size_hint_x=1,
            height="48dp",
        )
        self.ind_select_violation_btn.bind(on_release=self.show_violation_dialog)
        self.industrial_form.add_widget(self.ind_select_violation_btn)

        self.ind_violation_field = MDTextField(
            hint_text="Обраний вид порушення",
            text="",
            readonly=True,
            mode="rectangle",
            size_hint_y=None,
            height="50dp",
        )
        self.industrial_form.add_widget(self.ind_violation_field)

        self.ind_hidden_option_layout = MDBoxLayout(
            orientation="horizontal",
            spacing="10dp",
            size_hint_y=None,
            height="40dp",
            opacity=0,
        )

        ind_hidden_label = MDLabel(text="Ознака прихованості:", font_style="Body1", size_hint_x=0.7)
        self.ind_hidden_checkbox = MDCheckbox(
            size_hint=(None, None), size=("36dp", "36dp"), pos_hint={"center_y": 0.5}
        )
        self.ind_hidden_checkbox.bind(active=self.on_hidden_checkbox_change)

        self.ind_hidden_checkbox_label = MDLabel(
            text="Неприховане",
            font_style="Body2",
            theme_text_color="Secondary",
            size_hint_x=0.3,
        )

        self.ind_hidden_option_layout.add_widget(ind_hidden_label)
        self.ind_hidden_option_layout.add_widget(self.ind_hidden_checkbox)
        self.ind_hidden_option_layout.add_widget(self.ind_hidden_checkbox_label)
        self.industrial_form.add_widget(self.ind_hidden_option_layout)

        step4_ind_label = MDLabel(
            text="Крок 4: Період нарахування (Введіть дату та час у 24-год форматі)",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.industrial_form.add_widget(step4_ind_label)

        ind_dates_layout = MDBoxLayout(
            orientation="horizontal", spacing="10dp", size_hint_y=None, height="55dp"
        )

        self.ind_datetime_from_field = MDTextField(
            hint_text="Початок (ДД.ММ.РРРР ГГ:ХХ)",
            mode="rectangle",
            size_hint_x=0.5,
        )
        self.ind_datetime_from_field.bind(text=self.on_ind_datetime_from_change)

        self.ind_datetime_to_field = MDTextField(
            hint_text="Кінець (ДД.ММ.РРРР ГГ:ХХ)",
            mode="rectangle",
            size_hint_x=0.5,
        )
        self.ind_datetime_to_field.bind(text=self.on_ind_datetime_to_change)

        ind_dates_layout.add_widget(self.ind_datetime_from_field)
        ind_dates_layout.add_widget(self.ind_datetime_to_field)
        self.industrial_form.add_widget(ind_dates_layout)

        self.ind_duration_label = MDLabel(
            text="Тривалість: 0 год. 0 хв. (0 діб)",
            font_style="Subtitle2",
            theme_text_color="Primary",
            size_hint_y=None,
            height="30dp",
        )
        self.industrial_form.add_widget(self.ind_duration_label)

        step5_ind_label = MDLabel(
            text="Крок 5: Метод визначення об'єму для промисловості",
            font_style="Subtitle1",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.industrial_form.add_widget(step5_ind_label)

        self.ind_mode_segment = MDSegmentedControl(
            size_hint_x=1, pos_hint={"center_x": 0.5}
        )
        item_mode_eq = MDSegmentedControlItem(text="За потужністю обладнання")
        item_mode_hole = MDSegmentedControlItem(text="За формулою отвору")
        self.ind_mode_segment.add_widget(item_mode_eq)
        self.ind_mode_segment.add_widget(item_mode_hole)
        self.ind_mode_segment.bind(on_active=self.on_industrial_mode_change)
        self.industrial_form.add_widget(self.ind_mode_segment)

        self.ind_equipment_mode_container = MDBoxLayout(
            orientation="vertical", spacing="8dp", size_hint_y=None
        )
        self.ind_equipment_mode_container.bind(
            minimum_height=self.ind_equipment_mode_container.setter("height")
        )

        self.ind_equipment_count_field = MDTextField(
            hint_text="Кількість одиниць обладнання",
            mode="rectangle",
            input_filter="int",
            size_hint_y=None,
            height="50dp",
        )
        self.ind_equipment_count_field.bind(text=self.on_industrial_equipment_count_change)
        self.ind_equipment_mode_container.add_widget(self.ind_equipment_count_field)

        self.ind_equipment_dynamic_layout = MDBoxLayout(
            orientation="vertical", spacing="8dp", size_hint_y=None
        )
        self.ind_equipment_dynamic_layout.bind(
            minimum_height=self.ind_equipment_dynamic_layout.setter("height")
        )
        self.ind_equipment_mode_container.add_widget(self.ind_equipment_dynamic_layout)

        self.ind_hole_mode_container = MDBoxLayout(
            orientation="vertical", spacing="10dp", size_hint_y=None, opacity=0, disabled=True
        )
        self.ind_hole_mode_container.bind(
            minimum_height=self.ind_hole_mode_container.setter("height")
        )

        ind_hole_title = MDLabel(
            text="Параметри витоку через отвір (k = 1.4):",
            font_style="Subtitle2",
            theme_text_color="Primary",
            size_hint_y=None,
            height="25dp",
        )
        self.ind_hole_mode_container.add_widget(ind_hole_title)

        self.ind_hole_diameter_field = MDTextField(
            hint_text="Діаметр отвору d, м", mode="rectangle", input_filter="float", size_hint_y=None, height="50dp"
        )
        self.ind_hole_diameter_field.bind(text=lambda inst, val: self.on_float_param_change("hole_diameter", val))

        self.ind_gas_pressure_field = MDTextField(
            hint_text="Тиск газу P, кг/м2", mode="rectangle", input_filter="float", size_hint_y=None, height="50dp"
        )
        self.ind_gas_pressure_field.bind(text=lambda inst, val: self.on_float_param_change("gas_pressure", val))

        self.ind_specific_volume_field = MDTextField(
            hint_text="Питомий обсяг газу v, м3/кг", mode="rectangle", input_filter="float", size_hint_y=None, height="50dp"
        )
        self.ind_specific_volume_field.bind(text=lambda inst, val: self.on_float_param_change("specific_volume", val))

        self.ind_hole_mode_container.add_widget(self.ind_hole_diameter_field)
        self.ind_hole_mode_container.add_widget(self.ind_gas_pressure_field)
        self.ind_hole_mode_container.add_widget(self.ind_specific_volume_field)

        self.industrial_form.add_widget(self.ind_equipment_mode_container)
        self.industrial_form.add_widget(self.ind_hole_mode_container)

        # Панель кнопок для промисловості (Розрахувати та Скинути)
        ind_btn_layout = MDBoxLayout(
            orientation="horizontal", spacing="10dp", size_hint_y=None, height="50dp"
        )
        self.ind_calculate_btn = MDRaisedButton(
            text="РОЗРАХУВАТИ (ПРОМИСЛОВІСТЬ)",
            size_hint_x=0.6,
            height="50dp",
            md_bg_color=self.theme_cls.primary_color,
        )
        self.ind_calculate_btn.bind(on_release=self.on_calculate_click)

        self.reset_ind_btn = MDFlatButton(
            text="СКИНУТИ",
            size_hint_x=0.4,
            height="50dp",
            theme_text_color="Custom",
            text_color=self.theme_cls.primary_color,
        )
        self.reset_ind_btn.bind(on_release=lambda x: self.reset_form("industrial"))

        ind_btn_layout.add_widget(self.ind_calculate_btn)
        ind_btn_layout.add_widget(self.reset_ind_btn)
        self.industrial_form.add_widget(ind_btn_layout)

    def reset_form(self, form_type):
        """Метод повного очищення полів та скидання стану форми"""
        if form_type == "household":
            self.account_number = ""
            self.account_field.text = ""
            self.selected_violation_code = None
            self.selected_violation_text = None
            self.violation_field.text = ""
            self.hidden_option_layout.opacity = 0
            self.hidden_checkbox.active = False
            self.is_hidden_violation = False
            self.date_from = None
            self.date_to = None
            self.date_from_field.text = ""
            self.date_to_field.text = ""
            self.days_label.text = "Всього днів: 0 (0 год.)"
            self.days_label.theme_text_color = "Primary"
            self.period_details_label.text = "• Опалювальний період: 0 днів (0 год.)\n• Неопалювальний період: 0 днів (0 год.)"
            self.registered_persons = 0
            self.persons_field.text = ""
            self.heated_area = 0.0
            self.area_field.text = ""
            self.hole_diameter = 0.0
            self.gas_pressure = 0.0
            self.specific_volume = 0.0
            self.hole_diameter_field.text = ""
            self.gas_pressure_field.text = ""
            self.specific_volume_field.text = ""
            self.hole_formula_layout.opacity = 0
            self.hole_formula_layout.disabled = True
            self.selected_equipment.clear()
            for child in self.equipment_checkboxes_layout.children:
                for widget in child.children:
                    if isinstance(widget, MDCheckbox):
                        widget.active = False
        else:
            self.eic_code = ""
            self.ind_eic_field.text = ""
            self.selected_violation_code = None
            self.selected_violation_text = None
            self.ind_violation_field.text = ""
            self.ind_hidden_option_layout.opacity = 0
            self.ind_hidden_checkbox.active = False
            self.is_hidden_violation = False
            self.ind_datetime_from = None
            self.ind_datetime_to = None
            self.ind_datetime_from_field.text = ""
            self.ind_datetime_to_field.text = ""
            self.ind_duration_label.text = "Тривалість: 0 год. 0 хв. (0 діб)"
            self.ind_duration_label.theme_text_color = "Primary"
            self.industrial_calculation_mode = "equipment"
            self.industrial_equipment_count = 0
            self.ind_equipment_count_field.text = ""
            self.ind_equipment_dynamic_layout.clear_widgets()
            self.industrial_equipment_items = []
            self.hole_diameter = 0.0
            self.gas_pressure = 0.0
            self.specific_volume = 0.0
            self.ind_hole_diameter_field.text = ""
            self.ind_gas_pressure_field.text = ""
            self.ind_specific_volume_field.text = ""

        self.total_days = 0
        self.total_hours = 0.0
        self.heating_days = 0
        self.heating_hours = 0.0
        self.offseason_days = 0
        self.offseason_hours = 0.0

    def on_industrial_mode_change(self, instance_segmented_control, active_item):
        selected_text = active_item.text
        if "потужністю" in selected_text:
            self.industrial_calculation_mode = "equipment"
            self.ind_equipment_mode_container.opacity = 1
            self.ind_equipment_mode_container.disabled = False
            self.ind_hole_mode_container.opacity = 0
            self.ind_hole_mode_container.disabled = True
        else:
            self.industrial_calculation_mode = "hole"
            self.ind_equipment_mode_container.opacity = 0
            self.ind_equipment_mode_container.disabled = True
            self.ind_hole_mode_container.opacity = 1
            self.ind_hole_mode_container.disabled = False

    def on_industrial_equipment_count_change(self, instance, value):
        try:
            count = int(value) if value else 0
        except ValueError:
            count = 0

        self.industrial_equipment_count = count
        self.ind_equipment_dynamic_layout.clear_widgets()
        self.industrial_equipment_items = []

        if count > 0:
            for i in range(count):
                row_layout = MDBoxLayout(
                    orientation="horizontal", spacing="10dp", size_hint_y=None, height="50dp"
                )
                
                name_field = MDTextField(
                    hint_text=f"Назва обладнання #{i+1}",
                    mode="rectangle",
                    size_hint_x=0.4,
                )
                
                power_kw_field = MDTextField(
                    hint_text=f"Потужність (кВт) #{i+1}",
                    mode="rectangle",
                    input_filter="float",
                    size_hint_x=0.3,
                )
                
                flow_field = MDTextField(
                    hint_text="Витрата (м3/год)",
                    mode="rectangle",
                    readonly=True,
                    size_hint_x=0.3,
                )

                power_kw_field.flow_field_ref = flow_field
                power_kw_field.bind(text=self.on_power_field_text_change)

                row_layout.add_widget(name_field)
                row_layout.add_widget(power_kw_field)
                row_layout.add_widget(flow_field)

                self.ind_equipment_dynamic_layout.add_widget(row_layout)
                self.industrial_equipment_items.append({
                    "name_field": name_field,
                    "power_kw_field": power_kw_field,
                    "flow_field": flow_field
                })

    def on_power_field_text_change(self, instance, value):
        try:
            kw_val = float(value.replace(",", ".")) if value else 0.0
        except ValueError:
            kw_val = 0.0

        if hasattr(instance, "flow_field_ref") and instance.flow_field_ref:
            if kw_val > 0:
                flow_cubic_m = kw_val / 8.2
                instance.flow_field_ref.text = f"{flow_cubic_m:.3f}"
            else:
                instance.flow_field_ref.text = ""

    def create_equipment_checkboxes(self):
        self.equipment_checkboxes_layout.clear_widgets()
        for key, info in self.EQUIPMENT_DATA.items():
            row = MDBoxLayout(orientation="horizontal", spacing="8dp", size_hint_y=None, height="40dp")
            chk = MDCheckbox(size_hint=(None, None), size=("36dp", "36dp"), pos_hint={"center_y": 0.5})
            chk.bind(active=lambda instance, val, eq_key=key: self.on_equipment_toggle(eq_key, val))
            lbl = MDLabel(
                text=f"{info['name']} ({info['unit']})",
                font_style="Body2",
                theme_text_color="Primary",
                pos_hint={"center_y": 0.5},
            )
            row.add_widget(chk)
            row.add_widget(lbl)
            self.equipment_checkboxes_layout.add_widget(row)

    def on_account_change(self, instance, value):
        self.account_number = value.strip()

    def on_equipment_toggle(self, key, is_active):
        if is_active:
            self.selected_equipment.add(key)
        else:
            self.selected_equipment.discard(key)

        if key == "hole_diameter_calc":
            if is_active:
                self.hole_formula_layout.opacity = 1
                self.hole_formula_layout.disabled = False
            else:
                self.hole_formula_layout.opacity = 0
                self.hole_formula_layout.disabled = True
                self.hole_diameter_field.text = ""
                self.gas_pressure_field.text = ""
                self.specific_volume_field.text = ""

    def on_float_param_change(self, param_name, value):
        try:
            val = float(value.replace(",", ".")) if value else 0.0
            setattr(self, param_name, val)
        except ValueError:
            setattr(self, param_name, 0.0)

    def on_persons_change(self, instance, value):
        try:
            self.registered_persons = int(value) if value else 0
        except ValueError:
            self.registered_persons = 0

    def on_area_change(self, instance, value):
        try:
            formatted_val = value.replace(",", ".")
            self.heated_area = float(formatted_val) if formatted_val else 0.0
        except ValueError:
            self.heated_area = 0.0

    def on_household_date_from_change(self, instance, value):
        try:
            self.date_from = datetime.strptime(value.strip(), "%d.%m.%Y").date()
            self.calculate_days()
        except ValueError:
            self.date_from = None
            self.reset_day_counters()

    def on_household_date_to_change(self, instance, value):
        try:
            self.date_to = datetime.strptime(value.strip(), "%d.%m.%Y").date()
            self.calculate_days()
        except ValueError:
            self.date_to = None
            self.reset_day_counters()

    def calculate_days(self):
        if self.date_from and self.date_to:
            today = datetime.now().date()

            if self.date_to > today or self.date_to < self.date_from:
                self.days_label.text = "Помилка вибору періоду!"
                self.days_label.theme_text_color = "Error"
                self.reset_day_counters()
            else:
                self.heating_days = 0
                self.offseason_days = 0
                current_date = self.date_from
                while current_date <= self.date_to:
                    if current_date.month in [10, 11, 12, 1, 2, 3, 4]:
                        self.heating_days += 1
                    else:
                        self.offseason_days += 1
                    current_date += timedelta(days=1)

                self.total_days = self.heating_days + self.offseason_days
                self.total_hours = float(self.total_days * 24)
                self.heating_hours = float(self.heating_days * 24)
                self.offseason_hours = float(self.offseason_days * 24)

                self.days_label.text = f"Всього днів: {self.total_days} ({int(self.total_hours)} год.)"
                self.days_label.theme_text_color = "Primary"
                self.period_details_label.text = (
                    f"• Опалювальний період: {self.heating_days} дн. ({int(self.heating_hours)} год.)\n"
                    f"• Неопалювальний період: {self.offseason_days} дн. ({int(self.offseason_hours)} год.)"
                )

    def on_ind_datetime_from_change(self, instance, value):
        try:
            self.ind_datetime_from = datetime.strptime(value.strip(), "%d.%m.%Y %H:%M")
            self.calculate_industrial_duration()
        except ValueError:
            self.ind_datetime_from = None

    def on_ind_datetime_to_change(self, instance, value):
        try:
            self.ind_datetime_to = datetime.strptime(value.strip(), "%d.%m.%Y %H:%M")
            self.calculate_industrial_duration()
        except ValueError:
            self.ind_datetime_to = None

    def calculate_industrial_duration(self):
        if self.ind_datetime_from and self.ind_datetime_to:
            if self.ind_datetime_to < self.ind_datetime_from:
                self.ind_duration_label.text = "Помилка: кінець раніше за початок!"
                self.ind_duration_label.theme_text_color = "Error"
                self.total_hours = 0.0
                self.total_days = 0
            else:
                delta = self.ind_datetime_to - self.ind_datetime_from
                total_seconds = delta.total_seconds()
                self.total_hours = total_seconds / 3600.0
                self.total_days = math.ceil(self.total_hours / 24.0) if self.total_hours > 0 else 0

                hours_int = int(total_seconds // 3600)
                minutes_int = int((total_seconds % 3600) // 60)
                days_float = self.total_hours / 24.0

                self.ind_duration_label.text = (
                    f"Тривалість: {hours_int} год. {minutes_int} хв. ({days_float:.2f} діб)"
                )
                self.ind_duration_label.theme_text_color = "Primary"

    def reset_day_counters(self):
        self.total_days = 0
        self.total_hours = 0.0
        self.heating_days = 0
        self.heating_hours = 0.0
        self.offseason_days = 0
        self.offseason_hours = 0.0

    def show_violation_dialog(self, instance):
        scroll = MDScrollView(size_hint_y=None, height="320dp")
        list_box = MDBoxLayout(
            orientation="vertical", spacing="8dp", size_hint_y=None
        )
        list_box.bind(minimum_height=list_box.setter("height"))

        current_dict = (
            self.household_violations_dict
            if self.consumer_type == "household"
            else self.industrial_violations_dict
        )

        for code, text in current_dict.items():
            btn = MDFlatButton(
                text=text,
                size_hint_x=1,
                size_hint_y=None,
                height="50dp",
            )
            btn.ids.text_color = self.theme_cls.primary_color
            btn.bind(on_release=lambda x, c=code, t=text: self.set_violation(c, t))
            list_box.add_widget(btn)

        scroll.add_widget(list_box)

        title_text = (
            "Оберіть порушення (Побутовий)"
            if self.consumer_type == "household"
            else "Оберіть порушення (Промисловий)"
        )

        self.dialog = MDDialog(
            title=title_text,
            type="custom",
            content_cls=scroll,
            buttons=[
                MDFlatButton(
                    text="СКАСУВАТИ",
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release=lambda x: self.dialog.dismiss(),
                )
            ],
        )
        self.dialog.open()

    def set_violation(self, code, text):
        self.selected_violation_code = code
        self.selected_violation_text = text
        self.violation_field.text = text
        self.ind_violation_field.text = text

        if code in ["п. 1.1", "п. 1.3"]:
            self.hidden_option_layout.opacity = 1
            self.hidden_checkbox.disabled = False
            self.ind_hidden_option_layout.opacity = 1
            self.ind_hidden_checkbox.disabled = False
        else:
            self.hidden_option_layout.opacity = 0
            self.hidden_checkbox.disabled = True
            self.ind_hidden_option_layout.opacity = 0
            self.ind_hidden_checkbox.disabled = True
            self.is_hidden_violation = False
            self.hidden_checkbox.active = False
            self.ind_hidden_checkbox.active = False

        if self.dialog:
            self.dialog.dismiss()

    def on_hidden_checkbox_change(self, checkbox, value):
        self.is_hidden_violation = value
        self.hidden_checkbox.active = value
        self.ind_hidden_checkbox.active = value

        if value:
            self.hidden_checkbox_label.text = "Приховане"
            self.hidden_checkbox_label.theme_text_color = "Error"
            self.ind_hidden_checkbox_label.text = "Приховане"
            self.ind_hidden_checkbox_label.theme_text_color = "Error"
        else:
            self.hidden_checkbox_label.text = "Неприховане"
            self.hidden_checkbox_label.theme_text_color = "Secondary"
            self.ind_hidden_checkbox_label.text = "Неприховане"
            self.ind_hidden_checkbox_label.theme_text_color = "Secondary"

    def on_consumer_type_change(self, instance_segmented_control, active_item):
        selected_text = active_item.text
        if self.dialog:
            self.dialog.dismiss()

        self.forms_container.clear_widgets()
        if "Побутовий" in selected_text:
            self.consumer_type = "household"
            self.forms_container.add_widget(self.household_form)
        else:
            self.consumer_type = "industrial"
            self.forms_container.add_widget(self.industrial_form)

    def on_calculate_click(self, instance):
        if not self.selected_violation_code:
            self.show_alert_dialog("Помилка", "Оберіть вид порушення Кодексу ГРМ!")
            return

        if self.consumer_type == "household":
            if not self.date_from or not self.date_to or self.total_days <= 0:
                self.show_alert_dialog("Помилка", "Введіть коректний період у форматі ДД.ММ.РРРР!")
                return
            if not self.selected_equipment:
                self.show_alert_dialog("Помилка", "Оберіть принаймні один тип обладнання або розрахунок за формулою!")
                return
        else:
            if not self.ind_datetime_from or not self.ind_datetime_to or self.total_hours <= 0:
                self.show_alert_dialog("Помилка", "Введіть коректні дату та час у форматі ДД.ММ.РРРР ГГ:ХХ!")
                return
            if self.industrial_calculation_mode == "equipment" and self.industrial_equipment_count <= 0:
                self.show_alert_dialog("Помилка", "Вкажіть кількість неопломбованого обладнання!")
                return
            if self.industrial_calculation_mode == "hole" and (self.hole_diameter <= 0 or self.gas_pressure <= 0 or self.specific_volume <= 0):
                self.show_alert_dialog("Помилка", "Вкажіть коректні параметри отвору (діаметр, тиск, питомий обсяг)!")
                return

        total_volume = 0.0
        details_text = ""
        norm_factor = 0.7 if self.selected_violation_code == "п. 1.5" else 1.0

        if self.consumer_type == "household":
            is_hole_calc = "hole_diameter_calc" in self.selected_equipment

            if is_hole_calc:
                if self.hole_diameter <= 0 or self.gas_pressure <= 0 or self.specific_volume <= 0:
                    self.show_alert_dialog(
                        "Помилка", "Вкажіть коректні параметри отвору (діаметр d, тиск P, питомий обсяг v повинні бути > 0)!"
                    )
                    return

                g = 9.81
                k = 1.4
                area = (math.pi * (self.hole_diameter**2)) / 4.0
                adiabatic_factor = (2.0 * k) / (k + 1.0)
                velocity = math.sqrt(adiabatic_factor * g * self.gas_pressure * self.specific_volume)
                hourly_flow = 3600.0 * area * velocity
                total_volume = round(hourly_flow * self.total_hours, 4)

                details_text = (
                    f"РОЗРАХУНКОВА ФОРМУЛА ВИТОКУ ГАЗУ:\n"
                    f" Q = 3600 * f * w * T\n"
                    f" w = sqrt[ (2*k / (k + 1)) * g * P * v ]\n"
                    f" f = (pi * d^2) / 4\n\n"
                    f"ВХІДНІ ДАНІ ФОРМУЛИ:\n"
                    f"• Діаметр отвору (d): {self.hole_diameter} м\n"
                    f"• Тиск газу (P): {self.gas_pressure} кг/м2\n"
                    f"• Питомий обсяг газу (v): {self.specific_volume} м3/кг\n"
                    f"• Показник адіабати (k): 1.4\n"
                    f"• Прискорення вільн. падіння (g): 9.81 м/с2\n"
                    f"• Час витоку (T): {self.total_hours:.2f} год.\n\n"
                    f"ПОКРОКОВІ ОБЧИСЛЕННЯ:\n"
                    f"1) Площа отвору f = (pi * {self.hole_diameter}^2) / 4 = {area:.8f} м2\n"
                    f"2) Адіабатний коефіцієнт = {adiabatic_factor:.4f}\n"
                    f"3) Швидкість витоку w = {velocity:.4f} м/с\n"
                    f"4) Годинна витрата Qгод = {hourly_flow:.4f} м3/год\n"
                    f"5) Об'єм за весь період Qзаг = {total_volume} м3\n"
                )

            else:
                has_people_app = any(self.EQUIPMENT_DATA[eq]["type"] == "people" for eq in self.selected_equipment)
                has_area_app = any(self.EQUIPMENT_DATA[eq]["type"] == "area" for eq in self.selected_equipment)
                has_hybrid_app = "heating_stove" in self.selected_equipment

                if (has_people_app or has_hybrid_app) and self.registered_persons <= 0:
                    self.show_alert_dialog("Помилка", "Вкажіть кількість зареєстрованих осіб!")
                    return

                if (has_area_app or has_hybrid_app) and self.heated_area <= 0:
                    self.show_alert_dialog("Помилка", "Вкажіть опалювальну площу об'єкта!")
                    return

                details_text = "МЕТОД: За нормами споживання газових приладів (Додаток 15 Кодексу ГРМ)\n"
                if norm_factor == 0.7:
                    details_text += " (Застосовано коефіцієнт 0.7 за п. 1.5 Кодексу ГРМ)\n\n"
                else:
                    details_text += "\n"

                for eq in self.selected_equipment:
                    eq_info = self.EQUIPMENT_DATA[eq]

                    if eq_info["type"] == "people":
                        vol = ((self.registered_persons * eq_info["rate"] * self.total_days) / 30.0) * norm_factor
                        total_volume += vol
                        factor_str = " * 0.7" if norm_factor == 0.7 else ""
                        details_text += (
                            f"• {eq_info['name']}:\n"
                            f" {self.registered_persons} осіб * {eq_info['rate']} м3 * ({self.total_days}/30 дн){factor_str} = {vol:.3f} м3\n"
                        )

                    elif eq_info["type"] == "area":
                        vol_heat = ((self.heated_area * eq_info["rate_heating"] * self.heating_days) / 30.0) * norm_factor
                        vol_offseason = ((self.heated_area * (eq_info["rate_heating"] * 0.3) * self.offseason_days) / 30.0) * norm_factor
                        vol_eq = vol_heat + vol_offseason
                        total_volume += vol_eq
                        factor_str = " * 0.7" if norm_factor == 0.7 else ""

                        details_text += (
                            f"• {eq_info['name']}:\n"
                            f" - Опалювальний пер. ({self.heating_days} дн): {self.heated_area} м2 * {eq_info['rate_heating']} м3 * ({self.heating_days}/30){factor_str} = {vol_heat:.3f} м3\n"
                            f" - Неопалювальний пер. 30% ({self.offseason_days} дн): {self.heated_area} м2 * {eq_info['rate_heating']*0.3:.2f} м3 * ({self.offseason_days}/30){factor_str} = {vol_offseason:.3f} м3\n"
                        )

                    elif eq_info["type"] == "hybrid_stove":
                        # --- РОЗРАХУНОК ОПАЛЮВАЛЬНО-ВАРИЛЬНОЇ ПЕЧІ (ДОДАТОК 15) ---
                        
                        # 1. Міжопалювальний період (за кількістю осіб):
                        # 70 м3/міс (до 3 мешканців включно) + 45 м3/міс на кожного наступного
                        if self.registered_persons <= 3:
                            monthly_rate_offseason = 70.0
                        else:
                            monthly_rate_offseason = 70.0 + (self.registered_persons - 3) * 45.0

                        vol_offseason = ((monthly_rate_offseason * self.offseason_days) / 30.0) * norm_factor

                        # 2. Опалювальний період (за опалювальною площею):
                        # 20 м3/м2 (до 20 кв.м включно) + 11 м3/м2 на кожен наступний кв.м
                        if self.heated_area <= 20.0:
                            monthly_rate_heating = self.heated_area * 20.0
                        else:
                            monthly_rate_heating = (20.0 * 20.0) + ((self.heated_area - 20.0) * 11.0)

                        vol_heat = ((monthly_rate_heating * self.heating_days) / 30.0) * norm_factor

                        vol_eq = vol_heat + vol_offseason
                        total_volume += vol_eq
                        factor_str = " * 0.7" if norm_factor == 0.7 else ""

                        details_text += (
                            f"• {eq_info['name']} (Граничні норми):\n"
                            f" - Міжопалювальний пер. ({self.offseason_days} дн, {self.registered_persons} осіб): "
                            f"{monthly_rate_offseason:.1f} м3/міс * ({self.offseason_days}/30){factor_str} = {vol_offseason:.3f} м3\n"
                            f" - Опалювальний пер. ({self.heating_days} дн, площа {self.heated_area} м2): "
                            f"{monthly_rate_heating:.1f} м3/міс * ({self.heating_days}/30){factor_str} = {vol_heat:.3f} м3\n"
                        )

                total_volume = round(total_volume, 3)

        else:
            if self.industrial_calculation_mode == "equipment":
                details_text = "МЕТОД: За номінальною потужністю неопломбованого обладнання (кВт -> м3/год)\n\n"
                total_power_kw = 0.0
                total_hourly_flow = 0.0

                for idx, item in enumerate(self.industrial_equipment_items):
                    eq_name = item["name_field"].text.strip() or f"Обладнання #{idx+1}"
                    try:
                        kw_val = float(item["power_kw_field"].text.replace(",", ".")) if item["power_kw_field"].text else 0.0
                    except ValueError:
                        kw_val = 0.0

                    if kw_val <= 0:
                        self.show_alert_dialog("Помилка", f"Вкажіть коректну потужність (кВт) для обладнання: {eq_name}")
                        return

                    hourly_flow_val = kw_val / 8.2
                    total_power_kw += kw_val
                    total_hourly_flow += hourly_flow_val

                    item_vol = hourly_flow_val * self.total_hours
                    total_volume += item_vol

                    details_text += (
                        f"• {eq_name}:\n"
                        f" Потужність: {kw_val} кВт | Год. витрата: {hourly_flow_val:.3f} м3/год * {self.total_hours:.2f} год. = {item_vol:.3f} м3\n"
                    )

                total_volume = round(total_volume, 3)
                details_text += f"\nСумарна потужність: {total_power_kw:.3f} кВт\n"
                details_text += f"Сумарна годинна витрата: {total_hourly_flow:.3f} м3/год\n"
            else:
                g = 9.81
                k = 1.4
                area = (math.pi * (self.hole_diameter**2)) / 4.0
                adiabatic_factor = (2.0 * k) / (k + 1.0)
                velocity = math.sqrt(adiabatic_factor * g * self.gas_pressure * self.specific_volume)
                hourly_flow = 3600.0 * area * velocity
                total_volume = round(hourly_flow * self.total_hours, 4)

                details_text = (
                    f"РОЗРАХУНКОВА ФОРМУЛА ВИТОКУ ГАЗУ (Промисловість):\n"
                    f" Q = 3600 * f * w * T\n"
                    f" w = sqrt[ (2*k / (k + 1)) * g * P * v ]\n"
                    f" f = (pi * d^2) / 4\n\n"
                    f"ВХІДНІ ДАНІ ФОРМУЛИ:\n"
                    f"• Діаметр отвору (d): {self.hole_diameter} м\n"
                    f"• Тиск газу (P): {self.gas_pressure} кг/м2\n"
                    f"• Питомий обсяг газу (v): {self.specific_volume} м3/кг\n"
                    f"• Показник адіабати (k): 1.4\n"
                    f"• Час витоку (T): {self.total_hours:.2f} год.\n\n"
                    f"ПОКРОКОВІ ОБЧИСЛЕННЯ:\n"
                    f"1) Площа отвору f = (pi * {self.hole_diameter}^2) / 4 = {area:.8f} м2\n"
                    f"2) Адіабатний коефіцієнт = {adiabatic_factor:.4f}\n"
                    f"3) Швидкість витоку w = {velocity:.4f} м/с\n"
                    f"4) Годинна витрата Qгод = {hourly_flow:.4f} м3/год\n"
                    f"5) Об'єм за весь період Qзаг = {total_volume} м3\n"
                )

        self.last_total_volume = total_volume
        self.last_details_text = details_text
        self.show_calculation_card(total_volume, details_text)

    def show_calculation_card(self, total_volume, details_text):
        if self.consumer_type == "household":
            acc_str = self.account_number if self.account_number else "Не вказано"
            category_str = "Побутовий споживач"
            date_from_str = self.date_from.strftime('%d.%m.%Y') if self.date_from else ""
            date_to_str = self.date_to.strftime('%d.%m.%Y') if self.date_to else ""
            duration_str = f"{self.total_days} днів ({int(self.total_hours)} годин)"
        else:
            acc_str = f"EIC: {self.eic_code}" if self.eic_code else "Не вказано"
            category_str = "Промисловий споживач"
            date_from_str = self.ind_datetime_from.strftime('%d.%m.%Y %H:%M') if self.ind_datetime_from else ""
            date_to_str = self.ind_datetime_to.strftime('%d.%m.%Y %H:%M') if self.ind_datetime_to else ""
            duration_str = f"{self.total_hours:.2f} годин"

        detection_method = "Не застосовується"
        if self.selected_violation_code in ["п. 1.1", "п. 1.3"]:
            detection_method = "Приховане" if self.is_hidden_violation else "Неприховане"

        card_content = (
            f"ЗАГАЛЬНА ІНФОРМАЦІЯ ({category_str}):\n"
            f"• Ідентифікатор / Рахунок: {acc_str}\n"
            f"• Вид порушення: {self.selected_violation_text}\n"
            f"• Метод виявлення: {detection_method}\n"
            f"• Період: з {date_from_str} по {date_to_str}\n"
            f"• Тривалість: {duration_str}\n\n"
            f"{details_text}"
            f"----------------------------------------\n"
            f"КІНЦЕВИЙ РЕЗУЛЬТАТ:\n"
            f"Всього донараховано: {total_volume:.3f} м3"
        )

        scroll = MDScrollView(size_hint_y=None, height="380dp")

        label_content = MDLabel(
            text=card_content,
            font_style="Body2",
            size_hint_y=None,
            theme_text_color="Primary",
        )
        label_content.bind(
            texture_size=lambda instance, value: setattr(instance, "height", value[1])
        )

        scroll.add_widget(label_content)

        self.dialog = MDDialog(
            title="Картка розрахунку",
            type="custom",
            content_cls=scroll,
            buttons=[
                MDRaisedButton(
                    text="ЕКСПОРТ У WORD",
                    md_bg_color=self.theme_cls.primary_color,
                    on_release=lambda x: self.export_to_word(),
                ),
                MDFlatButton(
                    text="ЗАКРИТИ",
                    on_release=lambda x: self.dialog.dismiss(),
                ),
            ],
        )
        self.dialog.open()

    def export_to_word(self):
        try:
            doc = Document()

            title = doc.add_heading("Акт / Картка розрахунку необлікованого об'єму природного газу", level=1)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            curr_time = datetime.now().strftime('%d.%m.%Y %H:%M')
            doc.add_paragraph(f"Дата формування: {curr_time}\n")

            doc.add_heading("1. Загальна інформація", level=2)
            
            if self.consumer_type == "household":
                acc_str = self.account_number if self.account_number else "Не вказано"
                cat_str = "Побутовий споживач (Населення)"
                date_from_str = self.date_from.strftime('%d.%m.%Y') if self.date_from else ""
                date_to_str = self.date_to.strftime('%d.%m.%Y') if self.date_to else ""
                duration_str = f"{self.total_days} дн. ({int(self.total_hours)} год.)"
            else:
                acc_str = self.eic_code if self.eic_code else "Не вказано"
                cat_str = "Промисловий споживач (Непобутовий)"
                date_from_str = self.ind_datetime_from.strftime('%d.%m.%Y %H:%M') if self.ind_datetime_from else ""
                date_to_str = self.ind_datetime_to.strftime('%d.%m.%Y %H:%M') if self.ind_datetime_to else ""
                duration_str = f"{self.total_hours:.2f} год."

            detection_method = "Не застосовується"
            if self.selected_violation_code in ["п. 1.1", "п. 1.3"]:
                detection_method = "Приховане" if self.is_hidden_violation else "Неприховане"

            period_str = f"з {date_from_str} по {date_to_str}"

            info_data = [
                ("Категорія споживача", cat_str),
                ("Ідентифікатор / Рахунок", acc_str),
                ("Вид порушення", str(self.selected_violation_text or "Не вказано")),
                ("Метод виявлення", detection_method),
                ("Період нарахування", period_str),
                ("Тривалість", duration_str),
            ]

            table_info = doc.add_table(rows=len(info_data), cols=2)
            table_info.style = 'Table Grid'

            for idx, (label, val) in enumerate(info_data):
                table_info.cell(idx, 0).text = label
                table_info.cell(idx, 1).text = str(val)

            doc.add_paragraph("")

            doc.add_heading("2. Деталізація та покроковий розрахунок", level=2)
            doc.add_paragraph(self.last_details_text)

            doc.add_heading("3. Кінцевий результат", level=2)
            res_p = doc.add_paragraph()
            res_text = f"Всього донараховано: {self.last_total_volume:.3f} м3"
            run = res_p.add_run(res_text)
            run.bold = True
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(0, 128, 128)

            # --- Збереження у папку "Завантаження" (Downloads) ---
            if os.name == 'nt':  # Windows
                downloads_path = os.path.join(os.environ.get('USERPROFILE', ''), 'Downloads')
            else:  # Android / Linux / macOS
                downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

            if not os.path.exists(downloads_path):
                os.makedirs(downloads_path, exist_ok=True)
                downloads_path = ""

            file_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Calculation_{acc_str}_{file_timestamp}.docx"
            full_path = os.path.join(downloads_path, filename) if downloads_path else filename

            doc.save(full_path)

            if self.dialog:
                self.dialog.dismiss()

            msg_success = f"Документ успішно збережено у папку Завантаження:\n{full_path}"
            self.show_alert_dialog("Успіх", msg_success)

        except Exception as e:
            err_details = str(e)
            msg_error = f"Не вдалося зберегти документ Word:\n{err_details}"
            self.show_alert_dialog("Помилка експорту", msg_error)

    def show_alert_dialog(self, title, message):
        self.dialog = MDDialog(
            title=title,
            text=message,
            buttons=[
                MDFlatButton(
                    text="ОК",
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release=lambda x: self.dialog.dismiss(),
                )
            ],
        )
        self.dialog.open()


if __name__ == "__main__":
    GasCalculatorApp().run()