import sys
import os
import re
import traceback
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk
from generator import KPGenerator
from contract_generator import ContractGenerator


def log_error(error_msg):
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(error_msg + "\n" + "-" * 20 + "\n")
    except Exception:
        print(error_msg)


class KPApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("MetaManager Pro")
        self.geometry("820x750")
        self.minsize(700, 550)
        self.resizable(True, True)

        # Текущий режим: "kp" или "contract"
        self.current_mode = "kp"

        # ---------------------------------------------------------------
        # Основная компоновка
        # ---------------------------------------------------------------
        self.grid_rowconfigure(0, weight=0)   # переключатель режимов
        self.grid_rowconfigure(1, weight=1)   # прокручиваемая область
        self.grid_rowconfigure(2, weight=0)   # кнопка генерации
        self.grid_columnconfigure(0, weight=1)

        # --- Переключатель режимов ---
        self.mode_bar = ctk.CTkFrame(self, fg_color=("gray90", "gray20"), corner_radius=0)
        self.mode_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.mode_bar.grid_columnconfigure(0, weight=1)
        self.mode_bar.grid_columnconfigure(1, weight=1)

        self.kp_mode_btn = ctk.CTkButton(
            self.mode_bar,
            text="  Коммерческое предложение",
            command=lambda: self.switch_mode("kp"),
            corner_radius=0,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("gray75", "gray35"),
            hover_color=("gray65", "gray45"),
        )
        self.kp_mode_btn.grid(row=0, column=0, sticky="ew", padx=(0, 1), pady=0)

        self.contract_mode_btn = ctk.CTkButton(
            self.mode_bar,
            text="  Договор",
            command=lambda: self.switch_mode("contract"),
            corner_radius=0,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("gray90", "gray20"),
            hover_color=("gray65", "gray45"),
        )
        self.contract_mode_btn.grid(row=0, column=1, sticky="ew", padx=(1, 0), pady=0)

        # --- Прокручиваемая область ---
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        # --- Нижняя панель с кнопкой ---
        self.bottom_bar = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=(5, 15))
        self.bottom_bar.grid_columnconfigure(0, weight=1)

        self.generate_button = ctk.CTkButton(
            self.bottom_bar,
            text="Сгенерировать КП",
            command=self.on_generate,
            height=44,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.generate_button.grid(row=0, column=0)

        # --- Строим содержимое ---
        self._build_kp_content()
        self._build_contract_content()

        # Показываем режим КП по умолчанию
        self.switch_mode("kp")

        # Настройка навигации клавиатурой
        self._setup_keyboard_navigation()
        self._setup_global_shortcuts()

    # -----------------------------------------------------------------------
    # Переключение режимов
    # -----------------------------------------------------------------------

    def switch_mode(self, mode):
        self.current_mode = mode

        # Скрываем только виджеты режимов, не трогая внутренности scrollable-frame
        self._hide_mode_widgets(self.kp_widgets)
        self._hide_mode_widgets(self.contract_widgets)

        if mode == "kp":
            self.kp_mode_btn.configure(fg_color=("gray75", "gray35"))
            self.contract_mode_btn.configure(fg_color=("gray90", "gray20"))
            self._show_kp_content()
            self.generate_button.configure(text="Сгенерировать КП")
        else:
            self.kp_mode_btn.configure(fg_color=("gray90", "gray20"))
            self.contract_mode_btn.configure(fg_color=("gray75", "gray35"))
            self._show_contract_content()
            self.generate_button.configure(text="Сгенерировать Договор")

    def _hide_mode_widgets(self, widget_defs):
        """Скрывает только виджеты, которые реально размещены через grid."""
        for _, widget, _ in widget_defs:
            try:
                if widget.winfo_manager() == "grid":
                    widget.grid_forget()
            except Exception:
                continue

    # -----------------------------------------------------------------------
    # Построение содержимого КП
    # -----------------------------------------------------------------------

    def _build_kp_content(self):
        """Создаёт все виджеты для режима КП (не отображает их)."""
        f = self.scroll_frame

        self.kp_widgets = []  # список (kind, widget, row)

        # --- Заголовок ---
        lbl_title = ctk.CTkLabel(f, text="Коммерческое предложение",
                                  font=ctk.CTkFont(size=20, weight="bold"))
        self.kp_widgets.append(("label", lbl_title, 0))

        # --- Название КП ---
        lbl_kp_name = ctk.CTkLabel(f, text="Название КП:")
        self.kp_widgets.append(("label", lbl_kp_name, 1))
        self.kp_name_entry = ctk.CTkEntry(
            f, placeholder_text="Введите название файла КП", width=550
        )
        self.kp_widgets.append(("entry", self.kp_name_entry, 2))
        self._bind_entry_keys(self.kp_name_entry)

        # --- Титульный лист КП ---
        lbl_kp_title = ctk.CTkLabel(f, text="Титульный лист КП:")
        self.kp_widgets.append(("label", lbl_kp_title, 3))
        self.kp_title_entry = ctk.CTkEntry(
            f, placeholder_text="Введите текст для титульного листа КП", width=550
        )
        self.kp_widgets.append(("entry", self.kp_title_entry, 4))
        self._bind_entry_keys(self.kp_title_entry)

        # --- Путь сохранения ---
        lbl_save = ctk.CTkLabel(f, text="Путь сохранения:")
        self.kp_widgets.append(("label", lbl_save, 5))

        self.save_path_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.kp_widgets.append(("frame", self.save_path_frame, 6))

        desktop = (
            os.path.join(os.environ['USERPROFILE'], 'Desktop')
            if 'USERPROFILE' in os.environ
            else os.path.expanduser("~/Desktop")
        )
        self.save_path_var = ctk.StringVar(value=desktop)
        self.save_path_entry = ctk.CTkEntry(self.save_path_frame, textvariable=self.save_path_var, width=500)
        self.save_path_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.save_path_entry)
        ctk.CTkButton(self.save_path_frame, text="Обзор", width=80, command=self.browse_folder).pack(side="left")

        # --- Ветка ---
        lbl_branch = ctk.CTkLabel(f, text="Выберите ветку:")
        self.kp_widgets.append(("label", lbl_branch, 7))
        self.branch_var = ctk.StringVar(value="хоз.пит")
        self.branch_menu = ctk.CTkOptionMenu(
            f, values=["хоз.пит", "техническая лицензия"],
            variable=self.branch_var,
            command=self.on_branch_change
        )
        self.kp_widgets.append(("menu", self.branch_menu, 8))

        # --- Объём водопотребления ---
        lbl_vol = ctk.CTkLabel(f, text="Объем водопотребления (м3/сут):")
        self.kp_widgets.append(("label", lbl_vol, 9))
        self.volume_var = ctk.StringVar(value="до 100")
        self.volume_menu = ctk.CTkOptionMenu(
            f, values=["до 100", "100-500", "500+", "500+ с переоценкой запасов"],
            variable=self.volume_var
        )
        self.kp_widgets.append(("menu", self.volume_menu, 10))

        # --- Тип работ (СМР) ---
        lbl_smr = ctk.CTkLabel(f, text="Тип работ:")
        self.kp_widgets.append(("label", lbl_smr, 11))
        self.smr_var = ctk.StringVar(value="без смр")
        self.smr_menu = ctk.CTkOptionMenu(
            f, values=["с смр", "без смр"],
            variable=self.smr_var,
            command=self.toggle_smr_fields
        )
        self.kp_widgets.append(("menu", self.smr_menu, 12))

        # --- ПИР ---
        self.include_pir = ctk.BooleanVar(value=False)
        self.pir_cb = ctk.CTkCheckBox(
            f,
            text="ПИР",
            variable=self.include_pir,
            command=self.toggle_pir_fields
        )
        self.kp_widgets.append(("check", self.pir_cb, 13))

        self.pir_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.kp_widgets.append(("frame_pir", self.pir_frame, 14))
        ctk.CTkLabel(self.pir_frame, text="Количество ПИР:").pack(side="left", padx=(0, 8))
        self.pir_count_entry = ctk.CTkEntry(self.pir_frame, width=70)
        self.pir_count_entry.insert(0, "1")
        self.pir_count_entry.pack(side="left", padx=(0, 16))
        self._bind_entry_keys(self.pir_count_entry)

        ctk.CTkLabel(self.pir_frame, text="Цена за 1 ПИР:").pack(side="left", padx=(0, 8))
        self.pir_price_entry = ctk.CTkEntry(self.pir_frame, width=140)
        self.pir_price_entry.pack(side="left")
        self._bind_entry_keys(self.pir_price_entry)

        # --- Количество скважин (для режима без СМР) ---
        self.no_smr_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.kp_widgets.append(("frame_nosmr", self.no_smr_frame, 15))
        ctk.CTkLabel(self.no_smr_frame, text="Количество скважин:").pack(side="left", padx=(0, 10))
        self.no_smr_wells_count = ctk.CTkEntry(self.no_smr_frame, width=60)
        self.no_smr_wells_count.insert(0, "1")
        self.no_smr_wells_count.pack(side="left")
        self._bind_entry_keys(self.no_smr_wells_count)

        # --- Поля для режима с СМР ---
        self.smr_frame = ctk.CTkFrame(f)
        self.kp_widgets.append(("frame_smr", self.smr_frame, 16))
        self.setup_smr_fields()

    def _show_kp_content(self):
        """Отображает виджеты режима КП."""
        for kind, widget, row in self.kp_widgets:
            if kind == "frame_smr":
                if self.smr_var.get() == "с смр":
                    widget.grid(row=row, column=0, padx=20, pady=(0, 15), sticky="nsew")
            elif kind == "frame_nosmr":
                if self.smr_var.get() != "с смр":
                    widget.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="w")
            elif kind == "frame_pir":
                if self.include_pir.get():
                    widget.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="w")
            elif kind == "label":
                pady = (20, 0) if row == 0 else (10, 0)
                widget.grid(row=row, column=0, padx=20, pady=pady, sticky="w")
            elif kind in ("entry", "menu", "check"):
                widget.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="w")
            elif kind == "frame":
                widget.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="ew")

    # -----------------------------------------------------------------------
    # Построение содержимого Договора
    # -----------------------------------------------------------------------

    def _build_contract_content(self):
        """Создаёт все виджеты для режима Договор."""
        f = self.scroll_frame
        self.contract_widgets = []

        # --- Заголовок ---
        lbl_title = ctk.CTkLabel(f, text="Договор подряда",
                                  font=ctk.CTkFont(size=20, weight="bold"))
        self.contract_widgets.append(("label", lbl_title, 0))

        # --- Номер договора ---
        lbl_num = ctk.CTkLabel(f, text="Номер договора (после 0ЦЦБ-):")
        self.contract_widgets.append(("label", lbl_num, 1))
        self.contract_number_entry = ctk.CTkEntry(f, placeholder_text="Например: 0025", width=300)
        self.contract_widgets.append(("entry", self.contract_number_entry, 2))
        self._bind_entry_keys(self.contract_number_entry)

        # --- Путь сохранения договора ---
        lbl_save = ctk.CTkLabel(f, text="Путь сохранения:")
        self.contract_widgets.append(("label", lbl_save, 3))

        self.contract_save_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.contract_save_frame, 4))

        desktop = (
            os.path.join(os.environ['USERPROFILE'], 'Desktop')
            if 'USERPROFILE' in os.environ
            else os.path.expanduser("~/Desktop")
        )
        self.contract_save_var = ctk.StringVar(value=desktop)
        self.contract_save_entry = ctk.CTkEntry(
            self.contract_save_frame, textvariable=self.contract_save_var, width=450)
        self.contract_save_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.contract_save_entry)
        ctk.CTkButton(self.contract_save_frame, text="Обзор", width=80,
                      command=self.browse_contract_folder).pack(side="left")

        # --- Разделитель: Данные заказчика ---
        sep1 = ctk.CTkLabel(f, text="─── Данные заказчика ───────────────────────────────",
                             font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"))
        self.contract_widgets.append(("label_sep", sep1, 5))

        # --- ИНН заказчика ---
        lbl_inn = ctk.CTkLabel(f, text="ИНН заказчика:")
        self.contract_widgets.append(("label", lbl_inn, 6))

        self.inn_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.inn_frame, 7))

        self.inn_entry = ctk.CTkEntry(self.inn_frame, placeholder_text="Введите ИНН", width=220)
        self.inn_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.inn_entry)

        self.fetch_btn = ctk.CTkButton(
            self.inn_frame, text="Получить данные по ИНН", width=210,
            command=self.fetch_company_data
        )
        self.fetch_btn.pack(side="left")

        # --- Статус загрузки ---
        self.inn_status_label = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=11),
                                              text_color=("gray50", "gray60"))
        self.contract_widgets.append(("label", self.inn_status_label, 8))

        # --- Наименование заказчика (авто) ---
        lbl_cname = ctk.CTkLabel(f, text="Полное наименование заказчика:")
        self.contract_widgets.append(("label", lbl_cname, 9))
        self.customer_fullname_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически по ИНН", width=650)
        self.contract_widgets.append(("entry", self.customer_fullname_entry, 10))
        self._bind_entry_keys(self.customer_fullname_entry)

        # --- Краткое наименование (авто) ---
        lbl_cshort = ctk.CTkLabel(f, text="Краткое наименование заказчика:")
        self.contract_widgets.append(("label", lbl_cshort, 11))
        self.customer_shortname_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически по ИНН", width=450)
        self.contract_widgets.append(("entry", self.customer_shortname_entry, 12))
        self._bind_entry_keys(self.customer_shortname_entry)

        # --- Адрес (авто) ---
        lbl_addr = ctk.CTkLabel(f, text="Юр./Факт./Почт. адрес:")
        self.contract_widgets.append(("label", lbl_addr, 13))
        self.customer_address_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически по ИНН", width=650)
        self.contract_widgets.append(("entry", self.customer_address_entry, 14))
        self._bind_entry_keys(self.customer_address_entry)

        # --- ОГРН (авто) ---
        lbl_ogrn = ctk.CTkLabel(f, text="ОГРН:")
        self.contract_widgets.append(("label", lbl_ogrn, 15))
        self.customer_ogrn_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически", width=250)
        self.contract_widgets.append(("entry", self.customer_ogrn_entry, 16))
        self._bind_entry_keys(self.customer_ogrn_entry)

        # --- ИНН/КПП (авто) ---
        lbl_innkpp = ctk.CTkLabel(f, text="ИНН / КПП:")
        self.contract_widgets.append(("label", lbl_innkpp, 17))

        self.innkpp_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.innkpp_frame, 18))

        self.customer_inn_entry = ctk.CTkEntry(self.innkpp_frame, placeholder_text="ИНН", width=200)
        self.customer_inn_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.customer_inn_entry)
        self.customer_kpp_entry = ctk.CTkEntry(self.innkpp_frame, placeholder_text="КПП", width=200)
        self.customer_kpp_entry.pack(side="left")
        self._bind_entry_keys(self.customer_kpp_entry)

        # --- БИК (с автозаполнением банка) ---
        lbl_bik = ctk.CTkLabel(f, text="БИК:")
        self.contract_widgets.append(("label", lbl_bik, 19))

        self.bik_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.bik_frame, 20))

        self.customer_bik_entry = ctk.CTkEntry(
            self.bik_frame, placeholder_text="Введите БИК (9 цифр)", width=200)
        self.customer_bik_entry.pack(side="left", padx=(0, 8))
        self._bind_entry_keys(self.customer_bik_entry)

        self.bik_status_label = ctk.CTkLabel(
            self.bik_frame, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"), width=300, anchor="w")
        self.bik_status_label.pack(side="left")

        # Автозаполнение по БИК при вводе 9 цифр
        self.customer_bik_entry.bind('<KeyRelease>', self._on_bik_keyrelease)

        # --- Банк (автозаполняется по БИК) ---
        lbl_bank = ctk.CTkLabel(f, text="Банк:")
        self.contract_widgets.append(("label", lbl_bank, 21))
        self.customer_bank_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически по БИК", width=500)
        self.contract_widgets.append(("entry", self.customer_bank_entry, 22))
        self._bind_entry_keys(self.customer_bank_entry)

        # --- Р/с (ручной ввод) ---
        lbl_rs = ctk.CTkLabel(f, text="Расчётный счёт (Р/с):")
        self.contract_widgets.append(("label", lbl_rs, 23))
        self.customer_rs_entry = ctk.CTkEntry(
            f, placeholder_text="Введите расчётный счёт", width=350)
        self.contract_widgets.append(("entry", self.customer_rs_entry, 24))
        self._bind_entry_keys(self.customer_rs_entry)

        # --- К/с (ручной ввод) ---
        lbl_ks = ctk.CTkLabel(f, text="Корреспондентский счёт (К/с):")
        self.contract_widgets.append(("label", lbl_ks, 25))
        self.customer_ks_entry = ctk.CTkEntry(
            f, placeholder_text="Введите корреспондентский счёт", width=350)
        self.contract_widgets.append(("entry", self.customer_ks_entry, 26))
        self._bind_entry_keys(self.customer_ks_entry)

        # --- Телефон (авто) ---
        lbl_phone = ctk.CTkLabel(f, text="Телефон:")
        self.contract_widgets.append(("label", lbl_phone, 27))
        self.customer_phone_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически", width=300)
        self.contract_widgets.append(("entry", self.customer_phone_entry, 28))
        self._bind_entry_keys(self.customer_phone_entry)

        # --- Email (авто) ---
        lbl_email = ctk.CTkLabel(f, text="E-mail:")
        self.contract_widgets.append(("label", lbl_email, 29))
        self.customer_email_entry = ctk.CTkEntry(
            f, placeholder_text="Заполняется автоматически", width=350)
        self.contract_widgets.append(("entry", self.customer_email_entry, 30))
        self._bind_entry_keys(self.customer_email_entry)

        # --- Руководитель (авто) ---
        lbl_dir = ctk.CTkLabel(f, text="Должность и ФИО руководителя:")
        self.contract_widgets.append(("label", lbl_dir, 31))

        self.director_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.director_frame, 32))

        self.customer_director_title_entry = ctk.CTkEntry(
            self.director_frame, placeholder_text="Должность (напр. директора)", width=280)
        self.customer_director_title_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.customer_director_title_entry)

        self.customer_director_name_entry = ctk.CTkEntry(
            self.director_frame, placeholder_text="ФИО (напр. Иванов Иван Иванович)", width=320)
        self.customer_director_name_entry.pack(side="left")
        self._bind_entry_keys(self.customer_director_name_entry)

        # --- Основание ---
        lbl_basis = ctk.CTkLabel(f, text="Действует на основании:")
        self.contract_widgets.append(("label", lbl_basis, 33))
        self.customer_basis_entry = ctk.CTkEntry(
            f, placeholder_text="Устава / доверенности №...", width=400)
        self.contract_widgets.append(("entry", self.customer_basis_entry, 34))
        self._bind_entry_keys(self.customer_basis_entry)

        # --- Адрес проведения работ ---
        self.include_work_address = ctk.BooleanVar(value=False)
        self.work_address_cb = ctk.CTkCheckBox(
            f,
            text="Добавить адрес проведения работ",
            variable=self.include_work_address,
            command=self.toggle_work_address_fields
        )
        self.contract_widgets.append(("check", self.work_address_cb, 35))

        self.work_address_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame_work_address", self.work_address_frame, 36))
        ctk.CTkLabel(self.work_address_frame, text="Адрес проведения работ:").pack(
            side="left", padx=(0, 8)
        )
        self.work_address_entry = ctk.CTkEntry(self.work_address_frame, width=480)
        self.work_address_entry.pack(side="left")
        self._bind_entry_keys(self.work_address_entry)

        # --- Разделитель: Смета ---
        sep2 = ctk.CTkLabel(f, text="─── Смета (КП) ─────────────────────────────────────",
                             font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"))
        self.contract_widgets.append(("label_sep", sep2, 37))

        # --- Размер аванса (%) ---
        lbl_advance = ctk.CTkLabel(f, text="Размер аванса (%):",
                                    font=ctk.CTkFont(size=12))
        self.contract_widgets.append(("label", lbl_advance, 38))
        self.advance_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.advance_frame, 39))
        self.advance_percent_entry = ctk.CTkEntry(
            self.advance_frame, placeholder_text="30", width=80)
        self.advance_percent_entry.insert(0, "30")
        self.advance_percent_entry.pack(side="left", padx=(0, 8))
        self._bind_entry_keys(self.advance_percent_entry)
        ctk.CTkLabel(self.advance_frame, text="%  (от суммы договора)",
                     font=ctk.CTkFont(size=11), text_color=("gray50", "gray60")).pack(side="left")

        # --- Загрузка КП ---
        lbl_kp_file = ctk.CTkLabel(f, text="Файл КП (Word .docx) для вставки сметы:")
        self.contract_widgets.append(("label", lbl_kp_file, 40))

        self.kp_file_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.contract_widgets.append(("frame", self.kp_file_frame, 41))

        self.kp_file_var = ctk.StringVar(value="")
        self.kp_file_entry = ctk.CTkEntry(
            self.kp_file_frame, textvariable=self.kp_file_var, width=480)
        self.kp_file_entry.pack(side="left", padx=(0, 10))
        self._bind_entry_keys(self.kp_file_entry)
        ctk.CTkButton(self.kp_file_frame, text="Выбрать файл", width=120,
                      command=self.browse_kp_file).pack(side="left")

        # --- Статус файла КП ---
        self.kp_file_status = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=11),
                                            text_color=("gray50", "gray60"))
        self.contract_widgets.append(("label", self.kp_file_status, 42))

    def _show_contract_content(self):
        """Отображает виджеты режима Договор."""
        for kind, widget, row in self.contract_widgets:
            if kind == "label":
                pady = (20, 0) if row == 0 else (8, 0)
                widget.grid(row=row, column=0, padx=20, pady=pady, sticky="w")
            elif kind == "label_sep":
                widget.grid(row=row, column=0, padx=20, pady=(15, 2), sticky="w")
            elif kind == "entry":
                widget.grid(row=row, column=0, padx=20, pady=(0, 5), sticky="w")
            elif kind == "frame":
                widget.grid(row=row, column=0, padx=20, pady=(0, 5), sticky="ew")
            elif kind == "check":
                widget.grid(row=row, column=0, padx=20, pady=(6, 4), sticky="w")
            elif kind == "frame_work_address":
                if self.include_work_address.get():
                    widget.grid(row=row, column=0, padx=20, pady=(0, 8), sticky="ew")

    # -----------------------------------------------------------------------
    # Привязка клавиш к полям ввода
    # -----------------------------------------------------------------------

    def _bind_entry_keys(self, entry_widget):
        target_widgets = [entry_widget]
        inner_entry = getattr(entry_widget, "_entry", None)
        # Для CTkEntry фактический фокус часто на внутреннем tk.Entry.
        if inner_entry is not None:
            target_widgets.append(inner_entry)

        # Вставка из буфера обмена (Ctrl+V)
        for target in target_widgets:
            target.bind('<Control-v>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            target.bind('<Control-V>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            target.bind('<Control-KeyPress-v>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            target.bind('<<Paste>>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            target.bind('<Control-Insert>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            target.bind('<Shift-Insert>',
                lambda e, w=target: (self._paste_to_entry(w), 'break')[1])
            # Выделить всё (Ctrl+A)
            target.bind('<Control-a>',
                lambda e, w=target: (self._select_all(w), 'break')[1])
            target.bind('<Control-A>',
                lambda e, w=target: (self._select_all(w), 'break')[1])
            # Копировать (Ctrl+C) - разрешаем по умолчанию
            target.bind('<Control-c>', lambda e: None)
            target.bind('<Control-C>', lambda e: None)
            # Вырезать (Ctrl+X) - разрешаем по умолчанию
            target.bind('<Control-x>', lambda e: None)
            target.bind('<Control-X>', lambda e: None)
            # Контекстное меню по правой кнопке
            target.bind('<Button-3>',
                lambda e, w=target: self._show_context_menu(e, w))

    def _setup_global_shortcuts(self):
        """Глобальные горячие клавиши для полей ввода (в т.ч. при русской раскладке)."""
        self.bind_all('<Control-KeyPress>', self._handle_global_ctrl_key, add='+')

    def _handle_global_ctrl_key(self, event):
        focused = self.focus_get()
        if focused is None:
            return

        is_text_entry = all(
            hasattr(focused, method_name) for method_name in ("delete", "insert", "index")
        )
        if not is_text_entry:
            return

        key = (event.keysym or "").lower()
        # keycode 86 — физическая клавиша V на Windows;
        # 'м' — русская раскладка для той же клавиши.
        is_paste = event.keycode == 86 or key in ("v", "м")
        if is_paste:
            return self._paste_to_entry(focused)

    def _show_context_menu(self, event, entry_widget):
        """Shows right-click context menu with Copy/Paste/Cut/Select All."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Вырезать",
            command=lambda: entry_widget.event_generate('<<Cut>>'))
        menu.add_command(label="Копировать",
            command=lambda: entry_widget.event_generate('<<Copy>>'))
        menu.add_command(label="Вставить",
            command=lambda: self._paste_to_entry(entry_widget))
        menu.add_separator()
        menu.add_command(label="Выделить всё",
            command=lambda: self._select_all(entry_widget))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _paste_to_entry(self, entry_widget):
        try:
            text = self.clipboard_get()
            # Убираем переносы строк и лишние пробелы
            text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').strip()
            try:
                entry_widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception:
                pass
            try:
                idx = entry_widget.index(tk.INSERT)
                entry_widget.insert(idx, text)
            except Exception:
                entry_widget.insert(tk.END, text)
        except Exception:
            pass
        return 'break'

    def _select_all(self, entry_widget):
        try:
            entry_widget.select_range(0, tk.END)
            entry_widget.focus_set()
        except Exception:
            pass
        return 'break'

    def _setup_keyboard_navigation(self):
        self.after(200, self._collect_and_bind_focusable)

    def _collect_and_bind_focusable(self):
        self._focusable_widgets = [
            self.kp_name_entry,
            self.kp_title_entry,
            self.save_path_entry,
            self.pir_count_entry,
            self.pir_price_entry,
            self.no_smr_wells_count,
            self.wells_count,
            self.wells_design,
            self.wells_depth,
            self.wells_price_per_meter,
            self.wells_price,
            self.pump_price,
            self.bmz_size,
            self.bmz_price,
        ]
        for i, widget in enumerate(self._focusable_widgets):
            widget.bind('<Down>', lambda e, idx=i: self._focus_next(idx))
            widget.bind('<Up>', lambda e, idx=i: self._focus_prev(idx))

    def _focus_next(self, current_idx):
        for i in range(current_idx + 1, len(self._focusable_widgets)):
            w = self._focusable_widgets[i]
            if w.winfo_ismapped() and str(w.cget('state')) != 'disabled':
                w.focus_set()
                break
        return 'break'

    def _focus_prev(self, current_idx):
        for i in range(current_idx - 1, -1, -1):
            w = self._focusable_widgets[i]
            if w.winfo_ismapped() and str(w.cget('state')) != 'disabled':
                w.focus_set()
                break
        return 'break'

    # -----------------------------------------------------------------------
    # Обработчики событий UI
    # -----------------------------------------------------------------------

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.save_path_var.set(folder)

    def browse_contract_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.contract_save_var.set(folder)

    def browse_kp_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл КП",
            filetypes=[("Word документы", "*.docx"), ("Все файлы", "*.*")]
        )
        if filepath:
            self.kp_file_var.set(filepath)
            self.kp_file_status.configure(
                text=f"Файл выбран: {os.path.basename(filepath)}",
                text_color=("green", "lightgreen")
            )

    def on_branch_change(self, choice):
        if choice == "хоз.пит":
            self.volume_menu.configure(
                values=["до 100", "100-500", "500+", "500+ с переоценкой запасов"]
            )
        else:
            self.volume_menu.configure(values=["до 100", "100-500", "500+"])
            if self.volume_var.get() == "500+ с переоценкой запасов":
                self.volume_var.set("500+")

    def setup_smr_fields(self):
        """Создаёт поля для режима с СМР внутри smr_frame."""
        self.include_wells = ctk.BooleanVar(value=True)
        self.include_pump = ctk.BooleanVar(value=True)
        self.include_bmz = ctk.BooleanVar(value=True)

        # Скважины
        self.wells_cb = ctk.CTkCheckBox(self.smr_frame, text="Скважины", variable=self.include_wells)
        self.wells_cb.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(self.smr_frame, text="Кол-во:").grid(row=0, column=1, padx=5, pady=5, sticky="e")
        self.wells_count = ctk.CTkEntry(self.smr_frame, width=60)
        self.wells_count.insert(0, "1")
        self.wells_count.grid(row=0, column=2, padx=5, pady=5)
        self._bind_entry_keys(self.wells_count)

        ctk.CTkLabel(self.smr_frame, text="Конструктив:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.wells_design = ctk.CTkEntry(self.smr_frame, width=280)
        self.wells_design.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.wells_design)

        ctk.CTkLabel(self.smr_frame, text="Глубина (м):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.wells_depth_var = ctk.StringVar(value="")
        self.wells_depth = ctk.CTkEntry(self.smr_frame, textvariable=self.wells_depth_var, width=100)
        self.wells_depth.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.wells_depth)

        ctk.CTkLabel(self.smr_frame, text="Цена за м:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.wells_price_per_meter_var = ctk.StringVar(value="")
        self.wells_price_per_meter = ctk.CTkEntry(
            self.smr_frame, textvariable=self.wells_price_per_meter_var, width=140
        )
        self.wells_price_per_meter.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.wells_price_per_meter)

        ctk.CTkLabel(self.smr_frame, text="Цена за 1 скв:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.wells_price = ctk.CTkEntry(self.smr_frame, width=140)
        self.wells_price.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.wells_price)
        self.wells_depth_var.trace_add("write", self._on_well_params_change)
        self.wells_price_per_meter_var.trace_add("write", self._on_well_params_change)

        # Насосное оборудование
        self.pump_cb = ctk.CTkCheckBox(self.smr_frame, text="Насосное оборудование", variable=self.include_pump)
        self.pump_cb.grid(row=5, column=0, columnspan=2, padx=5, pady=(15, 5), sticky="w")
        ctk.CTkLabel(self.smr_frame, text="Цена за 1 компл:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
        self.pump_price = ctk.CTkEntry(self.smr_frame, width=140)
        self.pump_price.grid(row=6, column=1, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.pump_price)

        # БМЗ
        self.bmz_cb = ctk.CTkCheckBox(self.smr_frame, text="БМЗ", variable=self.include_bmz)
        self.bmz_cb.grid(row=7, column=0, padx=5, pady=(15, 5), sticky="w")
        ctk.CTkLabel(self.smr_frame, text="Размеры:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.bmz_size = ctk.CTkEntry(self.smr_frame, width=280)
        self.bmz_size.grid(row=8, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.bmz_size)
        ctk.CTkLabel(self.smr_frame, text="Стоимость:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        self.bmz_price = ctk.CTkEntry(self.smr_frame, width=140)
        self.bmz_price.grid(row=9, column=1, padx=5, pady=5, sticky="w")
        self._bind_entry_keys(self.bmz_price)

    def _to_float(self, value: str):
        if value is None:
            return None
        cleaned = str(value).strip().replace(" ", "").replace(",", ".")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    def _on_well_params_change(self, *_):
        """
        Автоматически рассчитывает цену за 1 скважину по формуле:
        глубина * цена за метр.
        Поле остаётся редактируемым вручную.
        """
        depth = self._to_float(self.wells_depth_var.get())
        price_per_meter = self._to_float(self.wells_price_per_meter_var.get())
        if depth is None or price_per_meter is None:
            return

        total = depth * price_per_meter
        total_text = str(int(round(total)))
        self.wells_price.delete(0, tk.END)
        self.wells_price.insert(0, total_text)

    def toggle_smr_fields(self, choice):
        """Показывает/скрывает поля в зависимости от типа работ."""
        if choice == "с смр":
            self.no_smr_frame.grid_forget()
            self.smr_frame.grid(row=16, column=0, padx=20, pady=(0, 15), sticky="nsew",
                                in_=self.scroll_frame)
        else:
            self.smr_frame.grid_forget()
            self.no_smr_frame.grid(row=15, column=0, padx=20, pady=(0, 10), sticky="w",
                                   in_=self.scroll_frame)

    def toggle_pir_fields(self):
        """Показывает/скрывает поля ПИР."""
        if self.include_pir.get():
            self.pir_frame.grid(row=14, column=0, padx=20, pady=(0, 10), sticky="w",
                                in_=self.scroll_frame)
        else:
            self.pir_frame.grid_forget()

    def toggle_work_address_fields(self):
        """Показывает/скрывает поле адреса проведения работ в режиме договора."""
        if self.include_work_address.get():
            self.work_address_frame.grid(row=36, column=0, padx=20, pady=(0, 8), sticky="ew",
                                         in_=self.scroll_frame)
        else:
            self.work_address_frame.grid_forget()

    # -----------------------------------------------------------------------
    # Получение данных по ИНН
    # -----------------------------------------------------------------------

    def fetch_company_data(self):
        inn = self.inn_entry.get().strip()
        if not inn:
            messagebox.showwarning("Предупреждение", "Введите ИНН заказчика.")
            return

        self.inn_status_label.configure(
            text="Загрузка данных...", text_color=("gray50", "gray60"))
        self.fetch_btn.configure(state="disabled")
        self.update()

        def do_fetch():
            try:
                import urllib.request
                import json
                url = f"https://api.checko.ru/v2/company?key=P9ZVrxKxVktK9fsx&inn={inn}"
                with urllib.request.urlopen(url, timeout=15) as resp:
                    data = json.loads(resp.read().decode("utf-8"))

                if "data" not in data:
                    self.after(0, lambda: self._on_fetch_error("Данные не найдены. Проверьте ИНН."))
                    return

                d = data["data"]
                self.after(0, lambda: self._fill_company_fields(d))

            except Exception as e:
                self.after(0, lambda: self._on_fetch_error(f"Ошибка запроса: {str(e)}"))

        threading.Thread(target=do_fetch, daemon=True).start()

    def _fill_company_fields(self, d):
        """Заполняет поля данными из API."""
        try:
            # Наименования
            full_name = d.get("НаимПолн", "")
            full_name_fmt = self._format_company_name(full_name, d)
            short_name = d.get("НаимСокр", "")
            short_name_fmt = self._format_short_name(short_name)

            self._set_entry(self.customer_fullname_entry, full_name_fmt)
            self._set_entry(self.customer_shortname_entry, short_name_fmt)

            # Адрес
            addr = ""
            if "ЮрАдрес" in d and d["ЮрАдрес"]:
                addr = d["ЮрАдрес"].get("АдресРФ", "")
            self._set_entry(self.customer_address_entry, addr)

            # ОГРН, ИНН, КПП
            self._set_entry(self.customer_ogrn_entry, d.get("ОГРН", ""))
            self._set_entry(self.customer_inn_entry, d.get("ИНН", ""))
            self._set_entry(self.customer_kpp_entry, d.get("КПП", ""))

            # Контакты
            contacts = d.get("Контакты", {}) or {}
            phones = contacts.get("Тел", []) or []
            emails = contacts.get("Емэйл", []) or []
            phone_str = ", ".join(phones[:2]) if phones else ""
            email_str = emails[0] if emails else ""
            self._set_entry(self.customer_phone_entry, phone_str)
            self._set_entry(self.customer_email_entry, email_str)

            # Руководитель
            rukovod = d.get("Руковод", []) or []
            if rukovod:
                r = rukovod[0]
                fio = r.get("ФИО", "")
                dolzhn_raw = r.get("НаимДолжн", "")
                # Приводим должность к нижнему регистру для использования в тексте
                dolzhn = dolzhn_raw.lower() if dolzhn_raw else ""
                self._set_entry(self.customer_director_title_entry, dolzhn)
                self._set_entry(self.customer_director_name_entry, fio)

                # Определяем основание
                okopf_code = ""
                if "ОКОПФ" in d and d["ОКОПФ"]:
                    okopf_code = str(d["ОКОПФ"].get("Код", ""))
                # ИП: коды ОКОПФ начинаются с 501xx
                if okopf_code.startswith("501"):
                    basis = "свидетельства о государственной регистрации"
                else:
                    basis = "Устава"
                self._set_entry(self.customer_basis_entry, basis)

            self.inn_status_label.configure(
                text=f"Данные загружены: {short_name_fmt}",
                text_color=("green", "lightgreen")
            )

        except Exception as e:
            self._on_fetch_error(f"Ошибка обработки данных: {str(e)}")
        finally:
            self.fetch_btn.configure(state="normal")

    def _format_company_name(self, name, d):
        """Форматирует полное наименование организации."""
        if not name:
            return name
        # Слова, которые должны писаться строчными (предлоги, союзы) вне кавычек
        LOWERCASE_WORDS = {'с', 'о', 'в', 'и', 'а', 'на', 'по', 'из', 'до', 'за', 'от', 'об',
                           'при', 'для', 'под', 'над', 'без', 'про', 'или', 'но', 'да', 'не'}
        # Если имя в ALL CAPS - конвертируем
        if name == name.upper():
            # Разбиваем по двойным кавычкам
            parts = re.split(r'(")', name)
            result = []
            in_quotes = False
            for part in parts:
                if part == '"':
                    if not in_quotes:
                        result.append('«')
                    else:
                        result.append('»')
                    in_quotes = not in_quotes
                elif in_quotes:
                    # Внутри кавычек - Title Case
                    result.append(part.title())
                else:
                    # Вне кавычек - первое слово с заглавной, предлоги/союзы строчными
                    words = part.split()
                    formatted = []
                    for wi, w in enumerate(words):
                        w_lower = w.lower()
                        if wi == 0:
                            # Первое слово всегда с заглавной
                            formatted.append(w.capitalize())
                        elif w_lower in LOWERCASE_WORDS:
                            # Предлоги/союзы строчными
                            formatted.append(w_lower)
                        elif len(w) <= 3 and w.isalpha() and w.isupper():
                            formatted.append(w)  # аббревиатуры оставляем
                        else:
                            formatted.append(w.capitalize())
                    result.append(" ".join(formatted))
            return "".join(result)
        # Уже в нормальном регистре — исправляем неправильные заглавные буквы в предлогах
        name = re.sub(
            r'Общество С Ограниченной Ответственностью',
            'Общество с ограниченной ответственностью',
            name
        )
        return name

    def _format_short_name(self, name):
        """Форматирует краткое наименование."""
        if not name:
            return name
        # Заменяем двойные кавычки на ёлочки
        name = name.replace('"', '«', 1).replace('"', '»', 1)
        # Если часть в кавычках в верхнем регистре - конвертируем в Title Case
        parts = re.split(r'([«»])', name)
        result = []
        in_quotes = False
        for part in parts:
            if part in ('«', '»'):
                if part == '«':
                    in_quotes = True
                else:
                    in_quotes = False
                result.append(part)
            elif in_quotes and part == part.upper() and any(c.isalpha() for c in part):
                result.append(part.title())
            else:
                result.append(part)
        return ''.join(result)

    def _set_entry(self, entry, value):
        """Устанавливает значение в поле ввода."""
        entry.delete(0, tk.END)
        entry.insert(0, str(value) if value else "")

    def _on_fetch_error(self, msg):
        self.inn_status_label.configure(
            text=f"Ошибка: {msg}", text_color=("red", "salmon"))
        self.fetch_btn.configure(state="normal")

    # -----------------------------------------------------------------------
    # Автозаполнение банка по БИК
    # -----------------------------------------------------------------------

    def _on_bik_keyrelease(self, event=None):
        """Triggered on each keypress in BIK field; fetches bank when 9 digits entered."""
        bik = self.customer_bik_entry.get().strip()
        if len(bik) == 9 and bik.isdigit():
            self.bik_status_label.configure(
                text="Поиск...", text_color=("gray50", "gray60"))
            threading.Thread(
                target=self._fetch_bank_by_bik, args=(bik,), daemon=True
            ).start()
        elif len(bik) < 9:
            self.bik_status_label.configure(text="", text_color=("gray50", "gray60"))

    def _fetch_bank_by_bik(self, bik: str):
        """Получает название банка по БИК через API Checko."""
        try:
            import urllib.request
            import json

            url = f"https://api.checko.ru/v2/bank?key=P9ZVrxKxVktK9fsx&bic={bik}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode('utf-8'))

            if 'data' not in data or not data['data']:
                self.after(0, lambda: self.bik_status_label.configure(
                    text="Банк не найден", text_color=("red", "salmon")))
                return

            d = data['data']

            # Название банка — ключ 'Наим' в ответе Checko
            bank_name_raw = (
                d.get('Наим', '') or
                d.get('НаимКред', '') or
                d.get('НаимПолн', '') or
                d.get('Наименование', '') or ''
            )

            # Приводим к нормальному регистру
            def title_bank(s):
                if not s:
                    return s
                # Заменяем двойные кавычки на ёлочки
                s = re.sub(r'"([^"]+)"', lambda m: '«' + m.group(1) + '»', s)
                # Title case с сохранением ёлочек
                parts = re.split(r'(«[^»]+»)', s)
                result = []
                for part in parts:
                    if part.startswith('«'):
                        result.append(part)
                    else:
                        words = part.split()
                        titled = []
                        for wi, w in enumerate(words):
                            if wi == 0 or w.lower() not in ('в', 'на', 'по', 'для', 'и', 'или', 'с', 'о', 'об', 'от', 'до', 'из', 'за', 'при', 'под', 'над', 'без', 'через'):
                                titled.append(w.capitalize())
                            else:
                                titled.append(w.lower())
                        result.append(' '.join(titled))
                return ''.join(result)

            bank_name = title_bank(bank_name_raw) if bank_name_raw == bank_name_raw.upper() else bank_name_raw

            # Кор. счёт — в ответе Checko это объект {'Номер': '...', 'Дата': '...'}
            ks_raw = d.get('КорСчет', '') or d.get('КС', '') or ''
            if isinstance(ks_raw, dict):
                ks_num = ks_raw.get('Номер', '') or ''
            elif isinstance(ks_raw, list) and ks_raw:
                first = ks_raw[0]
                ks_num = first.get('Номер', '') if isinstance(first, dict) else str(first)
            else:
                ks_num = str(ks_raw) if ks_raw else ''

            def update_ui():
                self._set_entry(self.customer_bank_entry, bank_name)
                if ks_num and not self.customer_ks_entry.get().strip():
                    self._set_entry(self.customer_ks_entry, ks_num)
                self.bik_status_label.configure(
                    text=f"✓ {bank_name[:40]}",
                    text_color=("green", "lightgreen"))

            self.after(0, update_ui)

        except Exception as e:
            self.after(0, lambda: self.bik_status_label.configure(
                text=f"Ошибка: {str(e)[:50]}",
                text_color=("red", "salmon")))

    # -----------------------------------------------------------------------
    # Главная кнопка генерации
    # -----------------------------------------------------------------------

    def on_generate(self):
        if self.current_mode == "kp":
            self.generate_kp()
        else:
            self.generate_contract()

    # -----------------------------------------------------------------------
    # Генерация КП
    # -----------------------------------------------------------------------

    def generate_kp(self):
        smr_type = self.smr_var.get()

        data = {
            "kp_name": self.kp_name_entry.get().strip(),
            "kp_title": self.kp_title_entry.get().strip(),
            "branch": self.branch_var.get(),
            "volume": self.volume_var.get(),
            "smr_type": smr_type,
            "include_wells": self.include_wells.get() if smr_type == "с смр" else False,
            "include_pump": self.include_pump.get() if smr_type == "с смр" else False,
            "include_bmz": self.include_bmz.get() if smr_type == "с смр" else False,
            "include_pir": self.include_pir.get(),
            "pir_count": self.pir_count_entry.get().strip() or "1",
            "pir_price": self.pir_price_entry.get().strip(),
            "save_dir": self.save_path_var.get().strip()
        }

        if smr_type == "с смр":
            data.update({
                "wells_count": self.wells_count.get().strip() or "1",
                "wells_design": self.wells_design.get().strip(),
                "wells_depth": self.wells_depth.get().strip(),
                "wells_price": self.wells_price.get().strip(),
                "pump_price": self.pump_price.get().strip(),
                "bmz_size": self.bmz_size.get().strip(),
                "bmz_price": self.bmz_price.get().strip()
            })
        else:
            data["wells_count"] = self.no_smr_wells_count.get().strip() or "1"

        if not data["kp_name"]:
            messagebox.showwarning("Предупреждение", "Введите название КП.")
            self.kp_name_entry.focus_set()
            return

        if not data["kp_title"]:
            messagebox.showwarning("Предупреждение", "Введите текст для титульного листа КП.")
            self.kp_title_entry.focus_set()
            return

        if data["include_pir"] and not data["pir_price"]:
            messagebox.showwarning("Предупреждение", "Укажите цену для ПИР.")
            self.pir_price_entry.focus_set()
            return

        if not data["save_dir"]:
            messagebox.showwarning("Предупреждение", "Укажите путь для сохранения.")
            return

        generator = KPGenerator()
        try:
            self.generate_button.configure(state="disabled", text="Генерация...")
            self.update()

            docx_path, pdf_path = generator.create_kp(data)

            msg = f"КП успешно создано!\n\nWord: {docx_path}"
            if pdf_path and os.path.exists(pdf_path):
                msg += f"\nPDF: {pdf_path}"
            else:
                msg += "\n\nPDF не создан (требуется LibreOffice или Microsoft Word)."

            messagebox.showinfo("Успех", msg)

        except PermissionError:
            messagebox.showerror(
                "Ошибка доступа",
                "Не удалось сохранить файл.\n"
                "Возможно, файл уже открыт в Word, или нет прав на запись в эту папку."
            )
        except FileNotFoundError as e:
            messagebox.showerror("Ошибка", f"Шаблон не найден:\n{e}")
        except Exception as e:
            error_msg = f"Ошибка при генерации КП: {str(e)}\n{traceback.format_exc()}"
            log_error(error_msg)
            messagebox.showerror("Ошибка", "Не удалось создать КП.\nПодробности в файле error_log.txt")
        finally:
            self.generate_button.configure(state="normal", text="Сгенерировать КП")

    # -----------------------------------------------------------------------
    # Генерация Договора
    # -----------------------------------------------------------------------

    def generate_contract(self):
        contract_num = self.contract_number_entry.get().strip()
        save_dir = self.contract_save_var.get().strip()

        if not contract_num:
            messagebox.showwarning("Предупреждение", "Введите номер договора.")
            self.contract_number_entry.focus_set()
            return

        if not save_dir:
            messagebox.showwarning("Предупреждение", "Укажите путь для сохранения.")
            return

        customer_fullname = self.customer_fullname_entry.get().strip()
        if not customer_fullname:
            messagebox.showwarning("Предупреждение",
                                   "Введите ИНН и получите данные заказчика, "
                                   "или заполните наименование вручную.")
            return

        data = {
            "contract_number": contract_num,
            "save_dir": save_dir,
            "customer_fullname": customer_fullname,
            "customer_shortname": self.customer_shortname_entry.get().strip(),
            "customer_address": self.customer_address_entry.get().strip(),
            "customer_ogrn": self.customer_ogrn_entry.get().strip(),
            "customer_inn": self.customer_inn_entry.get().strip(),
            "customer_kpp": self.customer_kpp_entry.get().strip(),
            "customer_bank": self.customer_bank_entry.get().strip(),
            "customer_bik": self.customer_bik_entry.get().strip(),
            "customer_rs": self.customer_rs_entry.get().strip(),
            "customer_ks": self.customer_ks_entry.get().strip(),
            "customer_phone": self.customer_phone_entry.get().strip(),
            "customer_email": self.customer_email_entry.get().strip(),
            "customer_director_title": self.customer_director_title_entry.get().strip(),
            "customer_director_name": self.customer_director_name_entry.get().strip(),
            "customer_basis": self.customer_basis_entry.get().strip(),
            "include_work_address": self.include_work_address.get(),
            "work_address": self.work_address_entry.get().strip(),
            "kp_file": self.kp_file_var.get().strip(),
            "advance_percent": self.advance_percent_entry.get().strip() or "30",
        }

        if data["include_work_address"] and not data["work_address"]:
            messagebox.showwarning("Предупреждение", "Введите адрес проведения работ.")
            self.work_address_entry.focus_set()
            return

        gen = ContractGenerator()
        try:
            self.generate_button.configure(state="disabled", text="Генерация...")
            self.update()

            docx_path = gen.create_contract(data)

            msg = f"Договор успешно создан!\n\nWord: {docx_path}"
            messagebox.showinfo("Успех", msg)

        except PermissionError:
            messagebox.showerror(
                "Ошибка доступа",
                "Не удалось сохранить файл.\n"
                "Возможно, файл уже открыт в Word, или нет прав на запись в эту папку."
            )
        except FileNotFoundError as e:
            messagebox.showerror("Ошибка", f"Файл не найден:\n{e}")
        except Exception as e:
            error_msg = f"Ошибка при генерации договора: {str(e)}\n{traceback.format_exc()}"
            log_error(error_msg)
            messagebox.showerror("Ошибка",
                                 f"Не удалось создать договор.\n{str(e)}\n"
                                 "Подробности в error_log.txt")
        finally:
            self.generate_button.configure(state="normal", text="Сгенерировать Договор")


if __name__ == "__main__":
    try:
        app = KPApp()
        app.mainloop()
    except Exception as e:
        error_msg = f"Критическая ошибка: {str(e)}\n{traceback.format_exc()}"
        log_error(error_msg)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Критическая ошибка",
                             "Приложение аварийно завершилось.\nПодробности в error_log.txt")
