#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конвертер расписания: графическая оболочка (tkinter).

Запуск:  python gui.py          (или двойным щелчком по «Запуск GUI.bat»)

Логика полностью находится в пакете converter/, окно — только оболочка,
поэтому тот же сценарий доступен из командной строки: python -m converter.cli
"""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
import warnings

warnings.filterwarnings("ignore")

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# при сборке в .exe файлы настроек и соответствий должны лежать рядом с программой
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from converter import core                                    # noqa: E402
from converter import reference as reference_mod               # noqa: E402
from converter.aliases import Aliases                          # noqa: E402
from converter.report import KIND_TITLES                       # noqa: E402

APP_TITLE = "Конвертер расписания: Excel → XML (TimeTableExchange)"
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
SECTION_NAMES = {"subjects": "предмет", "rooms": "кабинет", "classes": "класс", "days": "день",
                 "teachers": "учитель"}
IGNORE_LABEL = "— пропускать (не включать в XML) —"


class App(tk.Tk):
    # ------------------------------------------------------------------ #
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x780")
        self.minsize(940, 660)
        self.ref = None
        self._pending = None
        self._dict_values = {"subjects": [IGNORE_LABEL], "rooms": [IGNORE_LABEL],
                             "classes": [IGNORE_LABEL], "days": [IGNORE_LABEL]}
        self._current_item = None
        self.aliases = Aliases()
        self.result = None
        self.unresolved = []
        self.settings = self._load_settings()

        self.var_excel = tk.StringVar()
        self.var_reference = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_aliases = tk.StringVar(value=os.path.join(BASE_DIR, "aliases.json"))
        self.var_teachers = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_week = tk.StringVar()
        self.var_encoding = tk.StringVar(value="windows-1251")
        self.var_room_mode = tk.StringVar(value="skip")
        self.var_skip_subject = tk.BooleanVar(value=True)
        self.var_align = tk.BooleanVar(value=True)
        self.var_conflicts = tk.BooleanVar(value=True)
        self.var_strict = tk.BooleanVar(value=False)
        self.var_ignore_item = tk.BooleanVar(value=False)
        self.var_rev_xml = tk.StringVar()
        self.var_rev_out = tk.StringVar()
        self.var_rev_teachers = tk.BooleanVar(value=True)
        self.var_rev_summary = tk.BooleanVar(value=True)
        self.var_rule_class = tk.StringVar()
        self.var_target = tk.StringVar()
        self.var_status = tk.StringVar(value="Готово. Выберите файл таблицы и эталонный XML.")

        self._build_menu()
        self._build_ui()
        self._restore_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ #
    # настройки
    # ------------------------------------------------------------------ #
    def _load_settings(self) -> dict:
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _save_settings(self) -> None:
        data = {
            "excel": self.var_excel.get(), "reference": self.var_reference.get(),
            "output": self.var_output.get(), "aliases": self.var_aliases.get(),
            "teachers": self.var_teachers.get(),
            "sheet": self.var_sheet.get(), "week": self.var_week.get(),
            "encoding": self.var_encoding.get(), "room_mode": self.var_room_mode.get(),
            "skip_subject": self.var_skip_subject.get(), "align": self.var_align.get(),
            "conflicts": self.var_conflicts.get(), "strict": self.var_strict.get(),
            "geometry": self.geometry(),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _restore_settings(self) -> None:
        s = self.settings
        self.var_excel.set(s.get("excel", ""))
        self.var_reference.set(s.get("reference", ""))
        self.var_output.set(s.get("output", ""))
        self.var_aliases.set(s.get("aliases", self.var_aliases.get()))
        self.var_teachers.set(s.get("teachers", ""))
        self.var_sheet.set(s.get("sheet", ""))
        self.var_week.set(s.get("week", ""))
        self.var_encoding.set(s.get("encoding", "windows-1251"))
        self.var_room_mode.set(s.get("room_mode", "skip"))
        self.var_skip_subject.set(bool(s.get("skip_subject", True)))
        self.var_align.set(bool(s.get("align", True)))
        self.var_conflicts.set(bool(s.get("conflicts", True)))
        self.var_strict.set(bool(s.get("strict", False)))
        if s.get("geometry"):
            try:
                self.geometry(s["geometry"])
            except tk.TclError:
                pass
        if self.var_aliases.get():
            self.aliases = Aliases.load(self.var_aliases.get())
            try:
                self._aliases_mtime = os.path.getmtime(self.var_aliases.get())
            except OSError:
                pass
            self._refresh_rules()

    def _on_close(self) -> None:
        self._save_settings()
        self.destroy()

    # ------------------------------------------------------------------ #
    # интерфейс
    # ------------------------------------------------------------------ #
    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="Открыть таблицу «по предметам»…", command=self.pick_excel)
        m_file.add_command(label="Открыть таблицу «по учителям»…", command=self.pick_teachers)
        m_file.add_command(label="Открыть эталонный XML…", command=self.pick_reference)
        m_file.add_separator()
        m_file.add_command(label="Выход", command=self._on_close)
        menubar.add_cascade(label="Файл", menu=m_file)

        m_run = tk.Menu(menubar, tearoff=0)
        m_run.add_command(label="Конвертировать", command=self.on_convert, accelerator="F5")
        m_run.add_command(label="Проверить без записи файла", command=self.on_check, accelerator="F6")
        m_run.add_command(label="Сохранить отчёт…", command=self.save_report)
        m_run.add_separator()
        m_run.add_command(label="Проверить целостность эталонного XML", command=self.check_reference)
        m_run.add_separator()
        m_run.add_command(label="XML → Excel (обратное преобразование)…",
                          command=lambda: self.notebook.select(self.tab_reverse))
        menubar.add_cascade(label="Действия", menu=m_run)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="Как это работает", command=self.show_help)
        m_help.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=m_help)
        self.config(menu=menubar)
        self.bind_all("<F5>", lambda e: self.on_convert())
        self.bind_all("<F6>", lambda e: self.on_check())

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        for name in ("clam", "vista", "default"):
            if name in style.theme_names():
                style.theme_use(name)
                break
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(8, 2))
        self.tab_convert = ttk.Frame(self.notebook)
        self.tab_map = ttk.Frame(self.notebook)
        self.tab_report = ttk.Frame(self.notebook)
        self.tab_reverse = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_convert, text="  1. Конвертация  ")
        self.notebook.add(self.tab_map, text="  2. Соответствия  ")
        self.notebook.add(self.tab_report, text="  3. Отчёт  ")
        self.notebook.add(self.tab_reverse, text="  4. XML → Excel  ")

        self._build_tab_convert()
        self._build_tab_map()
        self._build_tab_report()
        self._build_tab_reverse()

        status = ttk.Frame(self)
        status.pack(fill="x", side="bottom")
        ttk.Label(status, textvariable=self.var_status, anchor="w").pack(fill="x", padx=8, pady=3)
        ttk.Separator(status, orient="horizontal").pack(fill="x")

    # ----------------------------- вкладка 1 --------------------------- #
    def _path_row(self, parent, row, label, variable, command, hint=""):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="Обзор…", width=10, command=command).grid(row=row, column=2, padx=6, pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666").grid(row=row, column=3, sticky="w", padx=4)
        return entry

    def _build_tab_convert(self) -> None:
        frame = self.tab_convert
        box_files = ttk.LabelFrame(frame, text=" Файлы ")
        box_files.pack(fill="x", padx=8, pady=6)
        box_files.columnconfigure(1, weight=1)
        self._path_row(box_files, 0, "Таблица «по предметам» (.xlsx):", self.var_excel, self.pick_excel,
                       hint="основной входной файл")
        self._path_row(box_files, 1, "Таблица «по учителям» (.xlsx):", self.var_teachers,
                       self.pick_teachers, hint="необязательно: ФИО для точного выбора групп")
        self._path_row(box_files, 2, "Эталонный XML системы:", self.var_reference, self.pick_reference,
                       hint="ExportCm.xml — источник ID")
        self._path_row(box_files, 3, "Результат (XML):", self.var_output, self.pick_output,
                       hint="по умолчанию — рядом с таблицей")
        self._path_row(box_files, 4, "Файл соответствий:", self.var_aliases, self.pick_aliases,
                       hint="aliases.json")

        row = ttk.Frame(box_files)
        row.grid(row=5, column=0, columnspan=4, sticky="ew", padx=6, pady=(0, 6))
        ttk.Label(row, text="Лист Excel:").pack(side="left")
        self.cmb_sheet = ttk.Combobox(row, textvariable=self.var_sheet, width=18, state="readonly")
        self.cmb_sheet.pack(side="left", padx=6)
        ttk.Button(row, text="Прочитать листы", width=16, command=self.load_sheets).pack(side="left", padx=4)
        ttk.Label(row, text="    Неделя (дата понедельника):").pack(side="left")
        ttk.Entry(row, textvariable=self.var_week, width=12).pack(side="left", padx=6)
        ttk.Label(row, text="дд.мм.гггг, пусто — как в эталоне", foreground="#666").pack(side="left")

        box_opts = ttk.LabelFrame(frame, text=" Параметры преобразования ")
        box_opts.pack(fill="x", padx=8, pady=6)
        ttk.Checkbutton(box_opts, variable=self.var_skip_subject,
                        text="Пропускать уроки с нераспознанным предметом (иначе — остановка с ошибкой)"
                        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(box_opts, variable=self.var_align,
                        text="Подбирать группу (подгруппу) по кабинетам из эталона"
                        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(box_opts, variable=self.var_conflicts,
                        text="Проверять конфликты учителей и кабинетов"
                        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(box_opts, variable=self.var_strict,
                        text="Строгий режим: не сохранять XML, пока есть нераспознанные значения"
                        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=2)

        ttk.Label(box_opts, text="Неизвестный кабинет:").grid(row=4, column=0, sticky="w", padx=6, pady=4)
        radio = ttk.Frame(box_opts)
        radio.grid(row=4, column=1, sticky="w")
        ttk.Radiobutton(radio, text="пропускать урок", variable=self.var_room_mode,
                        value="skip").pack(side="left")
        ttk.Radiobutton(radio, text="ставить урок без кабинета", variable=self.var_room_mode,
                        value="noroom").pack(side="left", padx=8)

        ttk.Label(box_opts, text="Кодировка результата:").grid(row=5, column=0, sticky="w", padx=6, pady=4)
        cmb_enc = ttk.Combobox(box_opts, textvariable=self.var_encoding, width=16,
                               values=["windows-1251", "utf-8"], state="readonly")
        cmb_enc.grid(row=5, column=1, sticky="w", pady=4)
        ttk.Label(box_opts, text="как в эталоне — windows-1251", foreground="#666"
                  ).grid(row=5, column=2, sticky="w", padx=8)

        box_run = ttk.Frame(frame)
        box_run.pack(fill="x", padx=8, pady=4)
        self.btn_convert = ttk.Button(box_run, text="Конвертировать  (F5)", command=self.on_convert)
        self.btn_convert.pack(side="left", padx=4)
        self.btn_check = ttk.Button(box_run, text="Проверить без записи  (F6)", command=self.on_check)
        self.btn_check.pack(side="left", padx=4)
        ttk.Button(box_run, text="Открыть папку результата", command=self.open_output_folder
                   ).pack(side="left", padx=4)
        ttk.Button(box_run, text="Сохранить отчёт…", command=self.save_report).pack(side="left", padx=4)
        ttk.Button(box_run, text="Проверить эталон", command=self.check_reference).pack(side="left", padx=4)
        self.progress = ttk.Progressbar(box_run, mode="indeterminate", length=160)
        self.progress.pack(side="right", padx=6)

        box_log = ttk.LabelFrame(frame, text=" Ход работы ")
        box_log.pack(fill="both", expand=True, padx=8, pady=6)
        self.log = scrolledtext.ScrolledText(box_log, height=14, wrap="word", state="disabled",
                                             font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

    # ----------------------------- вкладка 2 --------------------------- #
    def _build_tab_map(self) -> None:
        frame = self.tab_map
        top = ttk.Frame(frame)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Button(top, text="Загрузить справочники системы", command=self.load_reference_dicts
                   ).pack(side="left", padx=4)
        ttk.Button(top, text="Сохранить соответствия", command=self.save_aliases).pack(side="left", padx=4)
        ttk.Button(top, text="Пересчитать нераспознанное", command=self.on_check).pack(side="left", padx=4)
        self.lbl_map_hint = ttk.Label(top, foreground="#444",
                                      text="Сначала выполните «Проверить без записи» — здесь появится "
                                           "список того, что программа не смогла распознать.")
        self.lbl_map_hint.pack(side="left", padx=12)

        paned = ttk.Panedwindow(frame, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=4)

        left = ttk.Labelframe(paned, text=" Нераспознанные значения ")
        paned.add(left, weight=1)
        self.list_unresolved = tk.Listbox(left, activestyle="dotbox", exportselection=False,
                                          font=("Segoe UI", 10))
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.list_unresolved.yview)
        self.list_unresolved.configure(yscrollcommand=scroll.set)
        self.list_unresolved.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        scroll.pack(side="right", fill="y")
        self.list_unresolved.bind("<<ListboxSelect>>", self.on_unresolved_select)

        right = ttk.Labelframe(paned, text=" Правило соответствия ")
        paned.add(right, weight=2)
        right.columnconfigure(1, weight=1)

        ttk.Label(right, text="Значение в Excel:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.lbl_excel_value = ttk.Label(right, text="—", font=("Segoe UI", 10, "bold"))
        self.lbl_excel_value.grid(row=0, column=1, sticky="w", padx=6)
        self.lbl_problem = ttk.Label(right, text="", foreground="#a33", wraplength=520, justify="left")
        self.lbl_problem.grid(row=1, column=0, columnspan=3, sticky="w", padx=6)

        ttk.Label(right, text="Заменить на:").grid(row=2, column=0, sticky="nw", padx=6, pady=6)
        self.cmb_target = ttk.Combobox(right, textvariable=self.var_target, values=[])
        self.cmb_target.grid(row=2, column=1, sticky="ew", padx=6, pady=6)
        self.cmb_target.bind("<<ComboboxSelected>>", self._on_target_selected)
        ttk.Checkbutton(right, text="Пропускать это значение", variable=self.var_ignore_item,
                        command=self._on_ignore_toggle).grid(row=3, column=1, sticky="w", padx=6)

        ttk.Label(right, text="Только для класса:").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        self.cmb_rule_class = ttk.Combobox(right, textvariable=self.var_rule_class, values=[""],
                                           state="readonly", width=28)
        self.cmb_rule_class.grid(row=4, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(right, text="пусто — правило для всех классов", foreground="#666"
                  ).grid(row=4, column=2, sticky="w", padx=6)

        btns = ttk.Frame(right)
        btns.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=8)
        ttk.Button(btns, text="Назначить правило", command=self.apply_rule).pack(side="left", padx=4)
        ttk.Button(btns, text="Удалить правило", command=self.delete_rule).pack(side="left", padx=4)
        ttk.Button(btns, text="Подставить предложение программы", command=self.use_suggestion
                   ).pack(side="left", padx=4)

        ttk.Separator(right, orient="horizontal").grid(row=6, column=0, columnspan=3, sticky="ew",
                                                       padx=6, pady=6)
        ttk.Label(right, text="Сохранённые правила:").grid(row=7, column=0, columnspan=3,
                                                           sticky="w", padx=6)
        box = ttk.Frame(right)
        box.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=6, pady=4)
        right.rowconfigure(8, weight=1)
        self.tree_rules = ttk.Treeview(box, columns=("section", "excel", "target", "cls"),
                                       show="headings", height=8)
        for col, title, width in (("section", "Раздел", 90), ("excel", "В Excel", 220),
                                  ("target", "В системе", 260), ("cls", "Класс", 120)):
            self.tree_rules.heading(col, text=title)
            self.tree_rules.column(col, width=width, anchor="w")
        s2 = ttk.Scrollbar(box, orient="vertical", command=self.tree_rules.yview)
        self.tree_rules.configure(yscrollcommand=s2.set)
        self.tree_rules.pack(side="left", fill="both", expand=True)
        s2.pack(side="right", fill="y")

    # ----------------------------- вкладка 3 --------------------------- #
    def _build_tab_report(self) -> None:
        frame = self.tab_report
        top = ttk.Frame(frame)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Button(top, text="Сохранить отчёт (.txt)", command=self.save_report).pack(side="left", padx=4)
        ttk.Button(top, text="Нераспознанное (.csv)", command=self.save_unresolved).pack(side="left", padx=4)
        ttk.Button(top, text="Протокол сопоставления (.csv)", command=self.save_mapping
                   ).pack(side="left", padx=4)
        self.txt_report = scrolledtext.ScrolledText(frame, wrap="none", state="disabled",
                                                    font=("Consolas", 9))
        self.txt_report.pack(fill="both", expand=True, padx=8, pady=6)

    # ------------------------------------------------------------------ #
    # выбор файлов
    # ------------------------------------------------------------------ #
    def pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Таблица расписания",
            filetypes=[("Таблицы Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")],
            initialdir=os.path.dirname(self.var_excel.get() or BASE_DIR) or BASE_DIR)
        if path:
            self.var_excel.set(path)
            if not self.var_output.get():
                base = os.path.splitext(os.path.basename(path))[0]
                self.var_output.set(os.path.join(os.path.dirname(path), f"{base}_ExportCm.xml"))
            self.load_sheets(silent=True)

    def pick_teachers(self) -> None:
        path = filedialog.askopenfilename(
            title="Таблица «расписание по учителям» (необязательно)",
            filetypes=[("Таблицы Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")],
            initialdir=os.path.dirname(self.var_teachers.get() or self.var_excel.get() or BASE_DIR)
            or BASE_DIR)
        if path:
            self.var_teachers.set(path)

    def pick_reference(self) -> None:
        path = filedialog.askopenfilename(
            title="Эталонный XML из системы",
            filetypes=[("XML", "*.xml"), ("Все файлы", "*.*")],
            initialdir=os.path.dirname(self.var_reference.get() or BASE_DIR) or BASE_DIR)
        if path:
            self.var_reference.set(path)
            self.ref = None
            self.load_reference_dicts(silent=True)

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить XML", defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("Все файлы", "*.*")],
            initialfile=os.path.basename(self.var_output.get() or "ExportCm_new.xml"))
        if path:
            self.var_output.set(path)

    def pick_aliases(self) -> None:
        path = filedialog.askopenfilename(
            title="Файл соответствий", filetypes=[("JSON", "*.json"), ("Все файлы", "*.*")],
            initialdir=os.path.dirname(self.var_aliases.get() or BASE_DIR) or BASE_DIR)
        if path:
            self.var_aliases.set(path)
            self.aliases = Aliases.load(path)
            try:
                self._aliases_mtime = os.path.getmtime(path)
            except OSError:
                pass
            self._refresh_rules()
            self.write_log(f"Загружен файл соответствий: {path}")

    def load_sheets(self, silent: bool = False) -> None:
        path = self.var_excel.get()
        if not path or not os.path.exists(path):
            if not silent:
                messagebox.showwarning(APP_TITLE, "Сначала выберите файл таблицы Excel.")
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            names = wb.sheetnames
            wb.close()
        except Exception as exc:
            if not silent:
                messagebox.showerror(APP_TITLE, f"Не удалось прочитать книгу Excel:\n{exc}")
            return
        self.cmb_sheet.configure(values=["(первый лист)"] + names)
        if self.var_sheet.get() not in names:
            self.var_sheet.set("")
        if not silent:
            self.write_log(f"Листы книги: {', '.join(names)}")

    # ------------------------------------------------------------------ #
    # справочники и соответствия
    # ------------------------------------------------------------------ #
    def load_reference_dicts(self, silent: bool = False, refresh: bool = True) -> bool:
        path = self.var_reference.get()
        if not path or not os.path.exists(path):
            if not silent:
                messagebox.showwarning(APP_TITLE, "Сначала выберите эталонный XML системы.")
            return False
        if self.ref is not None and getattr(self.ref, "path", None) == path:
            return True
        try:
            self.ref = reference_mod.load(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Не удалось разобрать эталонный XML:\n{exc}")
            return False
        subjects = sorted(set(self.ref.subject_names))
        rooms = sorted(set(self.ref.room_names))
        classes = [""] + self.ref.class_names
        days = [d.name for d in self.ref.days]
        from converter.teachers import TeacherIndex
        self._dict_values = {"subjects": [IGNORE_LABEL] + subjects,
                             "rooms": [IGNORE_LABEL] + rooms,
                             "classes": [IGNORE_LABEL] + classes[1:],
                             "days": [IGNORE_LABEL] + days,
                             "teachers": [IGNORE_LABEL] + TeacherIndex(self.ref).fio_list}
        self.cmb_rule_class.configure(values=classes)
        if refresh:
            self._refresh_rules()   # проставить/снять метки «нет в эталоне»
        if not silent:
            self.write_log(
                f"Справочники загружены: классов {len(self.ref.classes)}, "
                f"предметов {len(self.ref.subjects)}, кабинетов {len(self.ref.rooms)}, "
                f"учителей {len(self.ref.teachers)}, записей плана {len(self.ref.plan)}.")
        return True

    def _refresh_rules(self) -> None:
        for item in self.tree_rules.get_children():
            self.tree_rules.delete(item)
        for section in ("subjects", "rooms", "classes", "days", "teachers"):
            for excel, target in self.aliases.items(section):
                if Aliases.is_ignore(target):
                    text = "пропускать"
                else:
                    text = _short(target)
                    if not self.target_exists(section, str(target)):
                        text += "  ⚠ нет в эталоне"
                self.tree_rules.insert("", "end", values=(SECTION_NAMES.get(section, section),
                                                          excel, text, ""))
        for cls, rules in (self.aliases.data.get("class_subjects") or {}).items():
            if not isinstance(rules, dict):
                continue
            for excel, target in rules.items():
                if Aliases.is_ignore(target):
                    text = "пропускать"
                else:
                    text = str(target)
                    if not self.target_exists("subjects", str(target)):
                        text += "  ⚠ нет в эталоне"
                self.tree_rules.insert("", "end", values=("предмет", excel, text, cls))

    def on_unresolved_select(self, _event=None) -> None:
        sel = self.list_unresolved.curselection()
        if not sel:
            return
        item = self.unresolved[sel[0]]
        self._current_item = item
        self.lbl_excel_value.config(text=f"«{item.value}»  ({item.count} раз)")
        problem = KIND_TITLES.get(item.kind, item.kind)
        if item.classes:
            problem += f"; классы: {item.classes}"
        if item.kind == "subject_not_in_plan":
            problem += (". В этих классах курс может называться в плане иначе — предложение "
                        "программы является гипотезой. Правило-замена действует только там, "
                        "где прямого названия нет в плане, поэтому глобальное правило безопасно.")
        self.lbl_problem.config(text=problem)
        self.var_ignore_item.set(False)
        self.var_rule_class.set("")
        self.load_reference_dicts(silent=True)
        values = self._dict_values.get(item.section, [IGNORE_LABEL]) if self.ref else [IGNORE_LABEL]
        self.cmb_target.configure(values=values)
        self.var_target.set(item.suggestion or "")
        self.cmb_target.configure(state="normal")

    def _on_target_selected(self, _event=None) -> None:
        if self.var_target.get() == IGNORE_LABEL:
            self.var_ignore_item.set(True)

    def _on_ignore_toggle(self) -> None:
        if self.var_ignore_item.get():
            self.var_target.set(IGNORE_LABEL)

    def use_suggestion(self) -> None:
        item = getattr(self, "_current_item", None)
        if item and item.suggestion:
            self.var_target.set(item.suggestion)
            self.var_ignore_item.set(False)

    def apply_rule(self) -> None:
        item = getattr(self, "_current_item", None)
        if item is None:
            messagebox.showinfo(APP_TITLE, "Выберите значение в списке слева.")
            return
        target = self.var_target.get().strip()
        ignore = self.var_ignore_item.get() or target == IGNORE_LABEL or not target
        value = None if ignore else target
        cls_name = self.var_rule_class.get().strip()
        if cls_name and item.section == "subjects":
            section = self.aliases.data.setdefault("class_subjects", {})
            rules = section.setdefault(cls_name, {})
            if value is None:
                rules.pop(item.value, None)
            else:
                rules[item.value] = value
        elif cls_name:
            messagebox.showinfo(APP_TITLE, "Правило «только для класса» поддерживается для предметов.")
            return
        else:
            if value is None:
                self.aliases.set(item.section, item.value, None)
            else:
                self.aliases.set(item.section, item.value, value)
        self._refresh_rules()
        idx = self.list_unresolved.curselection()
        if idx:
            self.list_unresolved.delete(idx[0])
            del self.unresolved[idx[0]]
        self.save_aliases(silent=True)
        if value is not None and not self.target_exists(item.section, value):
            messagebox.showwarning(
                APP_TITLE,
                f"Правило сохранено, но НЕ сработает при конвертации:\nв справочнике эталона "
                f"({SECTION_NAMES.get(item.section, item.section)}) нет значения «{value}».\n\n"
                f"Что делать: добавьте его в системе и выгрузите новый эталонный XML "
                f"(или исправьте название в правиле — список названий доступен в выпадающем списке).")
        self.var_status.set(f"Правило сохранено: «{item.value}» → "
                            f"{'пропускать' if value is None else value}. Нажмите «Пересчитать».")

    def delete_rule(self) -> None:
        sel = self.tree_rules.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Выберите правило в списке.")
            return
        values = self.tree_rules.item(sel[0], "values")
        section, excel, _target, cls = values[0], values[1], values[2], values[3]
        back = {v: k for k, v in SECTION_NAMES.items()}
        if cls:
            rules = (self.aliases.data.get("class_subjects") or {}).get(cls, {})
            rules.pop(excel, None)
        else:
            self.aliases.remove(back.get(section, section), excel)
        self.tree_rules.delete(sel[0])
        self.save_aliases(silent=True)

    def save_aliases(self, silent: bool = False) -> None:
        path = self.var_aliases.get().strip() or os.path.join(BASE_DIR, "aliases.json")
        try:
            self.aliases.save(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Не удалось сохранить файл соответствий:\n{exc}")
            return
        try:
            self._aliases_mtime = os.path.getmtime(path)
        except OSError:
            pass
        if not silent:
            self.write_log(f"Файл соответствий сохранён: {path}")
            self.var_status.set(f"Сохранено: {path}")

    def reload_aliases_if_changed(self) -> None:
        """Подхватывает ручные правки aliases.json, сделанные вне программы."""
        path = self.var_aliases.get().strip() or os.path.join(BASE_DIR, "aliases.json")
        if not os.path.exists(path):
            return
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return
        if getattr(self, "_aliases_mtime", None) == mtime:
            return                      # диск и память синхронны
        # файл другой/изменён вне программы (в т.ч. вручную) — перечитываем,
        # чтобы не затереть ручные правки сохранением из памяти
        self.aliases = Aliases.load(path)
        self._aliases_mtime = mtime
        self._refresh_rules()
        self.write_log(f"Файл соответствий перезагружен с диска: {path}")

    def target_exists(self, section: str, value: str) -> bool:
        """Есть ли целевое название правила в справочниках загруженного эталона."""
        if self.ref is None and not self.load_reference_dicts(silent=True, refresh=False):
            return True          # эталон не загружен — проверим при конвертации
        if section == "subjects":
            return self.ref.subject_by_text(value) is not None
        if section == "rooms":
            return self.ref.room_by_text(value) is not None
        if section == "classes":
            return bool(self.ref.classes_by_text(value))
        if section == "days":
            return self.ref.day_by_text(value) is not None
        if section == "teachers":
            from converter.teachers import TeacherIndex
            return bool(TeacherIndex(self.ref).tids_by_text(value)) or value in self.ref.teachers
        return True

    # ------------------------------------------------------------------ #
    # конвертация
    # ------------------------------------------------------------------ #
    def _collect_options(self, dry_run: bool) -> core.Options:
        excel = self.var_excel.get().strip()
        reference = self.var_reference.get().strip()
        output = self.var_output.get().strip()
        if not output and excel:
            base = os.path.splitext(os.path.basename(excel))[0]
            output = os.path.join(os.path.dirname(excel), f"{base}_ExportCm.xml")
            self.var_output.set(output)
        aliases_path = self.var_aliases.get().strip() or os.path.join(BASE_DIR, "aliases.json")
        sheet = self.var_sheet.get().strip()
        return core.Options(
            excel_path=excel,
            reference_path=reference,
            output_path="" if dry_run else output,
            teachers_path=self.var_teachers.get().strip(),
            aliases_path=aliases_path,
            sheet=sheet if sheet and sheet != "(первый лист)" else None,
            week_start=self.var_week.get().strip(),
            skip_unknown_subject=self.var_skip_subject.get(),
            unknown_room_mode=self.var_room_mode.get(),
            align_groups=self.var_align.get(),
            check_conflicts=self.var_conflicts.get(),
            encoding=self.var_encoding.get().strip() or None,
            write_output=not dry_run,
            strict=self.var_strict.get(),
        )

    def on_convert(self) -> None:
        self._run(dry_run=False)

    def on_check(self) -> None:
        self._run(dry_run=True)

    def _run(self, dry_run: bool) -> None:
        if not self.var_excel.get().strip():
            messagebox.showwarning(APP_TITLE, "Выберите файл таблицы расписания.")
            return
        if not self.var_reference.get().strip():
            messagebox.showwarning(APP_TITLE, "Выберите эталонный XML системы (ExportCm.xml).")
            return
        if not dry_run and not self.var_output.get().strip():
            messagebox.showwarning(APP_TITLE, "Укажите, куда сохранить результат.")
            return
        self.reload_aliases_if_changed()
        self.save_aliases(silent=True)
        opts = self._collect_options(dry_run)
        self._set_busy(True, "Выполняется конвертация…" if not dry_run else "Проверка без записи файла…")
        self.write_log("—" * 70)
        self.write_log("ПРОВЕРКА (файл не записывается)" if dry_run else "КОНВЕРТАЦИЯ")

        self._pending = None

        def worker():
            # результат только кладём в очередь: с Tk работает главный поток
            try:
                self._pending = (core.convert(opts), None)
            except Exception:
                self._pending = (None, traceback.format_exc())

        threading.Thread(target=worker, daemon=True).start()
        self._poll_result()

    def _poll_result(self) -> None:
        pending = getattr(self, "_pending", None)
        if pending is None:
            self.after(120, self._poll_result)
            return
        self._pending = None
        self._done(pending[0], pending[1])

    def _done(self, res, error) -> None:
        self._set_busy(False, "")
        if error:
            self.write_log("ОШИБКА:\n" + error)
            self.var_status.set("Ошибка — см. журнал.")
            messagebox.showerror(APP_TITLE, "Ошибка при конвертации:\n" + error.strip().splitlines()[-1])
            return
        self.result = res
        for line in res.messages:
            self.write_log(line)
        self._set_text(self.txt_report, res.report_text)
        self.unresolved = list(res.unresolved)
        self.list_unresolved.delete(0, "end")
        for u in self.unresolved:
            self.list_unresolved.insert(
                "end", f"[{SECTION_NAMES.get(u.section, u.section)}] {u.value}  ×{u.count}"
                       + (f"   → возможно: {u.suggestion}" if u.suggestion else ""))
        self.lbl_map_hint.config(
            text=f"Нераспознанных значений: {len(self.unresolved)}. Выберите значение слева и укажите "
                 f"соответствие справа — затем «Пересчитать»."
            if self.unresolved else
            "Всё распознано — соответствия не требуются.")
        st = res.stats
        msg = (f"Уроков в XML: {st.get('placements')}  |  ячеек пропущено: {st.get('cells_skipped')}  |  "
               f"нераспознано: {st.get('unresolved')}  |  ошибок: {st.get('errors')}  |  "
               f"предупреждений: {st.get('warnings')}")
        if res.output_path:
            msg += f"  |  файл: {res.output_path}"
        self.var_status.set(msg)
        self.write_log(msg)
        self.notebook.select(self.tab_report if not self.unresolved else self.tab_map)
        if self.unresolved:
            self.load_reference_dicts(silent=True)
        if not res.ok:
            messagebox.showwarning(APP_TITLE, "Строгий режим: файл не записан, есть нераспознанные значения.")
        elif res.output_path and st.get("errors"):
            messagebox.showwarning(
                APP_TITLE,
                f"Файл записан, но {st.get('errors')} значений не распознано — "
                f"эти уроки НЕ попали в XML.\nСмотрите вкладку «Соответствия» и отчёт.")

    def _set_busy(self, busy: bool, status: str) -> None:
        self.config(cursor="watch" if busy else "")
        state = "disabled" if busy else "normal"
        for widget in (self.btn_convert, self.btn_check):
            widget.configure(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
        if status:
            self.var_status.set(status)

    # ----------------------------- вкладка 4 --------------------------- #
    def _build_tab_reverse(self) -> None:
        frame = self.tab_reverse
        box = ttk.LabelFrame(frame, text=" Обратное преобразование: XML системы -> XLSX ")
        box.pack(fill="x", padx=8, pady=8)
        box.columnconfigure(1, weight=1)
        self._path_row(box, 0, "XML из системы:", self.var_rev_xml, self.pick_rev_xml,
                       hint="ExportCm.xml или любой файл формата TimeTableExchange")
        self._path_row(box, 1, "Куда сохранить XLSX:", self.var_rev_out, self.pick_rev_out,
                       hint="по умолчанию — рядом с XML")
        ttk.Checkbutton(box, text="Лист «Расписание по учителям»",
                        variable=self.var_rev_teachers).grid(row=2, column=0, columnspan=2,
                                                             sticky="w", padx=6, pady=2)
        ttk.Checkbutton(box, text="Лист «Сводка» (уроки по дням и классам)",
                        variable=self.var_rev_summary).grid(row=3, column=0, columnspan=2,
                                                            sticky="w", padx=6, pady=2)
        btns = ttk.Frame(frame)
        btns.pack(fill="x", padx=8, pady=4)
        self.btn_reverse = ttk.Button(btns, text="Преобразовать в Excel", command=self.on_reverse)
        self.btn_reverse.pack(side="left", padx=4)
        ttk.Button(btns, text="Открыть папку результата",
                   command=lambda: self._open_folder_of(self.var_rev_out.get())).pack(side="left", padx=4)
        note = ttk.Label(frame, foreground="#444", justify="left", wraplength=900, text=(
            "Получаются таблицы в привычном виде: «Дни | Уроки | классы…», ячейки "
            "«Предмет (кабинет)» и «Фамилия И.О. (кабинет)»; деление на группы — "
            "«Иностранный (английский) язык (305,307)». Дни недели объединены по строкам, "
            "шапка и колонка дней оформлены, окно фиксируется на C2."))
        note.pack(fill="x", padx=10, pady=4)
        box_log = ttk.LabelFrame(frame, text=" Результат ")
        box_log.pack(fill="both", expand=True, padx=8, pady=8)
        self.txt_reverse = scrolledtext.ScrolledText(box_log, wrap="word", state="disabled",
                                                     font=("Consolas", 9))
        self.txt_reverse.pack(fill="both", expand=True, padx=4, pady=4)

    def pick_rev_xml(self) -> None:
        path = filedialog.askopenfilename(
            title="XML из системы", filetypes=[("XML", "*.xml"), ("Все файлы", "*.*")],
            initialdir=os.path.dirname(self.var_rev_xml.get() or self.var_reference.get() or BASE_DIR)
            or BASE_DIR)
        if path:
            self.var_rev_xml.set(path)
            if not self.var_rev_out.get():
                base = os.path.splitext(os.path.basename(path))[0]
                self.var_rev_out.set(os.path.join(os.path.dirname(path), f"{base}_расписание.xlsx"))

    def pick_rev_out(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить XLSX", defaultextension=".xlsx",
            filetypes=[("Таблицы Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialfile=os.path.basename(self.var_rev_out.get() or "расписание.xlsx"))
        if path:
            self.var_rev_out.set(path)

    def _open_folder_of(self, path: str) -> None:
        folder = os.path.dirname(path or "") or BASE_DIR
        if not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, f"Папка не найдена: {folder}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as exc:
            messagebox.showwarning(APP_TITLE, f"Не удалось открыть папку:\n{exc}")

    def on_reverse(self) -> None:
        xml = self.var_rev_xml.get().strip()
        if not xml or not os.path.exists(xml):
            messagebox.showwarning(APP_TITLE, "Выберите XML-файл системы для обратного преобразования.")
            return
        from converter import reverse
        out = self.var_rev_out.get().strip()
        self._set_busy(True, "Обратное преобразование XML → Excel…")

        def worker():
            try:
                self._reverse_pending = (reverse.convert(reverse.ReverseOptions(
                    xml_path=xml, output_path=out,
                    teachers_sheet=self.var_rev_teachers.get(),
                    summary_sheet=self.var_rev_summary.get())), None)
            except Exception:
                self._reverse_pending = (None, traceback.format_exc())

        threading.Thread(target=worker, daemon=True).start()
        self._poll_reverse()

    def _poll_reverse(self) -> None:
        pending = getattr(self, "_reverse_pending", None)
        if pending is None:
            self.after(120, self._poll_reverse)
            return
        self._reverse_pending = None
        self._reverse_done(pending[0], pending[1])

    def _reverse_done(self, res, error) -> None:
        self._set_busy(False, "")
        if error:
            self._set_text(self.txt_reverse, "ОШИБКА:\n" + error)
            self.var_status.set("Ошибка обратного преобразования — см. вкладку «XML → Excel».")
            messagebox.showerror(APP_TITLE, "Ошибка обратного преобразования:\n"
                                 + error.strip().splitlines()[-1])
            return
        self._set_text(self.txt_reverse, "\n".join(res.messages) + "\n\n"
                       + f"Дней: {res.days}; номеров уроков: {res.periods}; классов: {res.classes};\n"
                       + f"заполненных ячеек: {res.cells}; записей <csg>: {res.csg_total}.")
        for line in res.messages:
            self.write_log(line)
        self.var_status.set(f"XLSX сохранён: {res.output_path} ({res.cells} ячеек).")

    # ------------------------------------------------------------------ #
    # журнал, отчёты, справка
    # ------------------------------------------------------------------ #
    def write_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    @staticmethod
    def _set_text(widget, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def save_report(self) -> None:
        if not self.result:
            messagebox.showinfo(APP_TITLE, "Сначала выполните конвертацию или проверку.")
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить отчёт", defaultextension=".txt",
            filetypes=[("Текст", "*.txt"), ("Все файлы", "*.*")], initialfile="отчёт_конвертации.txt")
        if path:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(self.result.report_text)
            self.var_status.set(f"Отчёт сохранён: {path}")

    def save_unresolved(self) -> None:
        self._save_csv(self.result.unresolved_csv if self.result else "",
                       "нераспознанное.csv", "Список нераспознанного сохранён")

    def save_mapping(self) -> None:
        self._save_csv(self.result.mapping_csv if self.result else "",
                       "протокол_сопоставления.csv", "Протокол сопоставления сохранён")

    def _save_csv(self, content: str, default_name: str, message: str) -> None:
        if not content:
            messagebox.showinfo(APP_TITLE, "Сначала выполните конвертацию или проверку.")
            return
        path = filedialog.asksaveasfilename(
            title=message, defaultextension=".csv", filetypes=[("CSV", "*.csv")],
            initialfile=default_name)
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as fh:
            fh.write(content)
        self.var_status.set(f"{message}: {path}")

    def open_output_folder(self) -> None:
        path = self.var_output.get().strip()
        folder = os.path.dirname(path) or os.path.dirname(self.var_excel.get()) or BASE_DIR
        if not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, f"Папка не найдена: {folder}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as exc:
            messagebox.showwarning(APP_TITLE, f"Не удалось открыть папку:\n{exc}")

    def check_reference(self) -> None:
        """Проверка самого эталонного файла: битые ссылки, дубли, конфликты слотов."""
        path = self.var_reference.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "Сначала выберите эталонный XML системы.")
            return
        self._set_busy(True, "Проверка эталонного XML…")
        try:
            from converter import validate
            ref = reference_mod.load(path)
            findings = validate.validate(ref)
            text = validate.format_findings(ref, findings)
        except Exception as exc:
            self._set_busy(False, "")
            self.write_log(f"Не удалось проверить эталон: {exc}")
            messagebox.showerror(APP_TITLE, f"Ошибка разбора эталонного XML:\n{exc}")
            return
        self._set_busy(False, "")
        self.write_log("—" * 70)
        self.write_log(text)
        self._set_text(self.txt_report, text)
        self.notebook.select(self.tab_report)
        errors = sum(1 for f in findings if f.level == "error")
        warns = sum(1 for f in findings if f.level == "warning")
        self.var_status.set(f"Проверка эталона: ошибок {errors}, предупреждений {warns}.")

    def show_help(self) -> None:
        messagebox.showinfo(APP_TITLE, HELP_TEXT)

    def show_about(self) -> None:
        messagebox.showinfo(APP_TITLE,
                            "Конвертер расписания, версия 1.0\n\n"
                            "Excel-таблица расписания → XML формата TimeTableExchange\n"
                            "(словари и идентификаторы берутся из эталонной выгрузки системы).\n\n"
                            "Требуется Python 3.9+ и библиотека openpyxl.")


HELP_TEXT = """Как работает конвертер

1. Эталонный XML (ExportCm.xml из системы) используется как справочник:
   учителя, предметы, кабинеты, классы, учебный план (записи csg) и звонки.
2. Таблица Excel читается как «Дни | Уроки | классы…», ячейка — «Предмет (кабинет)».
3. Названия сопоставляются со справочниками: сначала точное совпадение, затем
   аббревиатуры, перестановка слов и нечёткое сравнение; при неудаче — по файлу
   соответствий aliases.json (вкладка «Соответствия»).
4. В итоговом XML словари копируются из эталона без изменений, пересоздаётся
   только блок <TimeTable> с уроками из таблицы.
5. Всё нераспознанное попадает в отчёт и НЕ записывается в XML (настраивается).

Кнопки: F5 — конвертировать и сохранить файл, F6 — проверить без записи."""


def _short(value, limit: int = 60) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def main() -> int:
    try:
        app = App()
    except tk.TclError as exc:
        print("Не удалось открыть графическое окно:", exc)
        print("Используйте консольный режим:  python -m converter.cli --help")
        return 1
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
