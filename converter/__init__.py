# -*- coding: utf-8 -*-
"""Конвертер расписания: Excel -> XML (TimeTableExchange).

Пакет состоит из независимых модулей:
  reference    — чтение эталонной выгрузки системы (справочники и ID);
  excel_reader — разбор таблицы расписания;
  aliases      — файл соответствий названий (aliases.json);
  matching     — сопоставление и проверки (учителя, кабинеты, часы);
  builder      — генерация XML;
  report       — отчёт;
  core         — связка всего вместе (функция convert);
  cli          — запуск из командной строки;
  gui          — графическое окно (tkinter), модуль ../gui.py.
"""
from .core import ConvertResult, Options, convert, load_reference, unresolved  # noqa: F401

__version__ = "1.0.0"
