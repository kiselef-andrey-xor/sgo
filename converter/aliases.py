# -*- coding: utf-8 -*-
"""Файл соответствий (aliases.json).

Хранит правила «значение в Excel -> значение в системе»:
  * subjects — названия предметов;
  * rooms    — названия кабинетов;
  * classes  — названия колонок-классов (в т.ч. деление профильных классов);
  * days     — названия дней;
  * class_subjects — индивидуальные правила «класс → предмет → предмет в системе»
    (нужны, когда один и тот же текст в таблице означает разное в разных классах).

Значение может быть:
  * строкой  — название из справочника системы («Спортивный зал 1»);
  * null     — осознанно пропускать (например, «Разговоры о важном»);
  * для classes — строкой, списком или объектом
    {"candidates": [...], "by_subject": {"химия": "11и (ест.-научн.)"}}.
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

from .text_utils import norm, norm_strict

IGNORE_WORDS = {"", "-", "—", "пропустить", "skip", "ignore", "нет", "none", "null"}

DEFAULT_TEMPLATE: Dict[str, Any] = {
    "_комментарий": [
        "Файл соответствий для конвертера расписания Excel -> TimeTableExchange XML.",
        "Ключ — значение из таблицы Excel (регистр и пробелы не важны),",
        "значение — точное название из справочника системы либо null (пропускать).",
        "Правится вручную или через вкладку «Соответствия» в программе.",
    ],
    "subjects": {},
    "rooms": {},
    "classes": {},
    "days": {},
    "teachers": {},
    "class_subjects": {},
}


class Aliases:
    def __init__(self, data: Optional[Dict[str, Any]] = None, path: Optional[str] = None) -> None:
        data = data or {}
        self.path = path
        self.data: Dict[str, Any] = {
            "subjects": dict(data.get("subjects") or {}),
            "rooms": dict(data.get("rooms") or {}),
            "classes": dict(data.get("classes") or {}),
            "days": dict(data.get("days") or {}),
            "teachers": dict(data.get("teachers") or {}),
            "class_subjects": dict(data.get("class_subjects") or {}),
        }
        self._index: Dict[str, Dict[str, Any]] = {}
        for section in ("subjects", "rooms", "classes", "days", "teachers"):
            self._index[section] = {norm_strict(k): (k, v) for k, v in self.data[section].items()}

    # ------------------------------------------------------------------ #
    @classmethod
    def load(cls, path: Optional[str]) -> "Aliases":
        if not path or not os.path.exists(path):
            return cls({}, path)
        with open(path, "r", encoding="utf-8") as fh:
            return cls(json.load(fh), path)

    @classmethod
    def default_file(cls, path: str) -> "Aliases":
        """Создаёт файл-шаблон, если его ещё нет."""
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(DEFAULT_TEMPLATE, fh, ensure_ascii=False, indent=2)
        return cls.load(path)

    def save(self, path: Optional[str] = None) -> str:
        path = path or self.path
        if not path:
            raise ValueError("Не указан файл для сохранения соответствий")
        payload: Dict[str, Any] = {}
        comment = (self.data.get("_комментарий") if isinstance(self.data, dict) else None)
        payload["_комментарий"] = comment or DEFAULT_TEMPLATE["_комментарий"]
        for section in ("subjects", "rooms", "classes", "days", "teachers", "class_subjects"):
            payload[section] = self.data.get(section, {})
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2, sort_keys=False)
        self.path = path
        return path

    # ------------------------------------------------------------------ #
    def get(self, section: str, text: str) -> Optional[Any]:
        item = self._index.get(section, {}).get(norm_strict(text))
        return item[1] if item else None

    def has(self, section: str, text: str) -> bool:
        return norm_strict(text) in self._index.get(section, {})

    def set(self, section: str, excel_text: str, target: Any) -> None:
        """Добавляет/обновляет правило. target=None — пропускать значение."""
        key = norm_strict(excel_text)
        if not key:
            return
        original = excel_text if isinstance(excel_text, str) else str(excel_text)
        # убираем старую запись с тем же нормализованным ключом
        old = self._index.get(section, {}).get(key)
        if old:
            self.data[section].pop(old[0], None)
        self.data[section][original] = target
        self._index.setdefault(section, {})[key] = (original, target)

    def remove(self, section: str, excel_text: str) -> None:
        key = norm_strict(excel_text)
        old = self._index.get(section, {}).pop(key, None)
        if old:
            self.data[section].pop(old[0], None)

    def items(self, section: str) -> List[tuple]:
        return [(k, v) for k, v in self.data.get(section, {}).items()]

    @staticmethod
    def is_ignore(value: Any) -> bool:
        if value is None:
            return True
        if isinstance(value, str) and norm(value) in IGNORE_WORDS:
            return True
        return False
