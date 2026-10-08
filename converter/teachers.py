# -*- coding: utf-8 -*-
"""Распознавание учителей из таблицы «расписание по учителям».

Ячейка такой таблицы — «Фамилия И.О. (кабинет)» или несколько учителей через запятую.
Инициалы в таблице могут быть записаны по-разному: «С.Е.», «Е..С.» (опечатка),
«О.Вл.» (две буквы отчества — способ различить двух Ольг Кузнецовых).
Поэтому сопоставление идёт так: фамилия точно (иначе нечётко) + первая буква имени +
отчество по ПРЕФИКСУ («вл» входит в «владимировна», но не в «васильевна»).
Если осталось несколько кандидатов — выбор сужается по учебному плану (кто ведёт
данный предмет у данного класса).
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from .reference import Reference, Teacher
from .text_utils import best_match, norm

_INITIALS_RE = re.compile(r"[a-zа-яё]+")
_NAME_RE = re.compile(r"^([^\s,.]+)\s*(.*)$")


def parse_fio(text: str) -> Tuple[str, List[str]]:
    """'Кузнецова О.Вл.' -> ('кузнецова', ['о', 'вл']); 'Бочкарева Е..С.' -> ('бочкарева', ['е','с'])."""
    m = _NAME_RE.match(norm(text))
    if not m:
        return "", []
    surname = m.group(1)
    groups = _INITIALS_RE.findall(m.group(2))
    return surname, [g[:2] for g in groups]


class TeacherIndex:
    def __init__(self, ref: Reference) -> None:
        self.ref = ref
        self.by_surname: Dict[str, List[Teacher]] = defaultdict(list)
        self.fio_keys: Dict[str, List[str]] = {}
        for t in ref.teachers.values():
            self.by_surname[norm(t.lastname)].append(t)
            for key in (f"{t.lastname} {t.firstname} {t.middlename}",
                        f"{t.lastname} {t.firstname[:1]}.{t.middlename[:1]}.",
                        f"{t.lastname} {t.firstname[:1]}.{t.middlename[:2]}."):
                self.fio_keys.setdefault(norm(key), [])
                if t.tid not in self.fio_keys[norm(key)]:
                    self.fio_keys[norm(key)].append(t.tid)

    def tids_by_text(self, text: str) -> List[str]:
        """Точное совпадение строки ФИО; список, т.к. «Кузнецова О.В.» может быть двумя людьми."""
        return list(self.fio_keys.get(norm(text), []))

    def resolve(self, text: str) -> Tuple[List[str], str]:
        """-> (список tid, способ). Пустой список — учитель не распознан."""
        if not text or not text.strip():
            return [], ""
        exact = self.tids_by_text(text)
        if exact:
            return exact, "точное" if len(exact) == 1 else "точное, неоднозначно"
        surname, groups = parse_fio(text)
        if not surname:
            return [], ""
        candidates = self.by_surname.get(surname)
        how = "точное"
        if not candidates:
            guess, score, _ = best_match(surname, list(self.by_surname), threshold=0.85)
            if not guess:
                return [], ""
            candidates = self.by_surname.get(guess) or []
            how = f"фамилия нечётко ({score:.2f})"
        selected = []
        for t in candidates:
            if groups and groups[0] and t.firstname[:1].lower() != groups[0][0]:
                continue
            if len(groups) > 1 and groups[1] and not t.middlename.lower().startswith(groups[1]):
                continue
            selected.append(t)
        if not selected:
            # инициалы не подошли ни к одной записи справочника
            return [], ""
        return [t.tid for t in selected], how if len(selected) == 1 else f"{how}, неоднозначно"

    def display(self, tid: str) -> str:
        """Полное ФИО — чтобы в отчёте различались тёзки («Кузнецова Ольга Владимировна»)."""
        t = self.ref.teachers.get(tid)
        if not t:
            return tid or ""
        return f"{t.lastname} {t.firstname} {t.middlename}"

    @property
    def fio_list(self) -> List[str]:
        out = []
        for t in sorted(self.ref.teachers.values(), key=lambda x: norm(x.lastname + x.firstname)):
            out.append(f"{t.lastname} {t.firstname} {t.middlename}")
        return out
