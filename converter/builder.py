# -*- coding: utf-8 -*-
"""Сборка итогового XML в формате TimeTableExchange.

Словари (LessonTimes, teachers, subjects, Rooms, Plan) копируются из эталона
дословно — так гарантированно сохраняются идентификаторы системы, порядок
атрибутов и форматирование. Пересоздаётся только блок <TimeTable>.
"""
from __future__ import annotations

import datetime as _dt
import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from .matching import Placement
from .reference import Reference

INDENT = "  "


def attr(name: str, value) -> str:
    text = "" if value is None else str(value)
    text = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))
    return f' {name}="{text}"'


def week_name_from_date(start: _dt.date) -> str:
    """«Неделя 14.09.2026 - 20.09.2026» по дате понедельника."""
    monday = start - _dt.timedelta(days=start.weekday())
    sunday = monday + _dt.timedelta(days=6)
    return f"Неделя {monday.strftime('%d.%m.%Y')} - {sunday.strftime('%d.%m.%Y')}"


def parse_week_start(text: str) -> Optional[_dt.date]:
    if not text:
        return None
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", str(text))
    if not m:
        return None
    year = int(m.group(3))
    if year < 100:
        year += 2000
    try:
        return _dt.date(year, int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def week_name_from_reference(name: str) -> Optional[_dt.date]:
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", name or "")
    if not m:
        return None
    try:
        return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def build_timetable(ref: Reference, placements: List[Placement],
                    week_id: Optional[str] = None,
                    week_name: Optional[str] = None) -> Tuple[str, List[str]]:
    """Формирует блок <TimeTable>…</TimeTable> и список сообщений."""
    notes: List[str] = []
    by_day_time: Dict[Tuple[int, int], List[Placement]] = defaultdict(list)
    for p in placements:
        by_day_time[(p.day.order, int(p.time.id))].append(p)

    w_id = week_id or ref.week_id or "1"
    w_name = week_name or ref.week_name
    if not w_name:
        w_name = "Неделя 1"

    out: List[str] = []
    out.append(f"{INDENT}<TimeTable>")
    out.append(f"{INDENT*2}<Week{attr('id', w_id)}{attr('name', w_name)}>")
    days_used = 0
    lessons_total = 0
    for day in ref.days:
        day_slots = {k: v for k, v in by_day_time.items() if k[0] == day.order}
        if not day_slots:
            notes.append(f"День «{day.name}»: занятий нет — блок <Day> не создаётся.")
            continue
        days_used += 1
        out.append(f"{INDENT*3}<Day{attr('id', day.id)}{attr('name', day.name)}{attr('wd', day.wd)}>")
        for (_, time_id) in sorted(day_slots, key=lambda k: k[1]):
            items = sorted(day_slots[(_, time_id)], key=lambda p: p.sort_key())
            lessons_total += 1
            out.append(f"{INDENT*4}<Lesson{attr('timeId', time_id)}>")
            for p in items:
                line = f"{INDENT*5}<csg{attr('id', p.csg.id)}"
                if p.room is not None:
                    line += attr("roomid", p.room.id)
                out.append(line + " />")
            out.append(f"{INDENT*4}</Lesson>")
        out.append(f"{INDENT*3}</Day>")
    out.append(f"{INDENT*2}</Week>")
    out.append(f"{INDENT}</TimeTable>")
    notes.insert(0, f"Сформировано: дней — {days_used}, блоков уроков — {lessons_total}, "
                    f"записей <csg> — {len(placements)}.")
    return "\n".join(out), notes


def assemble(ref: Reference, timetable_block: str) -> str:
    """Склеивает словари из эталона и новый блок расписания."""
    header = ref.header_text
    if header and not header.endswith("\n"):
        header += "\n"
    return f"{header}{timetable_block}\n{ref.footer_text}"


def encode(text: str, encoding: str) -> Tuple[bytes, List[str]]:
    """Кодирует текст, сообщая о символах, которых нет в целевой кодировке.

    Такие символы заменяются на «?» — иначе файл в windows-1251 не собрать.
    """
    try:
        return text.encode(encoding), []
    except UnicodeEncodeError:
        pass
    bad: List[str] = []
    out = bytearray()
    for ch in text:
        try:
            out += ch.encode(encoding)
        except UnicodeEncodeError:
            if ch not in bad:
                bad.append(ch)
            out += b"?"
    return bytes(out), bad


def write(path: str, text: str, encoding: str) -> Tuple[int, List[str]]:
    data, bad = encode(text, encoding)
    with open(path, "wb") as fh:
        fh.write(data)
    return len(data), bad
