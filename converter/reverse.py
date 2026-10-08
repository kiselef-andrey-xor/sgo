# -*- coding: utf-8 -*-
"""Обратный конвертер: XML формата TimeTableExchange -> читабельный XLSX.

Строит таблицы в том же виде, в котором школа ведёт расписание вручную:
  * лист «по предметам»:  Дни | Уроки | классы…,  ячейка «Предмет (кабинет)»;
  * лист «по учителям»:   те же колонки,          ячейка «Фамилия И.О. (кабинет)»;
  * лист «Сводка»:        неделя, число уроков по дням и классам.

Деление класса на группы превращается в знакомую запись
«Иностранный (английский) язык (305,307)», разные предметы в одном слоте —
«Информатика, Труд (технология) (320,110)».
"""
from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .reference import Reference, load as load_reference

DAY_SHORT = {
    "понедельник": "Пн", "вторник": "Вт", "среда": "Ср",
    "четверг": "Чт", "пятница": "Пт", "суббота": "Сб", "воскресенье": "Вс",
}

HEADER_FILL = "DDEBF7"
DAY_FILL = "F2F2F2"


@dataclass
class ReverseOptions:
    xml_path: str
    output_path: str = ""
    teachers_sheet: bool = True
    summary_sheet: bool = True


@dataclass
class ReverseResult:
    ok: bool = True
    output_path: str = ""
    messages: List[str] = field(default_factory=list)
    days: int = 0
    periods: int = 0
    classes: int = 0
    cells: int = 0          # заполненных ячеек сетки (один раз, не по листам)
    lessons: int = 0        # блоков <Lesson>
    csg_total: int = 0      # записей <csg> в расписании
    extra_rows: List[tuple] = field(default_factory=list)


@dataclass
class _Item:
    subject: str
    teacher: str
    room: str
    order: int


def _short_teacher(ref: Reference, tid: str) -> str:
    t = ref.teachers.get(tid)
    if not t:
        return ""
    return f"{t.lastname} {t.firstname[:1]}.{t.middlename[:1]}."


def collect(ref: Reference) -> Tuple[Dict[tuple, List[_Item]], List[tuple], int]:
    """Разбирает <TimeTable> эталона в сетку {(день, номер, класс): [элементы]}."""
    root = ET.fromstring(ref.raw_text)
    grid: Dict[tuple, List[_Item]] = defaultdict(list)
    extra: List[tuple] = []
    lessons = 0
    csg_total = 0
    number_by_time = {lt.id: lt.number for lt in ref.lesson_times}
    for day_el in root.findall("./TimeTable/Week/Day"):
        day_name = day_el.get("name") or ""
        for lesson in day_el.findall("Lesson"):
            period_raw = number_by_time.get(lesson.get("timeId"), "")
            try:
                period = int(period_raw)
            except (TypeError, ValueError):
                continue
            lessons += 1
            for c_el in lesson.findall("csg"):
                csg_total += 1
                csg = ref.plan.get(c_el.get("id"))
                if csg is None:
                    extra.append((day_name, period, c_el.get("id"), "", "", ""))
                    continue
                cls = ref.classes.get(csg.owner_classid) or ref.classes.get(csg.pclassid)
                room = ref.rooms.get(c_el.get("roomid"))
                item = _Item(
                    subject=csg.base_name,
                    teacher=_short_teacher(ref, csg.tid),
                    room=room.name if room else "",
                    order=csg.order,
                )
                if cls is None:
                    extra.append((day_name, period, csg.name, csg.group_label,
                                  item.teacher, item.room))
                else:
                    grid[(day_name, period, cls.name)].append(item)
    for items in grid.values():
        items.sort(key=lambda x: x.order)
    return grid, extra, lessons, csg_total


def cell_text(items: List[_Item], kind: str) -> str:
    """Собирает текст ячейки в привычном школьном виде.

    Все группы с кабинетами  -> «Предмет1, Предмет2 (каб1,каб2)» (одна строка).
    Есть группа без кабинета -> по строке на группу: «Предмет (каб)» / «Предмет»,
    иначе пара «предмет-кабинет» терялась бы при обратном чтении.
    """
    def name_of(it: _Item) -> str:
        return (it.subject if kind == "subject" else it.teacher) or "?"

    if all(it.room for it in items):
        names: List[str] = []
        for it in items:
            nm = name_of(it)
            if nm not in names:
                names.append(nm)
        return ", ".join(names) + " (" + ",".join(it.room for it in items) + ")"
    lines = [f"{name_of(it)} ({it.room})" if it.room else name_of(it) for it in items]
    return "\n".join(lines)


def write_xlsx(ref: Reference, grid: Dict[tuple, List[_Item]], extra: List[tuple],
               opts: ReverseOptions, res: ReverseResult) -> None:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(bold=True)
    head_fill = PatternFill("solid", fgColor=HEADER_FILL)
    day_fill = PatternFill("solid", fgColor=DAY_FILL)
    center = Alignment(horizontal="center", vertical="center")

    used_days = sorted({d for (d, _p, _c) in grid}, key=lambda n: next(
        (x.order for x in ref.days if x.name == n), 99))
    periods = sorted({p for (_d, p, _c) in grid})
    classes = [ref.classes[i].name for i in ref.class_order
               if any(c == ref.classes[i].name for (_d, _p, c) in grid)]
    res.days, res.periods, res.classes = len(used_days), len(periods), len(classes)

    period_range = list(range(min(periods) if periods else 0, (max(periods) if periods else 0) + 1))

    def build_sheet(ws, kind: str) -> None:
        ws.cell(1, 1, "Дни").font = head_font
        ws.cell(1, 2, "Уроки").font = head_font
        for j, cname in enumerate(classes):
            ws.cell(1, 3 + j, cname)
        for col in range(1, 3 + len(classes)):
            c = ws.cell(1, col)
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = center
        row = 2
        for day in used_days:
            first = row
            for period in period_range:
                ws.cell(row, 1, DAY_SHORT.get(day, day) if period == period_range[0] else None)
                ws.cell(row, 2, period)
                ws.cell(row, 2).alignment = center
                for j, cname in enumerate(classes):
                    items = grid.get((day, period, cname))
                    if items:
                        ws.cell(row, 3 + j, cell_text(items, kind))
                for col in range(1, 3 + len(classes)):
                    ws.cell(row, col).border = border
                row += 1
            if row - 1 > first:
                ws.merge_cells(start_row=first, start_column=1, end_row=row - 1, end_column=1)
            dc = ws.cell(first, 1)
            dc.value = DAY_SHORT.get(day, day)
            dc.font = head_font
            dc.fill = day_fill
            dc.alignment = center
        ws.freeze_panes = "C2"
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 7
        for j in range(len(classes)):
            ws.column_dimensions[get_column_letter(3 + j)].width = 34 if kind == "subject" else 30

    ws1 = wb.active
    ws1.title = "Расписание по предметам"
    build_sheet(ws1, "subject")
    if opts.teachers_sheet:
        ws2 = wb.create_sheet("Расписание по учителям")
        build_sheet(ws2, "teacher")
    if extra:
        wsx = wb.create_sheet("Прочее")
        for j, title in enumerate(("День", "Урок", "Запись плана", "Группа", "Учитель", "Кабинет")):
            c = wsx.cell(1, 1 + j, title)
            c.font = head_font
            c.fill = head_fill
        for i, rowdata in enumerate(extra):
            for j, value in enumerate(rowdata):
                wsx.cell(2 + i, 1 + j, value)
        res.extra_rows = extra
    if opts.summary_sheet:
        wss = wb.create_sheet("Сводка")
        wss.cell(1, 1, "Неделя").font = head_font
        wss.cell(1, 2, ref.week_name or "")
        rows = [("Дней", res.days), ("Номеров уроков", res.periods), ("Классов", res.classes),
                ("Заполненных ячеек", res.cells), ("Блоков <Lesson>", res.lessons)]
        per_day = defaultdict(int)
        for (day, _p, _c), items in grid.items():
            per_day[day] += len(items)
        per_class = defaultdict(int)
        for (_d, _p, cls), items in grid.items():
            per_class[cls] += len(items)
        r = 3
        wss.cell(r, 1, "По дням").font = head_font
        r += 1
        for day in used_days:
            wss.cell(r, 1, DAY_SHORT.get(day, day))
            wss.cell(r, 2, per_day[day])
            r += 1
        r += 1
        wss.cell(r, 1, "По классам").font = head_font
        r += 1
        for cname in classes:
            wss.cell(r, 1, cname)
            wss.cell(r, 2, per_class[cname])
            r += 1
        for i, (k, v) in enumerate(rows):
            wss.cell(3 + i, 4, k).font = head_font
            wss.cell(3 + i, 5, v)
        wss.column_dimensions["A"].width = 22
        wss.column_dimensions["D"].width = 20

    folder = os.path.dirname(os.path.abspath(opts.output_path))
    if folder:
        os.makedirs(folder, exist_ok=True)
    wb.save(opts.output_path)


def convert(opts: ReverseOptions) -> ReverseResult:
    res = ReverseResult()
    if not os.path.exists(opts.xml_path):
        raise FileNotFoundError(f"Файл XML не найден: {opts.xml_path}")
    if not opts.output_path:
        base = os.path.splitext(os.path.basename(opts.xml_path))[0]
        opts.output_path = os.path.join(os.path.dirname(os.path.abspath(opts.xml_path)),
                                        f"{base}_расписание.xlsx")
    res.messages.append(f"Читаю XML: {opts.xml_path}")
    ref = load_reference(opts.xml_path)
    res.messages.append(f"  неделя «{ref.week_name}», классов {len(ref.classes)}, "
                        f"кабинетов {len(ref.rooms)}, учителей {len(ref.teachers)}")
    grid, extra, lessons, csg_total = collect(ref)
    res.lessons = lessons
    res.csg_total = csg_total
    res.cells = len(grid)
    res.messages.append(f"  блоков уроков: {lessons}, записей <csg>: {csg_total}, "
                        f"ячеек сетки: {len(grid)}"
                        + (f", записей вне классов: {len(extra)}" if extra else ""))
    write_xlsx(ref, grid, extra, opts, res)
    res.output_path = opts.output_path
    res.messages.append(f"Записан файл: {opts.output_path} "
                        f"(листы: по предметам{', по учителям' if opts.teachers_sheet else ''}"
                        f"{', сводка' if opts.summary_sheet else ''}"
                        f"{', прочее' if extra else ''})")
    res.ok = True
    return res
