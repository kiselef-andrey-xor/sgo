# -*- coding: utf-8 -*-
"""Проверка целостности эталонной выгрузки (и любого файла формата TimeTableExchange).

Ищет «битые» ссылки и противоречия до конвертации, чтобы импорт не падал:
дубликаты id, висячие tid/sid/roomid/csg id/timeId, дубли уроков в одном слоте,
занятость учителей и кабинетов, расхождения часов с планом.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Dict, List, Tuple

from .reference import Reference


@dataclass
class Finding:
    level: str      # error | warning | info
    text: str
    count: int = 1


def _duplicates(values: List[str]) -> List[str]:
    return [v for v, n in Counter(values).items() if n > 1]


def validate(ref: Reference) -> List[Finding]:
    out: List[Finding] = []

    def add(level: str, text: str, count: int = 1) -> None:
        out.append(Finding(level, text, count))

    # --- повторное чтение сырого XML: ловим дубликаты атрибутов-идентов
    try:
        root = ET.fromstring(ref.raw_text)
    except ET.ParseError as exc:
        add("error", f"XML не разбирается: {exc}")
        return out

    for section, tag, attr in (("LessonTimes", "LessonTime", "id"), ("teachers", "teacher", "tid"),
                               ("subjects", "subject", "sid"), ("Rooms", "room", "id"),
                               ("Plan", "class", "id")):
        ids = [el.get(attr) for el in root.findall(f"./{section}/{tag}")]
        dup = _duplicates(ids)
        if dup:
            add("error", f"{section}: повторяющиеся {attr}: {', '.join(dup[:10])}", len(dup))

    csg_ids = [c.get("id") for cls in root.findall("./Plan/class") for c in cls.findall("csg")]
    dup = _duplicates(csg_ids)
    if dup:
        add("error", f"Plan: повторяющиеся id записей плана (csg): {', '.join(dup[:10])}", len(dup))

    # --- висячие ссылки в словаре плана
    bad_sid = Counter(c.sid for cls in ref.classes.values() for c in cls.csg
                      if c.sid and c.sid not in ref.subjects)
    if bad_sid:
        add("error", f"Plan: csg ссылается на несуществующий предмет (sid): "
                     f"{', '.join(list(bad_sid)[:8])}", sum(bad_sid.values()))
    bad_tid = Counter(c.tid for cls in ref.classes.values() for c in cls.csg
                      if c.tid and c.tid not in ref.teachers)
    if bad_tid:
        add("warning", f"Plan: csg ссылается на несуществующего учителя (tid): "
                       f"{', '.join(list(bad_tid)[:8])}", sum(bad_tid.values()))
    bad_pclass = {c.pclassid for cls in ref.classes.values() for c in cls.csg
                  if c.pclassid and c.pclassid not in ref.classes}
    if bad_pclass:
        add("error", f"Plan: pclassid вне справочника классов: {', '.join(sorted(bad_pclass)[:8])}",
            len(bad_pclass))
    orphan_used = 0
    nested = {c.id for cls in ref.classes.values() for c in cls.csg}
    for c in ref.orphan_csg:
        if c.id not in nested:
            orphan_used += 1
    if orphan_used:
        add("info", f"Plan: csg без pclassid лежат вне элементов <class>: {orphan_used}")

    bad_subj_teacher = Counter(t.get("tid") for s in root.findall("./subjects/subject")
                               for t in s.findall("teacher")
                               if t.get("tid") and t.get("tid") not in ref.teachers)
    if bad_subj_teacher:
        add("warning", "subjects: ссылка на несуществующего учителя", sum(bad_subj_teacher.values()))

    # --- расписание: ссылки и дубли
    lesson_times = {lt.id for lt in ref.lesson_times}
    numbers = Counter(lt.number for lt in ref.lesson_times)
    dup_numbers = [n for n, k in numbers.items() if k > 1]
    if dup_numbers:
        add("error", f"LessonTimes: повторяющийся номер урока: {', '.join(dup_numbers)}", len(dup_numbers))

    tt_csg = 0
    dangling_csg: Counter = Counter()
    dangling_room: Counter = Counter()
    bad_time: Counter = Counter()
    dup_in_lesson = 0
    teacher_clash: Counter = Counter()
    room_clash: Counter = Counter()
    class_clash: Counter = Counter()
    week_counts: Counter = Counter()
    for day in root.findall("./TimeTable/Week/Day"):
        for lesson in day.findall("Lesson"):
            tid_ = lesson.get("timeId")
            if tid_ not in lesson_times:
                bad_time[tid_] += 1
            seen_csg, seen_teacher, seen_room, seen_class = set(), Counter(), Counter(), Counter()
            for c in lesson.findall("csg"):
                tt_csg += 1
                cid = c.get("id")
                week_counts[cid] += 1
                csg = ref.plan.get(cid)
                if csg is None:
                    dangling_csg[cid] += 1
                    continue
                if cid in seen_csg:
                    dup_in_lesson += 1
                seen_csg.add(cid)
                rid = c.get("roomid")
                if rid:
                    if rid not in ref.rooms:
                        dangling_room[rid] += 1
                    else:
                        seen_room[rid] += 1
                if csg.tid:
                    seen_teacher[csg.tid] += 1
                owner = csg.pclassid or csg.owner_classid
                if owner:
                    seen_class[owner] += 1
            for tid2, n in seen_teacher.items():
                if n > 1:
                    teacher_clash[ref.teacher_fio(tid2)] += n - 1
            for rid, n in seen_room.items():
                if n > 1:
                    room_clash[ref.rooms[rid].name] += n - 1
            for clsid, _n in seen_class.items():
                # несколько записей одного класса в одном слоте — это деление на группы
                # (groupid разные); подозрительно лишь когда одна и та же группа дважды
                pairs = Counter(
                    (ref.plan[c.get("id")].base_name, ref.plan[c.get("id")].groupid)
                    for c in lesson.findall("csg")
                    if c.get("id") in ref.plan and
                    (ref.plan[c.get("id")].pclassid or ref.plan[c.get("id")].owner_classid) == clsid)
                for (nm, _gid), k in pairs.items():
                    if k > 1:
                        class_clash[f"{ref.classes.get(clsid).name if clsid in ref.classes else clsid} / {nm}"] += 1
    if dangling_csg:
        add("error", f"TimeTable: csg id не найден в плане: {', '.join(list(dangling_csg)[:8])}…",
            sum(dangling_csg.values()))
    if dangling_room:
        add("error", f"TimeTable: roomid вне справочника кабинетов: {', '.join(list(dangling_room)[:8])}",
            sum(dangling_room.values()))
    if bad_time:
        add("error", f"TimeTable: timeId вне LessonTimes: {', '.join(sorted(bad_time)[:8])}",
            sum(bad_time.values()))
    if dup_in_lesson:
        add("error", "TimeTable: одна запись плана дважды в одном уроке", dup_in_lesson)
    if teacher_clash:
        add("warning", f"TimeTable: учитель дважды в одном слоте: "
                       f"{', '.join(list(teacher_clash)[:6])}…", sum(teacher_clash.values()))
    if room_clash:
        add("warning", f"TimeTable: кабинет занят дважды в одном слоте: "
                       f"{', '.join(list(room_clash)[:6])}…", sum(room_clash.values()))
    if class_clash:
        add("warning", f"TimeTable: один класс и один предмет дважды в одном слоте: "
                       f"{', '.join(list(class_clash)[:6])}…", sum(class_clash.values()))

    # --- часы против плана
    mismatch = [(ref.plan[c].owner_classname or ref.plan[c].pclassid, ref.plan[c].name,
                 ref.plan[c].hours, n) for c, n in week_counts.items()
                if c in ref.plan and ref.plan[c].hours and n != ref.plan[c].hours]
    if mismatch:
        add("info", f"TimeTable: расхождение фактических часов с hrsweek у {len(mismatch)} записей плана "
                    f"(например: {mismatch[0][0]} «{mismatch[0][1]}» план {mismatch[0][2]}, факт {mismatch[0][3]})")

    unused = len(ref.plan) - len(set(week_counts) & set(ref.plan))
    if unused:
        add("info", f"Plan: {unused} записей плана не используются в расписании недели")
    if not out:
        add("info", "замечаний нет: файл целостен")
    order = {"error": 0, "warning": 1, "info": 2}
    out.sort(key=lambda f: (order[f.level], -f.count, f.text))
    return out


def format_findings(ref: Reference, findings: List[Finding]) -> str:
    lines = ["=" * 78,
             f"ПРОВЕРКА ЦЕЛОСТНОСТИ ФАЙЛА: {ref.path}",
             f"неделя: {ref.week_name or '—'}; классов {len(ref.classes)}, предметов "
             f"{len(ref.subjects)}, кабинетов {len(ref.rooms)}, учителей {len(ref.teachers)}, "
             f"csg {len(ref.plan)}, ссылок в расписании "
             f"{sum(1 for _ in ref.raw_text.split('<csg id=')) - 1}",
             "=" * 78]
    for level in ("error", "warning", "info"):
        items = [f for f in findings if f.level == level]
        if not items:
            continue
        title = {"error": "ОШИБКИ", "warning": "ПРЕДУПРЕЖДЕНИЯ", "info": "СПРАВОЧНО"}[level]
        lines.append(f"\n{title} ({len(items)})")
        for f in items:
            suffix = f"  [×{f.count}]" if f.count > 1 else ""
            lines.append(f"  • {f.text}{suffix}")
    return "\n".join(lines)
