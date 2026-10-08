# -*- coding: utf-8 -*-
"""Загрузка эталонной выгрузки системы (формат TimeTableExchange).

Эталонный XML используется как справочник идентификаторов: учителя (tid),
предметы (sid), кабинеты (room id), классы (class id), записи учебного плана
(csg id), звонки (LessonTime id) и дни недели (Day id/wd).
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .text_utils import norm, norm_strict, token_key

DEFAULT_ENCODING = "windows-1251"

DAY_ALIASES = {
    "понедельник": "понедельник", "пн": "понедельник", "пон": "понедельник", "пнд": "понедельник",
    "вторник": "вторник", "вт": "вторник", "втр": "вторник",
    "среда": "среда", "ср": "среда", "сред": "среда",
    "четверг": "четверг", "чт": "четверг", "четв": "четверг", "4": "четверг",
    "пятница": "пятница", "пт": "пятница", "пятн": "пятница",
    "суббота": "суббота", "сб": "суббота", "суб": "суббота",
    "воскресенье": "воскресенье", "вс": "воскресенье", "воскр": "воскресенье",
}


@dataclass
class LessonTime:
    id: str
    number: str
    starttime: str
    endtime: str
    sm: str = "1"


@dataclass
class Teacher:
    tid: str
    lastname: str = ""
    firstname: str = ""
    middlename: str = ""

    @property
    def fio(self) -> str:
        parts = [self.lastname, self.firstname, self.middlename]
        return " ".join(p for p in parts if p)


@dataclass
class Subject:
    sid: str
    name: str
    abbr: str = ""
    teachers: List[str] = field(default_factory=list)


@dataclass
class Room:
    id: str
    name: str
    seats: str = ""
    floor: str = ""


SPLIT_GROUP_RE = re.compile(
    r"(?:п\.?\s*\d|подгруппа|подгр|группа\s*\d|\d\s*-?\s*(?:гр|группа)|\(\s*\d\s*\)|вариант)", re.IGNORECASE)


@dataclass
class Csg:
    """Запись учебного плана «класс-предмет-группа» (class-subject-group)."""
    id: str
    name: str
    tid: str = ""
    sid: str = ""
    pclassid: str = ""
    groupid: str = ""
    parentsubjectid: str = ""
    hrsweek: str = "0"
    studcnt: str = "0"
    order: int = 0
    owner_classid: str = ""   # класс, внутри которого записан <csg> в эталоне
    owner_classname: str = ""

    @property
    def is_attached(self) -> bool:
        """Есть ли явная привязка к классу (атрибут pclassid)."""
        return bool(self.pclassid)

    @property
    def students(self) -> int:
        try:
            return int(self.studcnt)
        except (TypeError, ValueError):
            return 0

    @property
    def base_name(self) -> str:
        """Название предмета без метки группы: «Информатика/п.1» -> «Информатика»."""
        return self.name.split("/")[0].strip()

    @property
    def group_label(self) -> str:
        return self.name.split("/", 1)[1].strip() if "/" in self.name else ""

    @property
    def hours(self) -> int:
        try:
            return int(self.hrsweek)
        except (TypeError, ValueError):
            return 0


@dataclass
class SchoolClass:
    id: str
    name: str
    grade: str = ""
    studcnt: str = ""
    boys: str = ""
    girls: str = ""
    order: int = 0
    csg: List[Csg] = field(default_factory=list)


@dataclass
class Day:
    id: str
    name: str
    wd: str
    order: int = 0


class Reference:
    """Разобранная эталонная выгрузка + её исходный текст."""

    def __init__(self) -> None:
        self.path: str = ""
        self.encoding: str = DEFAULT_ENCODING
        self.raw_text: str = ""
        self.lesson_times: List[LessonTime] = []
        self.teachers: Dict[str, Teacher] = {}
        self.subjects: Dict[str, Subject] = {}
        self.rooms: Dict[str, Room] = {}
        self.classes: Dict[str, SchoolClass] = {}
        self.class_order: List[str] = []
        self.days: List[Day] = []
        self.plan: Dict[str, Csg] = {}          # csg id -> Csg
        self.orphan_csg: List[Csg] = []        # csg без pclassid (элективы/индивидуальные)
        self.week_id: str = "1"
        self.week_name: str = ""
        self.header_text: str = ""             # всё до <TimeTable> — копируется в результат дословно
        self.footer_text: str = "</TimeTableExchange>"
        self.room_history: Dict[str, Counter] = defaultdict(Counter)  # csg id -> счётчик кабинетов
        # индексы для сопоставления
        self._room_keys: Dict[str, str] = {}
        self._subject_keys: Dict[str, str] = {}
        self._class_keys: Dict[str, List[str]] = defaultdict(list)
        self._day_keys: Dict[str, str] = {}
        self._time_by_number: Dict[str, LessonTime] = {}
        self._time_by_start: Dict[str, LessonTime] = {}

    # ------------------------------------------------------------------ #
    # справочные выборки
    # ------------------------------------------------------------------ #
    def room_by_text(self, key: str) -> Optional[Room]:
        rid = self._room_keys.get(norm_strict(key))
        return self.rooms.get(rid) if rid else None

    def subject_by_text(self, key: str) -> Optional[Subject]:
        sid = self._subject_keys.get(norm_strict(key)) or self._subject_keys.get(token_key(key))
        return self.subjects.get(sid) if sid else None

    def classes_by_text(self, key: str) -> List[SchoolClass]:
        """Классы, подходящие под название колонки таблицы.

        Точное совпадение даёт один класс; «11и» может дать два профильных
        класса «11и (ест.-научн.)» и «11и (техн.)».
        """
        nk = norm_strict(key)
        ids = self._class_keys.get(nk, [])
        if not ids:
            ids = self._class_keys.get(token_key(key), [])
        return [self.classes[i] for i in ids if i in self.classes]

    def day_by_text(self, key: str) -> Optional[Day]:
        nk = norm(key).strip()
        did = self._day_keys.get(nk) or self._day_keys.get(DAY_ALIASES.get(nk, ""))
        for d in self.days:
            if d.id == did:
                return d
        return None

    def time_by_number(self, number) -> Optional[LessonTime]:
        return self._time_by_number.get(str(number).strip())

    def time_by_start(self, text) -> Optional[LessonTime]:
        return self._time_by_start.get(norm_strict(text))

    @property
    def room_names(self) -> List[str]:
        return [r.name for r in self.rooms.values()]

    @property
    def subject_names(self) -> List[str]:
        out = []
        for s in self.subjects.values():
            out.append(s.name)
            if s.abbr and s.abbr not in out:
                out.append(s.abbr)
        return out

    @property
    def class_names(self) -> List[str]:
        return [self.classes[i].name for i in self.class_order]

    def teacher_fio(self, tid: str) -> str:
        t = self.teachers.get(tid)
        return t.fio if t else (tid or "")


# ---------------------------------------------------------------------- #
# загрузка
# ---------------------------------------------------------------------- #
def detect_encoding(data: bytes) -> str:
    m = re.match(rb'\s*<\?xml[^>]*encoding=["\']([\w\-]+)["\']', data[:200])
    if m:
        enc = m.group(1).decode("ascii", "ignore")
        try:
            data.decode(enc)
            return enc
        except (LookupError, UnicodeDecodeError):
            pass
    if data.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    for enc in ("utf-8", DEFAULT_ENCODING):
        try:
            data.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return DEFAULT_ENCODING


def load(path: str) -> Reference:
    """Читает эталонный XML и строит все индексы."""
    with open(path, "rb") as fh:
        data = fh.read()
    ref = Reference()
    ref.path = path
    ref.encoding = detect_encoding(data)
    text = data.decode(ref.encoding, errors="replace")
    ref.raw_text = text

    root = ET.fromstring(text)

    # --- звонки
    for i, el in enumerate(root.findall("./LessonTimes/LessonTime")):
        lt = LessonTime(
            id=el.get("id") or str(i + 1),
            number=el.get("number") or str(i),
            starttime=el.get("starttime") or "",
            endtime=el.get("endtime") or "",
            sm=el.get("sm") or "1",
        )
        ref.lesson_times.append(lt)
        ref._time_by_number[lt.number] = lt
        if lt.starttime:
            ref._time_by_start[norm_strict(lt.starttime)] = lt

    # --- учителя
    for el in root.findall("./teachers/teacher"):
        t = Teacher(
            tid=el.get("tid") or "",
            lastname=el.get("lastname") or "",
            firstname=el.get("firstname") or "",
            middlename=el.get("middlename") or "",
        )
        if t.tid:
            ref.teachers[t.tid] = t

    # --- предметы
    for el in root.findall("./subjects/subject"):
        s = Subject(sid=el.get("sid") or "", name=el.get("name") or "", abbr=el.get("abbr") or "")
        s.teachers = [t.get("tid") for t in el.findall("teacher") if t.get("tid")]
        if s.sid:
            ref.subjects[s.sid] = s
            for key in {norm_strict(s.name), token_key(s.name), norm_strict(s.abbr), token_key(s.abbr)}:
                if key:
                    ref._subject_keys.setdefault(key, s.sid)

    # --- кабинеты
    for el in root.findall("./Rooms/room"):
        r = Room(id=el.get("id") or "", name=el.get("name") or "",
                 seats=el.get("seats") or "", floor=el.get("floor") or "")
        if r.id:
            ref.rooms[r.id] = r
            key = norm_strict(r.name)
            if key:
                ref._room_keys.setdefault(key, r.id)

    # --- учебный план (классы + csg)
    for ci, el in enumerate(root.findall("./Plan/class")):
        cls = SchoolClass(
            id=el.get("id") or "", name=el.get("name") or "", grade=el.get("grade") or "",
            studcnt=el.get("studcnt") or "", boys=el.get("boys") or "", girls=el.get("girls") or "",
            order=ci,
        )
        for csg_el in el.findall("csg"):
            csg = Csg(
                id=csg_el.get("id") or "",
                name=csg_el.get("name") or "",
                tid=csg_el.get("tid") or "",
                sid=csg_el.get("sid") or "",
                pclassid=csg_el.get("pclassid") or "",
                groupid=csg_el.get("groupid") or "",
                parentsubjectid=csg_el.get("parentsubjectid") or "",
                hrsweek=csg_el.get("hrsweek") or "0",
                studcnt=csg_el.get("studcnt") or "0",
                order=len(cls.csg),
                owner_classid=cls.id,
                owner_classname=cls.name,
            )
            cls.csg.append(csg)
            if csg.id:
                ref.plan[csg.id] = csg
            if not csg.pclassid:
                # вложенные <csg> без pclassid: индивидуальные/групповые занятия
                ref.orphan_csg.append(csg)
        if cls.id:
            ref.classes[cls.id] = cls
            ref.class_order.append(cls.id)
            base = re.sub(r"\s*[\(\[].*?[\)\]]\s*", " ", cls.name)  # «11и (техн.)» -> «11и»
            for key in {norm_strict(cls.name), token_key(cls.name), norm_strict(base), token_key(base)}:
                if key and cls.id not in ref._class_keys[key]:
                    ref._class_keys[key].append(cls.id)

    # csg, записанные напрямую в <Plan> (встречается не во всех выгрузках)
    for csg_el in root.findall("./Plan/csg"):
        csg = Csg(id=csg_el.get("id") or "", name=csg_el.get("name") or "", tid=csg_el.get("tid") or "",
                  sid=csg_el.get("sid") or "", pclassid=csg_el.get("pclassid") or "",
                  groupid=csg_el.get("groupid") or "", parentsubjectid=csg_el.get("parentsubjectid") or "",
                  hrsweek=csg_el.get("hrsweek") or "0", studcnt=csg_el.get("studcnt") or "0",
                  order=len(ref.orphan_csg))
        ref.orphan_csg.append(csg)
        if csg.id:
            ref.plan[csg.id] = csg
            cls = ref.classes.get(csg.pclassid)
            if cls is not None:
                cls.csg.append(csg)

    # --- дни недели (эталонный порядок и wd)
    week = root.find("./TimeTable/Week")
    if week is not None:
        ref.week_id = week.get("id") or "1"
        ref.week_name = week.get("name") or ""
        for di, day_el in enumerate(week.findall("Day")):
            d = Day(id=day_el.get("id") or str(di + 1), name=day_el.get("name") or "",
                    wd=day_el.get("wd") or str(di + 2), order=di)
            ref.days.append(d)
            for key in {norm(d.name), norm_strict(d.name), DAY_ALIASES.get(norm(d.name), ""),
                        DAY_ALIASES.get(norm_strict(d.name), "")}:
                if key:
                    ref._day_keys.setdefault(key, d.id)
    if not ref.days:  # расписание в эталоне пустое — берём стандартную шестидневку
        for di, (nm, wd) in enumerate([("понедельник", "2"), ("вторник", "3"), ("среда", "4"),
                                       ("четверг", "5"), ("пятница", "6"), ("суббота", "7")]):
            d = Day(id=str(di + 1), name=nm, wd=wd, order=di)
            ref.days.append(d)
            ref._day_keys[nm] = d.id
            ref._day_keys[nm[:2]] = d.id

    # --- история кабинетов по csg (пригодится для распределения групп)
    for day_el in root.findall("./TimeTable/Week/Day"):
        for les in day_el.findall("Lesson"):
            for c in les.findall("csg"):
                rid = c.get("roomid")
                if rid and rid in ref.rooms:
                    ref.room_history[c.get("id") or ""][ref.rooms[rid].name] += 1

    ref.header_text, ref.footer_text = split_timetable(text)
    return ref


def split_timetable(text: str):
    """Разрезает исходный текст на «словари» и «расписание».

    Возвращает (header, footer), где header — всё до строки <TimeTable>
    (включая перевод строки), footer — закрывающий тег корневого элемента.
    Словари копируются в результат дословно, поэтому идентификаторы и
    форматирование эталона сохраняются байт в байт.
    """
    m = re.search(r"^[ \t]*<TimeTable[ >]", text, re.MULTILINE)
    if m:
        header = text[: m.start()]
    else:
        m2 = re.search(r"^[ \t]*</TimeTableExchange>", text, re.MULTILINE)
        header = text[: m2.start()] if m2 else text
    m3 = re.search(r"</TimeTableExchange>\s*$", text)
    footer = "</TimeTableExchange>" if not m3 else text[m3.start():].rstrip("\r\n")
    return header, footer
