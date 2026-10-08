# -*- coding: utf-8 -*-
"""Сопоставление данных Excel со справочниками системы и построение размещений.

Результат работы — список Placement (уроки, которые попадут в XML), список
Issue (всё нераспознанное/сомнительное) и протокол сопоставления
(что именно во что превратилось — для проверки человеком).
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .aliases import Aliases
from .excel_reader import Entry, ExcelSchedule, Pair
from .reference import Csg, Day, LessonTime, Reference, Room, SchoolClass
from .teachers import TeacherIndex
from .text_utils import best_match, norm, norm_strict, token_key

# Порог нечёткого сопоставления названий предметов
FUZZY_THRESHOLD = 0.78
MAX_CELLS_IN_ISSUE = 12
MAX_CLASSES_IN_ISSUE = 12

SPLIT_GROUP_RE = re.compile(
    r"(?:п\.?\s*\d|подгруп|подгр|\d\s*-?\s*гр|группа\s*\d|вариант|\(\s*\d\s*\))", re.IGNORECASE)


# --------------------------------------------------------------------------- #
# структуры данных
# --------------------------------------------------------------------------- #
@dataclass
class Placement:
    """Один урок в итоговом расписании (одна строка <csg> внутри <Lesson>)."""
    day: Day
    time: LessonTime
    period: int
    cls: SchoolClass
    csg: Csg
    room: Optional[Room]
    cell: str
    subject_raw: str
    room_raw: str
    matched_by: str = ""
    teacher_text: str = ""
    teacher_tid: str = ""

    @property
    def tid(self) -> str:
        return self.csg.tid or ""

    def sort_key(self):
        return (self.day.order, int(self.time.id), self.cls.order, self.csg.order)


@dataclass
class Issue:
    kind: str
    level: str                 # error | warning | info
    message: str
    value: str = ""            # значение из Excel, с которым проблема
    section: str = ""          # subjects | rooms | classes | days
    suggestion: str = ""       # что предложила программа
    count: int = 1
    cells: List[str] = field(default_factory=list)
    classes: List[str] = field(default_factory=list)


@dataclass
class MappingRow:
    section: str               # предмет / кабинет / класс / день / урок
    excel: str
    system: str
    how: str
    count: int = 0


@dataclass
class MatchResult:
    placements: List[Placement] = field(default_factory=list)
    issues: List[Issue] = field(default_factory=list)
    used_hours: Dict[str, int] = field(default_factory=dict)   # csg id -> сколько уроков поставлено
    mapping: List[MappingRow] = field(default_factory=list)
    cells_total: int = 0
    cells_converted: int = 0
    cells_skipped: int = 0


def group_rank(csg: Csg, cls: SchoolClass) -> Tuple[int, int]:
    """Приоритет записи плана при выборе группы (чем меньше, тем лучше).

    0 — обычный урок всего класса (метка группы пуста или равна названию класса);
    1 — деление на подгруппы («п.1», «подгруппа 2», «1 гр.»);
    2 — именные/индивидуальные занятия («Русский язык/Галкин»).
    Второй ключ — число учеников (больше = приоритетнее).
    """
    label = norm(csg.group_label)
    cname = norm(cls.name)
    if not label or label == cname or (cname and cname in label):
        first = 0
    elif SPLIT_GROUP_RE.search(label):
        first = 1
    else:
        first = 2
    return (first, -csg.students)


# --------------------------------------------------------------------------- #
class Matcher:
    def __init__(self, ref: Reference, aliases: Aliases,
                 skip_unknown_subject: bool = True,
                 unknown_room_mode: str = "skip",      # skip | noroom
                 align_groups: bool = True,
                 teacher_grid=None,                     # (день, номер, класс) -> пары «ФИО (кабинет)»
                 teacher_index: Optional[TeacherIndex] = None) -> None:
        self.ref = ref
        self.aliases = aliases
        self.skip_unknown_subject = skip_unknown_subject
        self.unknown_room_mode = unknown_room_mode
        self.align_groups = align_groups
        self.teacher_grid = teacher_grid or {}
        self.teacher_index = teacher_index

        self.result = MatchResult()
        self._used: Dict[Tuple[str, str], set] = defaultdict(set)   # (day, time) -> {csg id}
        self._issues: Dict[Tuple[str, str, str, str], Issue] = {}
        self._map: Dict[Tuple[str, str, str, str], MappingRow] = {}
        self._cls_index: Dict[str, Dict[str, List[Csg]]] = {}
        self._class_cache: Dict[str, List[SchoolClass]] = {}
        self._room_cache: Dict[str, Tuple[Optional[Room], str, str]] = {}
        self._subject_guess: Dict[str, str] = {}
        # автоподбор: (class_id, norm-предмет из Excel) -> название из плана класса
        self._auto_residual: Dict[Tuple[str, str], str] = {}
        # подсказка: единственный неиспользованный предмет плана класса
        self._residual_single: Dict[str, str] = {}

    # ------------------------------------------------------------------ #
    # служебное
    # ------------------------------------------------------------------ #
    def _index_for(self, cls: SchoolClass) -> Dict[str, List[Csg]]:
        idx = self._cls_index.get(cls.id)
        if idx is not None:
            return idx
        idx = defaultdict(list)
        for csg in cls.csg:
            keys = {norm_strict(csg.base_name), token_key(csg.base_name)}
            subj = self.ref.subjects.get(csg.sid)
            if subj:
                keys |= {norm_strict(subj.name), token_key(subj.name),
                         norm_strict(subj.abbr), token_key(subj.abbr)}
            for key in filter(None, keys):
                idx[key].append(csg)
        for key in list(idx):
            seen, ordered = set(), []
            for csg in sorted(idx[key], key=lambda c: c.order):
                if csg.id not in seen:
                    seen.add(csg.id)
                    ordered.append(csg)
            idx[key] = ordered
        self._cls_index[cls.id] = dict(idx)
        return self._cls_index[cls.id]

    def add_issue(self, kind: str, level: str, message: str, value: str = "", section: str = "",
                  suggestion: str = "", cell: str = "", cls_name: str = "") -> None:
        key = (kind, section, value, message)
        issue = self._issues.get(key)
        if issue is None:
            issue = Issue(kind=kind, level=level, message=message, value=value, section=section,
                          suggestion=suggestion, count=0)
            self._issues[key] = issue
            self.result.issues.append(issue)
        issue.count += 1
        if suggestion and not issue.suggestion:
            issue.suggestion = suggestion
        if cell and cell not in issue.cells and len(issue.cells) < MAX_CELLS_IN_ISSUE:
            issue.cells.append(cell)
        if cls_name and cls_name not in issue.classes and len(issue.classes) < MAX_CLASSES_IN_ISSUE:
            issue.classes.append(cls_name)

    def note_mapping(self, section: str, excel: str, system: str, how: str) -> None:
        if not excel:
            return
        key = (section, norm_strict(excel), system, how)
        row = self._map.get(key)
        if row is None:
            row = MappingRow(section=section, excel=excel, system=system, how=how)
            self._map[key] = row
            self.result.mapping.append(row)
        row.count += 1

    # ------------------------------------------------------------------ #
    # дни, звонки, классы
    # ------------------------------------------------------------------ #
    def resolve_day(self, entry: Entry) -> Optional[Day]:
        text = entry.day_raw or entry.day
        if not text:
            return None
        if self.aliases.has("days", text):
            value = self.aliases.get("days", text)
            if Aliases.is_ignore(value):
                return None
            text = str(value)
        day = self.ref.day_by_text(text) or self.ref.day_by_text(entry.day)
        if day is None:
            guess, _, _ = best_match(text, [d.name for d in self.ref.days], threshold=0.6)
            day = next((d for d in self.ref.days if d.name == guess), None) if guess else None
        return day

    def resolve_time(self, entry: Entry) -> Optional[LessonTime]:
        if entry.period is not None:
            lt = self.ref.time_by_number(entry.period)
            if lt:
                return lt
        if entry.time_text:
            lt = self.ref.time_by_start(entry.time_text)
            if lt:
                return lt
        return None

    def resolve_classes(self, excel_class: str) -> Tuple[List[SchoolClass], str]:
        if excel_class in self._class_cache:
            return self._class_cache[excel_class], ""
        classes: List[SchoolClass] = []
        how = ""
        rule = self.aliases.get("classes", excel_class)
        names: List[str] = []
        if isinstance(rule, dict):
            names = [str(x) for x in (rule.get("candidates") or [])]
        elif isinstance(rule, list):
            names = [str(x) for x in rule]
        elif isinstance(rule, str) and rule.strip() and not Aliases.is_ignore(rule):
            names = [rule]
        for nm in names:
            for c in self.ref.classes_by_text(nm):
                if c not in classes:
                    classes.append(c)
            how = "соответствие"
        if not classes:
            classes = self.ref.classes_by_text(excel_class)
            how = "точное" if classes else ""
        if not classes:
            guess, score, _ = best_match(excel_class, self.ref.class_names, threshold=0.85)
            if guess:
                classes = self.ref.classes_by_text(guess)
                how = f"нечёткое ({score:.2f})"
        self._class_cache[excel_class] = classes
        if classes:
            self.note_mapping("класс", excel_class, ", ".join(c.name for c in classes), how)
        return classes, how

    def class_split_rule(self, excel_class: str) -> Dict[str, str]:
        rule = self.aliases.get("classes", excel_class)
        if isinstance(rule, dict) and isinstance(rule.get("by_subject"), dict):
            return {norm_strict(k): str(v) for k, v in rule["by_subject"].items()}
        return {}

    # ------------------------------------------------------------------ #
    # предметы
    # ------------------------------------------------------------------ #
    def class_subject_alias(self, cls: SchoolClass, text: str):
        """Персональное правило для класса: aliases.json → class_subjects."""
        section = self.aliases.data.get("class_subjects") or {}
        for key, rules in section.items():
            if norm_strict(key) == norm_strict(cls.name) and isinstance(rules, dict):
                for src, dst in rules.items():
                    if norm_strict(src) == norm_strict(text):
                        return dst, True
        return None, False

    def find_csg(self, cls: SchoolClass, text: str) -> Tuple[List[Csg], str, Optional[str]]:
        """Ищет записи плана класса по названию предмета.

        -> (кандидаты, способ совпадения, sid найденного предмета справочника)
        """
        key, tkey = norm_strict(text), token_key(text)
        idx = self._index_for(cls)
        for k, how in ((key, "точное"), (tkey, "перестановка слов")):
            if k and idx.get(k):
                return self._rank(list(idx[k]), cls), how, self._sid_of(idx[k])
        subj = self.ref.subject_by_text(text)
        how_subj = "по справочнику предметов"
        if subj is None:
            guess, score, ambiguous = best_match(text, self.ref.subject_names, threshold=FUZZY_THRESHOLD)
            if guess:
                subj = self.ref.subject_by_text(guess)
                how_subj = f"по справочнику (нечётко {score:.2f})"
        if subj is not None:
            by_sid = [c for c in cls.csg if c.sid == subj.sid]
            if by_sid:
                return self._rank(by_sid, cls), how_subj, subj.sid
            return [], "нет в плане класса", subj.sid
        names = sorted({c.base_name for c in cls.csg})
        guess, score, ambiguous = best_match(text, names, threshold=FUZZY_THRESHOLD + 0.04)
        if guess and not ambiguous:
            by_name = [c for c in cls.csg if norm_strict(c.base_name) == norm_strict(guess)]
            if by_name:
                return self._rank(by_name, cls), f"нечёткое ({score:.2f})", None
        return [], "", None

    def _sid_of(self, csgs: Sequence[Csg]) -> Optional[str]:
        sids = {c.sid for c in csgs if c.sid}
        return next(iter(sids)) if len(sids) == 1 else None

    def _rank(self, csgs: Sequence[Csg], cls: SchoolClass) -> List[Csg]:
        return sorted(csgs, key=lambda c: (group_rank(c, cls), c.order))

    def resolve_teacher(self, text: str) -> Tuple[List[str], str]:
        """ФИО из таблицы по учителям -> список tid (с учётом aliases → teachers)."""
        if self.teacher_index is None or not text:
            return [], ""
        if self.aliases.has("teachers", text):
            value = self.aliases.get("teachers", text)
            if Aliases.is_ignore(value):
                return [], "ignore"
            tids = self.teacher_index.tids_by_text(str(value))
            if not tids and str(value) in self.ref.teachers:
                tids = [str(value)]
            if not tids:
                self.add_issue(
                    "alias_target_missing", "warning",
                    f"Правило соответствия «{text} → {value}» не сработало: учителя «{value}» "
                    f"нет в справочнике эталона",
                    value=f"{text} -> {value}", section="teachers")
            return (tids, "соответствие") if tids else ([], "")
        return self.teacher_index.resolve(text)

    def suggest_teacher(self, text: str) -> str:
        guess, _, _ = best_match(text, self.teacher_index.fio_list if self.teacher_index else [],
                                 threshold=0.6)
        return guess or ""

    def suggest_subject(self, text: str) -> str:
        key = norm_strict(text)
        if key in self._subject_guess:
            return self._subject_guess[key]
        guess, _, _ = best_match(text, self.ref.subject_names, threshold=0.5)
        self._subject_guess[key] = guess or ""
        return guess or ""

    # ------------------------------------------------------------------ #
    # кабинеты
    # ------------------------------------------------------------------ #
    def resolve_room(self, text: str) -> Tuple[Optional[Room], bool, str, str]:
        """-> (кабинет, пропускать?, способ совпадения, подсказка)."""
        if not text:
            return None, False, "", ""
        cached = self._room_cache.get(norm_strict(text))
        if cached is not None:
            room, how, suggestion = cached
            return room, how == "__ignore__", how if how != "__ignore__" else "", suggestion
        room, how, suggestion = None, "", ""
        if not self.aliases.has("rooms", text):
            room = self.ref.room_by_text(text)
            how = "точное" if room else ""
            if room is None:
                guess, score, _ = best_match(text, self.ref.room_names, threshold=0.86)
                if guess:
                    room = self.ref.room_by_text(guess)
                    how = f"нечёткое ({score:.2f})"
        else:
            value = self.aliases.get("rooms", text)
            if Aliases.is_ignore(value):
                self._room_cache[norm_strict(text)] = (None, "__ignore__", "")
                return None, True, "", ""
            room = self.ref.room_by_text(str(value))
            how = "соответствие" if room else ""
            if room is None:
                self.add_issue(
                    "alias_target_missing", "warning",
                    f"Правило соответствия «{text} → {value}» не сработало: кабинета "
                    f"«{value}» нет в справочнике эталона. Добавьте кабинет в системе и "
                    f"выгрузите новый эталон либо исправьте название в правиле",
                    value=f"{text} -> {value}", section="rooms")
        if room is None:
            guess, _, _ = best_match(text, self.ref.room_names, threshold=0.5)
            suggestion = guess or ""
        self._room_cache[norm_strict(text)] = (room, how, suggestion)
        return room, False, how, suggestion

    # ------------------------------------------------------------------ #
    # главный проход
    # ------------------------------------------------------------------ #
    def build(self, sched: ExcelSchedule) -> MatchResult:
        self._prescan_residuals(sched)
        if self.teacher_grid:
            self._cross_check_grids(sched)
        for entry in sched.entries:
            self.result.cells_total += 1
            self._process_entry(entry)
        self._check_conflicts()
        self._check_hours()
        return self.result

    def _prescan_residuals(self, sched: ExcelSchedule) -> None:
        """Автоподбор «по остатку плана».

        Если в таблице у класса есть предмет, которого нет в плане (например,
        «Информатика» в 5–6 классах), и при этом в плане класса ровно один предмет,
        не покрытый таблицей (например, «Программирование и алгоритмика»), а в таблице
        ровно один такой «лишний» предмет — считаем их одним и тем же курсом.
        """
        matched_sids: Dict[str, set] = defaultdict(set)
        unmatched: Dict[str, set] = defaultdict(set)
        unmatched_tids: Dict[Tuple[str, str], set] = defaultdict(set)
        for entry in sched.entries:
            classes, _ = self.resolve_classes(entry.excel_class)
            tpairs = self.teacher_grid.get((entry.day, entry.period, entry.excel_class))
            for cls in classes:
                for pos, pair in enumerate(entry.pairs):
                    text, ignore = pair.subject, False
                    alias, has_alias = self.class_subject_alias(cls, pair.subject)
                    if has_alias:
                        if Aliases.is_ignore(alias):
                            continue
                        text = str(alias)
                    elif self.aliases.has("subjects", pair.subject):
                        value = self.aliases.get("subjects", pair.subject)
                        if Aliases.is_ignore(value):
                            continue
                    csgs, _how, _sid = self.find_csg(cls, text)
                    if not csgs and not has_alias and self.aliases.has("subjects", pair.subject):
                        value = self.aliases.get("subjects", pair.subject)
                        if not Aliases.is_ignore(value):
                            csgs, _how, _sid = self.find_csg(cls, str(value))
                            if csgs:
                                text = str(value)
                    if csgs:
                        matched_sids[cls.id] |= {c.sid for c in csgs if c.sid}
                    elif self.ref.subject_by_text(text) is not None:
                        key = norm_strict(text)
                        unmatched[cls.id].add(key)
                        if self.teacher_index is not None and tpairs and pos < len(tpairs):
                            tids, _how = self.teacher_index.resolve(tpairs[pos].subject)
                            unmatched_tids[(cls.id, key)] |= set(tids)
        for cls in self.ref.classes.values():
            resid = {c.base_name for c in cls.csg
                     if c.sid and c.sid not in matched_sids.get(cls.id, ())}
            if len(resid) == 1:
                self._residual_single[cls.id] = next(iter(resid))
            un = unmatched.get(cls.id, set())
            if len(un) == 1 and len(resid) == 1:
                subj_name = next(iter(un))
                resid_name = next(iter(resid)
                                   )
                # надёжность: учитель из таблицы «по учителям» действительно ведёт
                # остаточный предмет плана — иначе это два разных курса и автоподбор запрещён
                tids = unmatched_tids.get((cls.id, subj_name), set())
                resid_tids = {c.tid for c in cls.csg if c.base_name == resid_name and c.tid}
                if tids and (tids & resid_tids):
                    self._auto_residual[(cls.id, subj_name)] = resid_name

    def _cross_check_grids(self, sched: ExcelSchedule) -> None:
        """Сверка таблицы «по предметам» с таблицей «по учителям»."""
        subj_keys = {(e.day, e.period, e.excel_class): e for e in sched.entries}
        for key, tentry in self.teacher_grid_items():
            sentry = subj_keys.get(key)
            if sentry is None:
                self.add_issue("teacher_only_slot", "info",
                               "Ячейка есть только в таблице по учителям (в таблицу по предметам "
                               "не попала)", value=f"{key[2]} {key[0]} {key[1]}",
                               cell=tentry[0].cell if tentry else "")
                continue
            if len(tentry) != len(sentry.pairs):
                self.add_issue("grid_pairs_mismatch", "warning",
                               "В двух таблицах разное число групп в одной ячейке",
                               value=f"{key[2]} {key[0]} {key[1]}", cell=sentry.pairs[0].cell)
                continue
            for sp, tp in zip(sentry.pairs, tentry):
                if norm(sp.room) != norm(tp.room):
                    self.add_issue("grid_rooms_mismatch", "warning",
                                   "Кабинеты в таблице по предметам и по учителям различаются "
                                   "(в XML уходит вариант из таблицы по предметам)",
                                   value=f"{sp.room} / {tp.room}", cell=sp.cell)
        for key in subj_keys:
            if self.teacher_grid and key not in self.teacher_grid:
                self.add_issue("teacher_missing", "info",
                               "Для ячейки нет данных в таблице по учителям",
                               value=key[2], cell=subj_keys[key].pairs[0].cell)

    def teacher_grid_items(self):
        return self.teacher_grid.items()

    def _process_entry(self, entry: Entry) -> None:
        day = self.resolve_day(entry)
        cell = f"{entry.excel_class}:{entry.day_raw}{entry.period_raw}"
        if day is None:
            self.add_issue("day_unknown", "error", "День недели не распознан",
                           value=entry.day_raw, section="days", cell=cell)
            self.result.cells_skipped += 1
            return
        lt = self.resolve_time(entry)
        if lt is None:
            numbers = ", ".join(t.number for t in self.ref.lesson_times) or "—"
            self.add_issue("period_unknown", "error",
                           f"Номер урока не найден среди звонков эталона (доступны: {numbers})",
                           value=entry.period_raw or entry.time_text or "", section="days", cell=cell)
            self.result.cells_skipped += 1
            return
        self.note_mapping("день", entry.day_raw or entry.day, day.name, "")
        self.note_mapping("урок", str(entry.period_raw), f"№{lt.number} ({lt.starttime}–{lt.endtime})", "")
        classes, how_classes = self.resolve_classes(entry.excel_class)
        if not classes:
            self.add_issue("class_unknown", "error", "Класс (колонка таблицы) не найден в справочнике",
                           value=entry.excel_class, section="classes", cell=cell,
                           suggestion=best_match(entry.excel_class, self.ref.class_names, threshold=0.5)[0] or "")
            self.result.cells_skipped += 1
            return

        tpairs = self.teacher_grid.get((entry.day, entry.period, entry.excel_class))
        assignments = self._assign_to_classes(entry, classes, day, lt)
        placed_any = False
        for pos, (cls, pair) in enumerate(assignments):
            teacher_text = ""
            if tpairs and pos < len(tpairs):
                teacher_text = tpairs[pos].subject
            if self._place(cls, pair, day, lt, how_classes, teacher_text):
                placed_any = True
        if placed_any:
            self.result.cells_converted += 1
        else:
            self.result.cells_skipped += 1

    def _assign_to_classes(self, entry: Entry, classes: List[SchoolClass], day: Day,
                           lt: LessonTime) -> List[Tuple[SchoolClass, Pair]]:
        """Раскладывает пары из ячейки по классам (актуально для профильных классов)."""
        if len(classes) == 1:
            return [(classes[0], p) for p in entry.pairs]
        split = self.class_split_rule(entry.excel_class)
        slot = (day.id, lt.id)
        out: List[Tuple[SchoolClass, Pair]] = []
        taken: set = set()
        for pair in entry.pairs:
            target: Optional[SchoolClass] = None
            subject_text = pair.subject
            for candidate in classes:
                alias, has_alias = self.class_subject_alias(candidate, subject_text)
                if has_alias and not Aliases.is_ignore(alias):
                    subject_text = str(alias)
                    break
            key = norm_strict(subject_text)
            if key in split:
                found = self.ref.classes_by_text(split[key])
                target = next((c for c in found if c in classes and c.id not in taken), None)
            if target is None:
                scored = []
                for cls in classes:
                    if cls.id in taken:
                        continue
                    csgs, _, _ = self.find_csg(cls, subject_text)
                    available = [c for c in csgs if c.id not in self._used[slot]]
                    if available:
                        scored.append((-max(c.hours for c in available), cls.order, cls))
                if scored:
                    scored.sort(key=lambda x: (x[0], x[1]))
                    target = scored[0][2]
                    if len(scored) > 1:
                        self.add_issue(
                            "class_ambiguous", "info",
                            "Колонка подходит нескольким классам — урок отнесён к классу с "
                            "большим числом часов в плане",
                            value=f"{entry.excel_class} / {subject_text}", section="classes",
                            cell=f"{entry.excel_class}:{entry.day_raw}{entry.period_raw}",
                            cls_name=target.name,
                            suggestion=", ".join(c.name for c in classes))
            if target is None:
                target = next((c for c in classes if c.id not in taken), classes[0])
            taken.add(target.id)
            out.append((target, pair))
        return out

    def _place(self, cls: SchoolClass, pair: Pair, day: Day, lt: LessonTime, how_classes: str,
               teacher_text: str = "") -> bool:
        slot = (day.id, lt.id)
        cell = pair.cell
        subject_text, how_subject = pair.subject, ""

        # 1) персональное правило для класса
        alias, has_alias = self.class_subject_alias(cls, subject_text)
        if has_alias:
            if Aliases.is_ignore(alias):
                self.add_issue("ignored_subject", "info",
                               "Пропущено по правилу для класса в файле соответствий",
                               value=pair.subject, section="subjects", cell=cell, cls_name=cls.name)
                return False
            subject_text, how_subject = str(alias), "правило для класса"

        candidates, how, sid = self.find_csg(cls, subject_text)

        # 2) общее соответствие (aliases.json → subjects) — как запасной вариант
        if not candidates and self.aliases.has("subjects", pair.subject) and not how_subject:
            value = self.aliases.get("subjects", pair.subject)
            if Aliases.is_ignore(value):
                self.add_issue("ignored_subject", "info",
                               "Пропущено по правилу в файле соответствий",
                               value=pair.subject, section="subjects", cell=cell, cls_name=cls.name)
                return False
            subject_text, how_subject = str(value), "соответствие"
            candidates, how, sid = self.find_csg(cls, subject_text)
            if not candidates:
                self.add_issue(
                    "alias_target_missing", "warning",
                    f"Правило соответствия «{pair.subject} → {value}» не сработало: предмета "
                    f"«{value}» нет в справочнике/плане эталона. Добавьте предмет в системе и "
                    f"выгрузите новый эталон либо исправьте название в правиле",
                    value=f"{pair.subject} -> {value}", section="subjects", cls_name=cls.name)

        if not candidates:
            auto = self._auto_residual.get((cls.id, norm_strict(pair.subject)))
            if auto:
                candidates, how, sid = self.find_csg(cls, auto)
                if candidates:
                    how_subject = f"автоподбор по плану класса ({auto})"
                    self.add_issue("subject_auto_residual", "info",
                                   f"Предмет «{pair.subject}» сопоставлен единственному непокрытому "
                                   f"предмету плана класса: «{auto}»",
                                   value=f"{pair.subject} -> {auto}", section="subjects",
                                   cell=cell, cls_name=cls.name)
        if not candidates:
            subj = self.ref.subjects.get(sid) if sid else None
            if subj is not None:
                residual = self._residual_single.get(cls.id, "")
                self.add_issue("subject_not_in_plan", "error",
                               "Предмет есть в справочнике, но отсутствует в учебном плане класса",
                               value=pair.subject, section="subjects", cell=cell, cls_name=cls.name,
                               suggestion=residual or subj.name)
            else:
                self.add_issue("subject_unknown", "error",
                               "Предмет не найден в справочнике системы",
                               value=pair.subject, section="subjects", cell=cell, cls_name=cls.name,
                               suggestion=self.suggest_subject(pair.subject))
            return False

        available = [c for c in candidates if c.id not in self._used[slot]]
        if not available:
            self.add_issue("group_overflow", "warning",
                           "В плане класса нет свободной группы на этот слот (в таблице больше "
                           "параллельных занятий, чем групп)",
                           value=f"{cls.name} / {pair.subject}", section="subjects", cell=cell,
                           cls_name=cls.name)
            return False

        room, ignore_room, how_room, room_suggestion = self.resolve_room(pair.room)
        if ignore_room:
            self.add_issue("ignored_room", "info",
                           "Кабинет пропущен по правилу в файле соответствий (урок без кабинета)",
                           value=pair.room, section="rooms", cell=cell, cls_name=cls.name)
        elif pair.room and room is None:
            self.add_issue("room_unknown", "warning", "Кабинет не найден в справочнике системы",
                           value=pair.room, section="rooms", cell=cell, cls_name=cls.name,
                           suggestion=room_suggestion)
            if self.unknown_room_mode == "skip":
                return False

        teacher_note = ""
        if teacher_text and self.teacher_index is not None:
            tids, how_t = self.resolve_teacher(teacher_text)
            if tids == [] and how_t != "ignore":
                self.add_issue("teacher_unknown", "warning",
                               "Учитель из таблицы по учителям не найден в справочнике системы",
                               value=teacher_text, section="teachers", cell=cell, cls_name=cls.name,
                               suggestion=self.suggest_teacher(teacher_text))
            elif tids:
                pref = [c for c in available if c.tid in tids]
                if pref:
                    available = pref
                    teacher_note = how_t
                elif len(tids) > 1:
                    self.add_issue("teacher_ambiguous", "info",
                                   "ФИО допускает нескольких учителей, план не позволяет выбрать",
                                   value=teacher_text, section="teachers", cell=cell, cls_name=cls.name,
                                   suggestion="; ".join(self.teacher_index.display(t) for t in tids))
                else:
                    plan_teachers = "; ".join(sorted({self.ref.teacher_fio(c.tid) for c in available}))
                    self.add_issue("teacher_mismatch", "warning",
                                   "Учитель по таблице «по учителям» не ведёт эту строку плана — "
                                   "взята строка плана, проверьте распределение",
                                   value=teacher_text, section="teachers", cell=cell, cls_name=cls.name,
                                   suggestion=plan_teachers)

        csg = self._pick_csg(available, room)
        if teacher_text and self.teacher_index is not None:
            self.note_mapping("учитель", teacher_text, self.teacher_index.display(csg.tid),
                              teacher_note or "по плану")
        self._used[slot].add(csg.id)
        self.result.used_hours[csg.id] = self.result.used_hours.get(csg.id, 0) + 1
        self.note_mapping("предмет", pair.subject, csg.base_name + (f" [{csg.group_label}]" if csg.group_label else ""),
                          how_subject or how)
        if pair.room:
            self.note_mapping("кабинет", pair.room, room.name if room else "— не найден —", how_room)
        try:
            period_no = int(lt.number)
        except (TypeError, ValueError):
            period_no = 0
        self.result.placements.append(Placement(
            day=day, time=lt,
            period=period_no,
            cls=cls, csg=csg, room=room, cell=cell,
            subject_raw=pair.subject, room_raw=pair.room,
            matched_by=";".join(x for x in (how_subject or how, how_room, how_classes) if x),
            teacher_text=teacher_text, teacher_tid=csg.tid,
        ))
        return True

    def _pick_csg(self, available: List[Csg], room: Optional[Room]) -> Csg:
        """Выбирает группу; если кабинет «исторически» закреплён за группой — учитываем."""
        if len(available) == 1 or room is None or not self.align_groups:
            return available[0]
        best, best_score = available[0], -1
        for csg in available:
            score = self.ref.room_history.get(csg.id, Counter()).get(room.name, 0)
            if score > best_score:
                best, best_score = csg, score
        return best

    # ------------------------------------------------------------------ #
    # проверки
    # ------------------------------------------------------------------ #
    def _check_conflicts(self) -> None:
        by_teacher: Dict[Tuple[str, str, str], List[Placement]] = defaultdict(list)
        by_room: Dict[Tuple[str, str, str], List[Placement]] = defaultdict(list)
        for p in self.result.placements:
            if p.tid:
                by_teacher[(p.day.id, p.time.id, p.tid)].append(p)
            if p.room:
                by_room[(p.day.id, p.time.id, p.room.id)].append(p)
        for (day_id, time_id, tid), items in sorted(by_teacher.items()):
            if len(items) < 2:
                continue
            day = next((d.name for d in self.ref.days if d.id == day_id), day_id)
            lt = items[0].time
            details = ", ".join(sorted({f"{i.cls.name}—{i.csg.base_name}" for i in items}))
            self.add_issue("conflict_teacher", "warning",
                           f"Учитель {self.ref.teacher_fio(tid)} занят одновременно "
                           f"({day}, урок {lt.number}): {details}",
                           value=f"{self.ref.teacher_fio(tid)}|{day}|{lt.number}",
                           cell=items[0].cell, suggestion=self.ref.teacher_fio(tid))
        for (day_id, time_id, rid), items in sorted(by_room.items()):
            if len({i.csg.id for i in items}) < 2:
                continue
            room = self.ref.rooms[rid]
            day = next((d.name for d in self.ref.days if d.id == day_id), day_id)
            lt = items[0].time
            details = ", ".join(sorted({f"{i.cls.name}—{i.csg.base_name}" for i in items}))
            self.add_issue("conflict_room", "warning",
                           f"Кабинет {room.name} занят одновременно ({day}, урок {lt.number}): {details}",
                           value=f"{room.name}|{day}|{lt.number}", cell=items[0].cell,
                           suggestion=room.name)

    def _check_hours(self) -> None:
        for csg_id, count in sorted(self.result.used_hours.items()):
            csg = self.ref.plan.get(csg_id)
            if not csg or not csg.hours or count == csg.hours:
                continue
            cls = self.ref.classes.get(csg.owner_classid) or self.ref.classes.get(csg.pclassid)
            cls_name = cls.name if cls else "?"
            self.add_issue("hours_mismatch", "info",
                           f"«{csg.name}»: в плане {csg.hours} ч/нед, в расписании {count}",
                           value=f"{cls_name}|{csg.name}", cls_name=cls_name)
