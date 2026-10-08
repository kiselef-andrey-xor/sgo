# -*- coding: utf-8 -*-
"""Формирование отчёта о конвертации (текст, CSV нераспознанного, CSV соответствий)."""
from __future__ import annotations

import csv
import io
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Dict, List, Optional

from .matching import Issue, MatchResult

KIND_TITLES: Dict[str, str] = {
    "subject_unknown": "Предмет не найден в справочнике системы",
    "subject_not_in_plan": "Предмет есть в справочнике, но отсутствует в учебном плане класса",
    "room_unknown": "Кабинет не найден в справочнике системы",
    "class_unknown": "Класс (колонка таблицы) не найден в справочнике системы",
    "day_unknown": "Не распознан день недели",
    "period_unknown": "Номер урока не найден среди звонков эталона",
    "group_overflow": "Не хватило групп в учебном плане",
    "class_ambiguous": "Колонка подходит нескольким классам (профильное деление)",
    "conflict_teacher": "Конфликт: учитель занят одновременно в двух местах",
    "conflict_room": "Конфликт: кабинет занят одновременно",
    "hours_mismatch": "Расхождение с учебным планом по часам",
    "ignored_subject": "Осознанно пропущено (правило в файле соответствий)",
    "ignored_room": "Кабинет пропущен по правилу в файле соответствий",
    "empty_subject": "Пустое название предмета",
    "teacher_unknown": "Учитель из таблицы по учителям не найден в справочнике системы",
    "teacher_ambiguous": "ФИО учителя допускает несколько толкований",
    "teacher_mismatch": "Учитель по таблице «по учителям» не совпадает со строкой плана",
    "grid_pairs_mismatch": "Разное число групп в ячейке двух таблиц",
    "grid_rooms_mismatch": "Кабинеты в двух таблицах различаются",
    "teacher_only_slot": "Ячейка есть только в таблице по учителям",
    "teacher_missing": "Нет данных в таблице по учителям для ячейки",
    "alias_target_missing": "Правило соответствия указывает на значение, которого нет в эталоне",
    "subject_auto_residual": "Предмет сопоставлен автоматически по остатку учебного плана",
}

# проблемы, которые чинятся добавлением соответствия
MAPPABLE = {
    "subject_unknown": "subjects",
    "subject_not_in_plan": "subjects",
    "room_unknown": "rooms",
    "class_unknown": "classes",
    "day_unknown": "days",
    "teacher_unknown": "teachers",
}

ERROR_KINDS = ("class_unknown", "day_unknown", "period_unknown", "subject_unknown",
               "subject_not_in_plan", "group_overflow")
WARNING_KINDS = ("room_unknown", "conflict_teacher", "conflict_room", "empty_subject",
                 "class_ambiguous", "teacher_unknown", "teacher_mismatch",
                 "grid_pairs_mismatch", "grid_rooms_mismatch", "alias_target_missing")
INFO_KINDS = ("ignored_subject", "ignored_room", "hours_mismatch", "teacher_ambiguous",
              "teacher_only_slot", "teacher_missing", "subject_auto_residual")


@dataclass
class Unresolved:
    section: str          # subjects | rooms | classes | days
    value: str            # значение из Excel
    count: int = 0
    suggestion: str = ""
    kind: str = ""
    example: str = ""
    classes: str = ""


def unresolved_items(result: MatchResult) -> List[Unresolved]:
    agg: Dict[tuple, Unresolved] = {}
    for issue in result.issues:
        section = MAPPABLE.get(issue.kind)
        if not section or not issue.value:
            continue
        key = (section, issue.value)
        item = agg.get(key)
        if item is None:
            item = Unresolved(section=section, value=issue.value, kind=issue.kind,
                              suggestion=issue.suggestion, example=", ".join(issue.cells[:3]))
            agg[key] = item
        item.count += issue.count
        if issue.suggestion and not item.suggestion:
            item.suggestion = issue.suggestion
        for c in issue.classes:
            if c and c not in item.classes:
                item.classes = (item.classes + ", " + c) if item.classes else c
    order = {"subjects": 0, "rooms": 1, "classes": 2, "days": 3, "teachers": 4}
    return sorted(agg.values(), key=lambda u: (order.get(u.section, 9), -u.count, u.value))


def stats(result: MatchResult, extra: Optional[Dict] = None) -> Dict[str, object]:
    by_level = Counter(i.level for i in result.issues)
    by_kind = Counter()
    for i in result.issues:
        by_kind[i.kind] += i.count
    data: Dict[str, object] = {
        "cells_total": result.cells_total,
        "cells_converted": result.cells_converted,
        "cells_skipped": result.cells_skipped,
        "placements": len(result.placements),
        "errors": by_level.get("error", 0),
        "warnings": by_level.get("warning", 0),
        "info": by_level.get("info", 0),
        "by_kind": dict(by_kind),
        "unresolved": len(unresolved_items(result)),
    }
    if extra:
        data.update(extra)
    return data


def per_day_summary(result: MatchResult) -> List[tuple]:
    counter: Dict[str, Counter] = defaultdict(Counter)
    order: Dict[str, int] = {}
    for p in result.placements:
        counter[p.day.name][p.time.number] += 1
        order[p.day.name] = p.day.order
    rows = [(day, sum(cnt.values()), dict(sorted(cnt.items(), key=lambda x: int(x[0]))))
            for day, cnt in counter.items()]
    rows.sort(key=lambda r: order.get(r[0], 99))
    return rows


def per_class_summary(result: MatchResult) -> List[tuple]:
    counter: Counter = Counter()
    for p in result.placements:
        counter[p.cls.name] += 1
    return sorted(counter.items(), key=lambda x: x[0])


def _fmt_list(items: List[str], limit: int = 8) -> str:
    if not items:
        return ""
    shown = items[:limit]
    tail = f" … (+{len(items) - limit})" if len(items) > limit else ""
    return ", ".join(shown) + tail


def format_report(result: MatchResult, extra: Optional[Dict[str, object]] = None,
                  source_info: Optional[Dict[str, str]] = None) -> str:
    out: List[str] = []
    line = "=" * 78
    out += [line, "ОТЧЁТ О КОНВЕРТАЦИИ РАСПИСАНИЯ: Excel -> TimeTableExchange XML", line]
    if source_info:
        for k, v in source_info.items():
            out.append(f"{k}: {v}")
        out.append("-" * 78)

    st = stats(result, extra)
    out.append("СВОДКА")
    out.append(f"  ячеек в таблице              : {st['cells_total']}")
    out.append(f"  преобразовано ячеек          : {st['cells_converted']}")
    out.append(f"  пропущено ячеек              : {st['cells_skipped']}")
    out.append(f"  уроков в итоговом XML (<csg>) : {st['placements']}")
    out.append(f"  ошибок / предупреждений      : {st['errors']} / {st['warnings']}")
    out.append(f"  значений требуют соответствий : {st['unresolved']}")
    if st.get("days"):
        out.append(f"  дней в расписании            : {st['days']}")
    if st.get("classes"):
        out.append(f"  классов в расписании         : {st['classes']}")

    grouped: Dict[str, List[Issue]] = defaultdict(list)
    for issue in result.issues:
        grouped[issue.kind].append(issue)

    def emit(kind: str, limit: Optional[int] = None) -> None:
        items = grouped.get(kind)
        if not items:
            return
        items = sorted(items, key=lambda i: (-i.count, i.value))
        out.append("")
        out.append(f"{KIND_TITLES.get(kind, kind).upper()} — {len(items)} поз., "
                   f"{sum(i.count for i in items)} случ.")
        for it in (items if limit is None else items[:limit]):
            out.append(f"  • «{it.value}» — {it.message}" if it.value else f"  • {it.message}")
            extra_bits = []
            if it.count > 1:
                extra_bits.append(f"повторов: {it.count}")
            if it.classes:
                extra_bits.append(f"классы: {_fmt_list(it.classes)}")
            if it.cells:
                extra_bits.append(f"ячейки: {_fmt_list(it.cells)}")
            if extra_bits:
                out.append(f"      {'; '.join(extra_bits)}")
            if it.suggestion:
                out.append(f"      похоже на: «{it.suggestion}»")
        if limit is not None and len(items) > limit:
            out.append(f"  … и ещё {len(items) - limit} (полный список — в CSV)")

    out += ["", line, "ЧТО НУЖНО ИСПРАВИТЬ (эти уроки НЕ попали в XML)", line]
    for kind in ERROR_KINDS:
        emit(kind, limit=30)
    if not any(grouped.get(k) for k in ERROR_KINDS):
        out.append("  нет — все ячейки таблицы распознаны")

    out += ["", line, "ПРЕДУПРЕЖДЕНИЯ", line]
    for kind in WARNING_KINDS:
        emit(kind, limit=30)
    if not any(grouped.get(k) for k in WARNING_KINDS):
        out.append("  нет")

    out += ["", line, "СПРАВОЧНО", line]
    for kind in INFO_KINDS:
        emit(kind, limit=20)
    if not any(grouped.get(k) for k in INFO_KINDS):
        out.append("  нет")

    out += ["", line, "ИТОГОВОЕ РАСПИСАНИЕ ПО ДНЯМ (уроков по номерам)", line]
    for day, total, cnt in per_day_summary(result):
        detail = ", ".join(f"№{k}:{v}" for k, v in cnt.items())
        out.append(f"  {day:<12} всего {total:>4}   ({detail})")

    out += ["", "ПО КЛАССАМ (число уроков)"]
    rows = per_class_summary(result)
    chunk = [f"{name}={cnt}" for name, cnt in rows]
    for i in range(0, len(chunk), 8):
        out.append("  " + "  ".join(chunk[i:i + 8]))
    out.append(line)
    return "\n".join(out)


def unresolved_csv(result: MatchResult) -> str:
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["Раздел", "Значение в Excel", "Встречается раз", "Предложение программы",
                "Классы", "Пример ячеек", "Проблема"])
    for u in unresolved_items(result):
        w.writerow([u.section, u.value, u.count, u.suggestion, u.classes, u.example,
                    KIND_TITLES.get(u.kind, u.kind)])
    return buf.getvalue()


def mapping_csv(result: MatchResult) -> str:
    """Протокол сопоставления: что во что превратилось (для проверки глазами)."""
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["Раздел", "Значение в Excel", "Значение в системе", "Способ сопоставления", "Раз"])
    order = {"класс": 0, "день": 1, "урок": 2, "предмет": 3, "кабинет": 4}
    rows = sorted(result.mapping, key=lambda r: (order.get(r.section, 9), r.excel, r.system))
    for r in rows:
        w.writerow([r.section, r.excel, r.system, r.how, r.count])
    return buf.getvalue()
