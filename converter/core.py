# -*- coding: utf-8 -*-
"""Ядро конвертера: Excel-расписание -> XML формата TimeTableExchange.

Типовой вызов:

    from converter import core
    res = core.convert(core.Options(
        excel_path="расписание.xlsx",
        reference_path="ExportCm.xml",
        output_path="ExportCm_new.xml",
        aliases_path="aliases.json",
    ))
    print(res.report_text)
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import builder, excel_reader, reference, report as report_mod
from .aliases import Aliases
from .teachers import TeacherIndex
from .matching import Issue, Matcher, MatchResult, Placement

__all__ = ["Options", "ConvertResult", "convert", "load_reference", "unresolved"]


@dataclass
class Options:
    excel_path: str
    reference_path: str
    output_path: str = ""
    teachers_path: str = ""          # необязательная таблица «расписание по учителям»
    aliases_path: Optional[str] = None
    sheet: Optional[str] = None
    # неделя: '' — оставить как в эталоне, 'дд.мм.гггг' — задать понедельник недели
    week_start: str = ""
    week_name: str = ""
    week_id: str = ""
    # поведение на нераспознанных данных
    skip_unknown_subject: bool = True      # иначе — ошибка и файл не пишется
    unknown_room_mode: str = "skip"        # skip | noroom
    align_groups: bool = True              # сопоставлять группы по кабинетам из эталона
    check_conflicts: bool = True
    encoding: Optional[str] = None         # None — как в эталоне (windows-1251)
    write_output: bool = True
    strict: bool = False                   # не писать файл при наличии ошибок


@dataclass
class ConvertResult:
    ok: bool = True
    xml_text: str = ""
    output_path: str = ""
    encoding: str = "windows-1251"
    bytes_written: int = 0
    report_text: str = ""
    unresolved_csv: str = ""
    mapping_csv: str = ""
    unresolved: List[report_mod.Unresolved] = field(default_factory=list)
    issues: List[Issue] = field(default_factory=list)
    placements: List[Placement] = field(default_factory=list)
    stats: Dict[str, object] = field(default_factory=dict)
    messages: List[str] = field(default_factory=list)
    bad_chars: List[str] = field(default_factory=list)


def load_reference(path: str) -> reference.Reference:
    return reference.load(path)


def unresolved(result: MatchResult) -> List[report_mod.Unresolved]:
    return report_mod.unresolved_items(result)


def convert(opts: Options) -> ConvertResult:
    res = ConvertResult()
    log = res.messages.append

    if not os.path.exists(opts.excel_path):
        raise FileNotFoundError(f"Файл таблицы не найден: {opts.excel_path}")
    if not os.path.exists(opts.reference_path):
        raise FileNotFoundError(f"Эталонный XML не найден: {opts.reference_path}")

    log(f"Читаю эталон: {opts.reference_path}")
    ref = reference.load(opts.reference_path)
    log(f"  классов: {len(ref.classes)}, предметов: {len(ref.subjects)}, "
        f"кабинетов: {len(ref.rooms)}, учителей: {len(ref.teachers)}, "
        f"записей плана (csg): {len(ref.plan)}, звонков: {len(ref.lesson_times)}")
    if not ref.lesson_times:
        log("  ВНИМАНИЕ: в эталоне нет блока <LessonTimes> — номера уроков взять неоткуда.")

    aliases = Aliases.load(opts.aliases_path) if opts.aliases_path else Aliases()
    if opts.aliases_path and os.path.exists(opts.aliases_path):
        total = sum(len(aliases.data.get(s, {})) for s in ("subjects", "rooms", "classes", "days"))
        log(f"Файл соответствий: {opts.aliases_path} (правил: {total})")

    log(f"Читаю таблицу: {opts.excel_path}")
    sched = excel_reader.read(opts.excel_path, opts.sheet, ref.lesson_times)
    log(f"  лист «{sched.sheet}», шапка в строке {sched.header_row}, "
        f"колонок классов: {len(sched.class_cols)}, ячеек с занятиями: {len(sched.entries)}")
    for w in sched.warnings:
        log(f"  ВНИМАНИЕ: {w}")

    teacher_grid = {}
    if opts.teachers_path:
        if not os.path.exists(opts.teachers_path):
            raise FileNotFoundError(f"Файл таблицы по учителям не найден: {opts.teachers_path}")
        tsched = excel_reader.read(opts.teachers_path, opts.sheet, ref.lesson_times)
        for te in tsched.entries:
            teacher_grid[(te.day, te.period, te.excel_class)] = te.pairs
        log(f"Читаю таблицу по учителям: {opts.teachers_path}")
        log(f"  лист «{tsched.sheet}», ячеек с занятиями: {len(tsched.entries)}; "
            f"учителей в справочнике: {len(ref.teachers)}")
        for w in tsched.warnings:
            log(f"  ВНИМАНИЕ: {w}")

    matcher = Matcher(
        ref, aliases,
        skip_unknown_subject=opts.skip_unknown_subject,
        unknown_room_mode=opts.unknown_room_mode,
        align_groups=opts.align_groups,
        teacher_grid=teacher_grid,
        teacher_index=TeacherIndex(ref) if teacher_grid else None,
    )
    result = matcher.build(sched)
    if not opts.check_conflicts:
        result.issues = [i for i in result.issues
                         if i.kind not in ("conflict_teacher", "conflict_room")]
    log(f"Распознано уроков: {len(result.placements)} из ячеек {result.cells_total} "
        f"(пропущено ячеек: {result.cells_skipped})")

    # --- неделя
    week_id = opts.week_id or ref.week_id
    week_name = opts.week_name
    if not week_name and opts.week_start:
        date = builder.parse_week_start(opts.week_start)
        if date:
            week_name = builder.week_name_from_date(date)
        else:
            log(f"  ВНИМАНИЕ: не удалось разобрать дату недели «{opts.week_start}»")
    if not week_name:
        week_name = ref.week_name or "Неделя 1"
    log(f"Неделя: id={week_id}, «{week_name}»")

    # --- сборка XML
    block, notes = builder.build_timetable(ref, result.placements, week_id, week_name)
    res.xml_text = builder.assemble(ref, block)
    for n in notes:
        log(f"  {n}")

    res.issues = result.issues
    res.placements = result.placements
    res.unresolved = report_mod.unresolved_items(result)
    res.unresolved_csv = report_mod.unresolved_csv(result)
    res.mapping_csv = report_mod.mapping_csv(result)

    errors = sum(1 for i in result.issues if i.level == "error")
    warnings = sum(1 for i in result.issues if i.level == "warning")
    days_used = len({p.day.id for p in result.placements})
    classes_used = len({p.cls.id for p in result.placements})
    res.stats = report_mod.stats(result, {"days": days_used, "classes": classes_used})
    res.stats.update({"errors": errors, "warnings": warnings})

    source_info = {
        "Таблица": os.path.basename(opts.excel_path),
        "Таблица по учителям": os.path.basename(opts.teachers_path) if opts.teachers_path else "—",
        "Эталон": os.path.basename(opts.reference_path),
        "Лист": sched.sheet,
        "Неделя": week_name,
        "Кодировка результата": opts.encoding or ref.encoding,
    }
    res.report_text = report_mod.format_report(result, res.stats, source_info)

    if opts.strict and errors:
        res.ok = False
        log(f"Режим strict: ошибок {errors} — файл не записан.")
        return res

    res.encoding = opts.encoding or ref.encoding
    if opts.write_output and opts.output_path:
        folder = os.path.dirname(os.path.abspath(opts.output_path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        size, bad = builder.write(opts.output_path, res.xml_text, res.encoding)
        res.bytes_written = size
        res.output_path = opts.output_path
        res.bad_chars = bad
        log(f"Записан файл: {opts.output_path} ({size:,} байт, кодировка {res.encoding})".replace(",", " "))
        if bad:
            log(f"  ВНИМАНИЕ: символы вне кодировки {res.encoding} заменены на «?»: {' '.join(bad)}")
    else:
        log("Файл не записывался (режим проверки).")
    res.ok = True
    return res
