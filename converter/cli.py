# -*- coding: utf-8 -*-
"""Запуск конвертера из командной строки.

Примеры:
    python -m converter.cli "расписание.xlsx" --reference ExportCm.xml
    python -m converter.cli "расписание.xlsx" -r ExportCm.xml -o ExportCm_new.xml ^
        --aliases aliases.json --report report.txt --week 15.09.2025
    python -m converter.cli "расписание.xlsx" -r ExportCm.xml --dry-run   # только проверка
"""
from __future__ import annotations

import argparse
import os
import sys
import warnings

warnings.filterwarnings("ignore")   # openpyxl шумит на служебные колонтитулы книги

from . import core
from .aliases import Aliases


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="converter",
        description="Конвертация таблицы расписания Excel в XML формата TimeTableExchange "
                    "(импорт в систему). Идентификаторы берутся из эталонной выгрузки системы.")
    p.add_argument("excel", help="файл таблицы расписания (.xlsx/.xlsm)")
    p.add_argument("-r", "--reference", default="",
                   help="эталонный XML из системы (ExportCm.xml) — источник ID "
                        "(не нужен для --reverse, --check-reference, --init-aliases)")
    p.add_argument("-o", "--output", default="",
                   help="куда сохранить результат (по умолчанию — <имя таблицы>_ExportCm.xml)")
    p.add_argument("-t", "--teachers", default="",
                   help="дополнительная таблица «расписание по учителям» (.xlsx): даёт ФИО "
                        "и позволяет точно выбирать группу/строку плана")
    p.add_argument("-a", "--aliases", default="",
                   help="файл соответствий aliases.json (создаётся автоматически)")
    p.add_argument("--sheet", default="", help="имя листа Excel (по умолчанию — первый)")
    p.add_argument("--report", default="", help="сохранить текстовый отчёт в файл")
    p.add_argument("--unresolved", default="", help="сохранить CSV с нераспознанными значениями")
    p.add_argument("--mapping", default="", help="сохранить CSV-протокол сопоставления")
    p.add_argument("--week", default="", help="дата понедельника недели, дд.мм.гггг")
    p.add_argument("--week-name", default="", help="название недели (переопределяет --week)")
    p.add_argument("--week-id", default="", help="идентификатор недели (по умолчанию — из эталона)")
    p.add_argument("--room-mode", choices=["skip", "noroom"], default="skip",
                   help="skip — пропускать урок с неизвестным кабинетом (по умолчанию), "
                        "noroom — ставить урок без кабинета")
    p.add_argument("--no-align", action="store_true",
                   help="не подбирать группы по кабинетам из эталона")
    p.add_argument("--no-conflicts", action="store_true", help="не проверять конфликты учителей/кабинетов")
    p.add_argument("--encoding", default="", help="кодировка результата (по умолчанию — как в эталоне)")
    p.add_argument("--strict", action="store_true",
                   help="не писать XML, если есть нераспознанные предметы/классы")
    p.add_argument("--dry-run", action="store_true", help="только проверка: XML не записывается")
    p.add_argument("--reverse", action="store_true",
                   help="обратное преобразование: позиционный аргумент — XML системы, "
                        "-o — куда сохранить XLSX (листы «по предметам» и «по учителям»)")
    p.add_argument("--check-reference", action="store_true",
                   help="проверить целостность эталонного XML (дубликаты id, висячие ссылки, "
                        "конфликты слотов) и выйти")
    p.add_argument("--init-aliases", action="store_true",
                   help="создать пустой файл соответствий и выйти")
    p.add_argument("-q", "--quiet", action="store_true", help="не печатать ход работы")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

    needs_reference = not (args.reverse or args.check_reference or args.init_aliases)
    if needs_reference and not args.reference:
        print("Ошибка: укажите эталонный XML: --reference ExportCm.xml", file=sys.stderr)
        return 2

    if args.init_aliases:
        path = args.aliases or "aliases.json"
        Aliases.default_file(path)
        print(f"Создан файл соответствий: {path}")
        return 0

    if args.reverse:
        from . import reverse
        try:
            res = reverse.convert(reverse.ReverseOptions(
                xml_path=args.excel,
                output_path=args.output,
                teachers_sheet=True,
                summary_sheet=True,
            ))
        except Exception as exc:  # noqa: BLE001
            print(f"Ошибка обратного преобразования: {exc}", file=sys.stderr)
            return 3
        if not args.quiet:
            for line in res.messages:
                print(line)
        print(f"\nИТОГ: дней {res.days}, номеров уроков {res.periods}, классов {res.classes}, "
              f"заполненных ячеек {res.cells} -> {res.output_path}")
        return 0

    if args.check_reference:
        from . import validate
        if not os.path.exists(args.reference):
            print(f"Ошибка: эталонный XML не найден: {args.reference}", file=sys.stderr)
            return 2
        ref = core.load_reference(args.reference)
        findings = validate.validate(ref)
        print(validate.format_findings(ref, findings))
        return 1 if any(f.level == "error" for f in findings) else 0

    if not os.path.exists(args.excel):
        print(f"Ошибка: файл таблицы не найден: {args.excel}", file=sys.stderr)
        return 2
    if not os.path.exists(args.reference):
        print(f"Ошибка: эталонный XML не найден: {args.reference}", file=sys.stderr)
        return 2

    output = args.output
    if not output and not args.dry_run:
        base = os.path.splitext(os.path.basename(args.excel))[0]
        refname = os.path.basename(args.reference) or "ExportCm.xml"
        output = os.path.join(os.path.dirname(os.path.abspath(args.excel)),
                              f"{base}_{refname}")

    aliases_path = args.aliases or os.path.join(os.path.dirname(os.path.abspath(args.reference)),
                                                "aliases.json")
    Aliases.default_file(aliases_path)

    opts = core.Options(
        excel_path=args.excel,
        reference_path=args.reference,
        output_path=output,
        teachers_path=args.teachers,
        aliases_path=aliases_path,
        sheet=args.sheet or None,
        week_start=args.week,
        week_name=args.week_name,
        week_id=args.week_id,
        unknown_room_mode=args.room_mode,
        align_groups=not args.no_align,
        check_conflicts=not args.no_conflicts,
        encoding=args.encoding or None,
        write_output=not args.dry_run,
        strict=args.strict,
    )

    try:
        res = core.convert(opts)
    except Exception as exc:  # noqa: BLE001 — пользователю нужен понятный текст
        print(f"Ошибка конвертации: {exc}", file=sys.stderr)
        return 3

    if not args.quiet:
        for line in res.messages:
            print(line)
        print()
        print(res.report_text)

    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            fh.write(res.report_text)
        print(f"Отчёт сохранён: {args.report}")
    if args.unresolved:
        with open(args.unresolved, "w", encoding="utf-8-sig", newline="") as fh:
            fh.write(res.unresolved_csv)
        print(f"Список нераспознанного сохранён: {args.unresolved}")
    if args.mapping:
        with open(args.mapping, "w", encoding="utf-8-sig", newline="") as fh:
            fh.write(res.mapping_csv)
        print(f"Протокол сопоставления сохранён: {args.mapping}")

    errors = int(res.stats.get("errors", 0))
    warnings = int(res.stats.get("warnings", 0))
    print(f"\nИТОГ: уроков в XML — {res.stats.get('placements')}, "
          f"нераспознанных значений — {res.stats.get('unresolved')}, "
          f"ошибок — {errors}, предупреждений — {warnings}")
    if args.strict and errors:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
