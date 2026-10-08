# -*- coding: utf-8 -*-
"""Нормализация и нечёткое сопоставление текстовых названий.

Названия предметов/кабинетов/классов в Excel и в выгрузке системы почти
никогда не совпадают посимвольно ("Иностранный язык (английский)" против
"Иностранный (английский) язык"), поэтому используется несколько уровней
сравнения: точное совпадение нормализованной строки -> совпадение набора
слов -> нечёткое сравнение.
"""
from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher

# Слова, которые не влияют на смысл названия предмета
STOPWORDS = {"и", "в", "во", "по", "на", "с", "со", "для", "из", "the", "of"}

_PUNCT_RE = re.compile(r"[^\w\s]+", re.UNICODE)
_SPACE_RE = re.compile(r"\s+")


def norm(value) -> str:
    """Лёгкая нормализация: регистр, ё, пробелы, «хвостовые» подчёркивания.

    Используется как основной ключ сопоставления.
    """
    if value is None:
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFC", text)
    text = text.replace("Ё", "Е").replace("ё", "е")
    # неразрывные пробелы и прочие пробельные символы
    text = re.sub(r"[\u00a0\u2007\u202f\u2009]", " ", text)
    text = _SPACE_RE.sub(" ", text)
    # «Химия_», «Физика -» и подобное: служебные хвосты в таблице
    text = text.rstrip("_-—–·.:; ")
    text = _SPACE_RE.sub(" ", text).strip()
    return text.lower()


def norm_strict(value) -> str:
    """Жёсткая нормализация: только буквы/цифры/пробелы (без знаков препинания)."""
    text = _PUNCT_RE.sub(" ", norm(value))
    return _SPACE_RE.sub(" ", text).strip()


def tokens(value) -> frozenset:
    """Набор значимых слов (без предлогов/союзов). Порядок слов не важен."""
    words = [w for w in norm_strict(value).split() if w and w not in STOPWORDS]
    return frozenset(words)


def token_key(value) -> str:
    """Сортированный набор слов — для сравнения переставленных названий."""
    return " ".join(sorted(tokens(value)))


def fuzzy_score(a: str, b: str) -> float:
    """Оценка похожести двух названий в диапазоне 0..1.

    Комбинация посимвольного сходства и перекрытия наборов слов.
    """
    na, nb = norm_strict(a), norm_strict(b)
    if not na or not nb:
        return 0.0
    ratio = SequenceMatcher(None, na, nb).ratio()
    ta, tb = tokens(a), tokens(b)
    if ta and tb:
        inter = len(ta & tb)
        union = len(ta | tb)
        jaccard = inter / union
        containment = inter / min(len(ta), len(tb))
    else:
        jaccard = containment = 0.0
    return round(0.35 * ratio + 0.25 * jaccard + 0.40 * containment, 4)


def best_match(text: str, candidates, threshold: float = 0.72, margin: float = 0.04):
    """Лучший кандидат из последовательности строк.

    Возвращает (candidate, score, ambiguous). ambiguous=True, если второй
    результат почти так же хорош — такое совпаждение лучше показать человеку.
    """
    scored = sorted(((fuzzy_score(text, c), c) for c in candidates), key=lambda x: (-x[0], str(x[1])))
    if not scored:
        return None, 0.0, False
    best_score, best = scored[0]
    if best_score < threshold:
        return None, best_score, False
    ambiguous = len(scored) > 1 and (best_score - scored[1][0]) < margin
    return best, best_score, ambiguous


def digits_only(value) -> str:
    return re.sub(r"\D", "", str(value or ""))
