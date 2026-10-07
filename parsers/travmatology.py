# parsers/travmatology.py
"""Парсер графика травматологии в формате заведующего.

Особенности формата:
- Дни месяца 1..31 — в отдельной строке (6-я Excel), обычно колонки 1..31.
- В колонке 0 — ФИО врача, в той же строке — времена ДЕЖУРСТВ.
- Строки БЕЗ ФИО (пустая колонка 0) — времена дневной работы, игнорируются.
- Время в формате "08.30-08.30" (точки, не двоеточия).
- Месяц и год — в заголовке (ячейка C1): "... на ОКТЯБРЬ 2026 года".
- Второй блок чисел 1..31 (с ~44-й строки) — дубль для расчёта часов, игнорируется.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from core.config import MONTHS_SEARCH
from core.doctor_validator import is_valid_doctor_name
from core.utils import read_excel_or_csv


_TIME_RE = re.compile(r'^\s*(\d{1,2})[.:](\d{2})\s*[-—]\s*(\d{1,2})[.:](\d{2})\s*$')


def _normalize_travma_time(value) -> str | None:
    """Приводит "08.30-08.30" к "08:30-08:30". None, если это не время."""
    if value is None or pd.isna(value):
        return None
    s = str(value).strip()
    m = _TIME_RE.match(s)
    if not m:
        return None
    h1, m1, h2, m2 = m.groups()
    return f"{int(h1):02d}:{m1}-{int(h2):02d}:{m2}"


def _find_days_row(df: pd.DataFrame):
    """Ищет строку с числами 1..31 (первый непрерывный диапазон).

    Возвращает (row_idx, {col_idx: day_num}) или (None, {}).
    Второй блок чисел (дубль для часов) игнорируется: останавливаемся
    на первом встреченном 31.
    """
    for i in range(min(15, len(df))):
        row = df.iloc[i].values
        day_by_col: dict[int, int] = {}
        expected = 1
        for col_idx, v in enumerate(row):
            if isinstance(v, (int, float)) and not pd.isna(v):
                iv = int(v)
                if iv == expected and iv <= 31:
                    day_by_col[col_idx] = iv
                    expected += 1
                    if iv == 31:
                        break
        if len(day_by_col) >= 28:
            return i, day_by_col
    return None, {}


def _extract_month_year(df: pd.DataFrame):
    """Берём месяц и год из ячейки C1 (row 0, col 2) — там заголовок
    типа "... на ОКТЯБРЬ 2026 года"."""
    if len(df) == 0 or df.shape[1] <= 2:
        return None, None

    header_cell = df.iloc[0, 2]
    if pd.isna(header_cell):
        return None, None

    text = str(header_cell)
    text_lower = text.lower()

    month = None
    for name, num in MONTHS_SEARCH.items():
        if name in text_lower:
            month = num
            break

    year = None
    m = re.search(r'\b(20\d{2})\b', text)
    if m:
        year = int(m.group(1))

    return month, year


def parse_travmatology_graph(file_path, dept_name: str = "Травматология (зав.)"):
    """Парсит график травматологии в формате заведующего.

    Возвращает: (data, dept_name, months_set, years_set).
    """
    print(f"\n   🔍 Парсим {Path(file_path).name} (травматология, формат заведующего)...")
    df = read_excel_or_csv(file_path)
    if df is None:
        return {}, dept_name, None, None

    days_row_idx, day_by_col = _find_days_row(df)
    if days_row_idx is None:
        print("   ⚠️ Не найдена строка с числами месяца 1..31")
        return {}, dept_name, None, None
    print(f"   📅 Строка с числами месяца: {days_row_idx}, "
          f"дней: {len(day_by_col)} (колонки {min(day_by_col)}..{max(day_by_col)})")

    month, year = _extract_month_year(df)
    months_found = {month} if month else set()
    years_found = {year} if year else set()
    print(f"   🗓️  Месяц: {month}, Год: {year}")

    result: dict[int, dict] = {}
    doctors_found = 0
    doctors_with_shifts = 0

    for i in range(days_row_idx + 1, len(df)):
        first_cell = df.iloc[i, 0]
        if pd.isna(first_cell) or not str(first_cell).strip():
            continue

        doctor = str(first_cell).strip()
        if not is_valid_doctor_name(doctor):
            continue

        doctors_found += 1
        row = df.iloc[i].values
        saved = 0
        for col_idx, day_num in day_by_col.items():
            if col_idx >= len(row):
                continue
            normalized = _normalize_travma_time(row[col_idx])
            if normalized:
                if day_num not in result:
                    result[day_num] = {}
                result[day_num][(doctor, dept_name)] = normalized
                saved += 1

        if saved:
            doctors_with_shifts += 1
            print(f"   👨‍⚕️  {doctor}: {saved} дежурств")

    print(f"   📊 Всего врачей: {doctors_found}, "
          f"с дежурствами: {doctors_with_shifts}, "
          f"дней с дежурствами: {len(result)}")
    return result, dept_name, (months_found or None), (years_found or None)