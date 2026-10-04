# parsers/oar.py
"""Парсер для ОАР с извлечением времени из предпоследней строки."""
import re
import pandas as pd
from pathlib import Path
from core.utils import read_excel_or_csv, extract_day_from_date, extract_month_from_date, extract_year_from_date, find_month_in_text
from core.doctor_validator import is_valid_doctor_name

def parse_oar_graph(file_path: str | Path, dept_name: str):
    print(f"\n   🔍 Парсим {Path(file_path).name} (ОАР)...")
    df = read_excel_or_csv(file_path)
    if df is None:
        return {}, dept_name, None, None

    header_row_idx = None
    time_row_idx = None
    months_found = set()
    years_found = set()

    # Поиск месяца в первых строках
    for i in range(min(5, len(df))):
        row = df.iloc[i].values
        row_str = ' '.join(str(v) for v in row if pd.notna(v))
        found_month = find_month_in_text(row_str)
        if found_month:
            months_found.add(found_month)

    # Поиск строки "Ф.И.О." и строки с временем (предпоследняя)
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
        if 'Ф.И.О.' in row_str:
            header_row_idx = i
        # Ищем строку с временем (формат 16/09, 09/09 и т.д.), но не строку с днями (1, 2, 3...)
        if re.search(r'\d{2}/\d{2}', row_str) and 'Ф.И.О.' not in row_str:
            numbers = [str(v).strip() for v in row.values if pd.notna(v)]
            # Если в первых 5 ячейках есть только цифры 1-31, это строка дней, пропускаем
            if len(numbers) > 10 and any(re.match(r'^\d{1,2}$', n) for n in numbers[:5]):
                continue
            time_row_idx = i

    if header_row_idx is None:
        return {}, dept_name, months_found if months_found else None, years_found if years_found else None

    header_row = df.iloc[header_row_idx].values
    day_cols = []
    for i, val in enumerate(header_row):
        if pd.notna(val):
            day_num = extract_day_from_date(val)
            month_num = extract_month_from_date(val)
            year_num = extract_year_from_date(val)
            if month_num: months_found.add(month_num)
            if year_num: years_found.add(year_num)
            if day_num and 1 <= day_num <= 31:
                day_cols.append((i, day_num))

    # Парсим время из предпоследней строки
    day_times = {}
    if time_row_idx is not None:
        time_row = df.iloc[time_row_idx].values
        for col_idx, day_num in day_cols:
            if col_idx < len(time_row) and pd.notna(time_row[col_idx]):
                time_str = str(time_row[col_idx]).strip()
                match = re.match(r'(\d{1,2})/(\d{1,2})', time_str)
                if match:
                    h1, h2 = match.groups()
                    day_times[day_num] = f"{int(h1):02d}-{int(h2):02d}"

    result = {day: {} for _, day in day_cols}
    doctors_found = 0

    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i].values
        if len(row) == 0 or pd.isna(row[0]) or str(row[0]).strip() == '':
            continue
        doctor = str(row[0]).strip()
        if not is_valid_doctor_name(doctor):
            continue
        doctors_found += 1

        for col_idx, day_num in day_cols:
            if col_idx < len(row) and pd.notna(row[col_idx]):
                val = str(row[col_idx]).strip()
                if val:
                    match = re.search(r'([АРЭ])', val)
                    if match:
                        letter = match.group(1)
                        time_str = day_times.get(day_num, '')
                        # ИСПРАВЛЕНИЕ: буква ПЕРЕД временем (А16-09 вместо 16-09А)
                        result[day_num][(doctor, dept_name)] = f"{letter}{time_str}" if time_str else letter

    print(f"   👨‍⚕️ Найдено {doctors_found} врачей")
    return result, dept_name, months_found if months_found else None, years_found if years_found else None