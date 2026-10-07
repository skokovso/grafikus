# parsers/administration.py
"""Парсер для графика дежурств администрации с фильтрацией по площадке."""
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from core.utils import read_excel_or_csv

def parse_administration_graph(file_path: str | Path, selected_location: str):
    print(f"\n   🔍 Парсим {Path(file_path).name} (Администрация)...")
    print(f"   📍 Выбранная площадка: '{selected_location}'")
    
    df = read_excel_or_csv(file_path)
    if df is None:
        return {}, "АДМ", None, None

    # Нормализуем название выбранной площадки
    selected_location_clean = selected_location.strip().lower()
    print(f"   🔎 Ищем площадку: '{selected_location_clean}'")
    
    target_locations = ["ломоносова, д.18", "володарского, д.18", "перинатальный центр", "родильный дом"]
    
    # 1. Ищем строку-заголовок с названиями площадок
    header_idx = None
    location_cols = {}
    
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v)).lower()
        
        if any(loc in row_str for loc in target_locations):
            header_idx = i
            print(f"   📋 Найдена строка заголовка на индексе {i}")
            
            for col_idx, val in enumerate(row):
                if pd.notna(val):
                    val_str = str(val).strip().lower()
                    for target in target_locations:
                        if target in val_str:
                            location_cols[col_idx] = target
                            print(f"      📍 Колонка {col_idx}: '{val_str}' -> '{target}'")
                            break
            break

    if header_idx is None:
        print(f"   ⚠️ Не найдена строка с названиями площадок")
        return {}, "АДМ", None, None

    if not location_cols:
        print(f"   ⚠️ Не найдены колонки с площадками")
        return {}, "АДМ", None, None

    print(f"   📊 Найдено колонок с площадками: {len(location_cols)}")

    result = {}
    months_found = set()
    years_found = set()

    # 2. Парсим строки с датами
    print(f"\n   🔎 Начинаем парсинг данных со строки {header_idx + 1}...")
    
    for i in range(header_idx + 1, len(df)):
        row = df.iloc[i].values
        if len(row) == 0 or pd.isna(row[0]):
            continue

        date_val = row[0]
        
        # Пропускаем служебные строки
        if isinstance(date_val, str) and ('справочник' in date_val.lower() or 'месяц' in date_val.lower() or 'год' in date_val.lower()):
            print(f"   ⏭️ Пропускаем служебную строку: {date_val}")
            break

        # Извлекаем день из даты
        day_num = None
        month_num = None
        year_num = None

        # datetime / pd.Timestamp — основной случай для этого файла
        if isinstance(date_val, (datetime, pd.Timestamp)):
            day_num = date_val.day
            month_num = date_val.month
            year_num = date_val.year
        elif isinstance(date_val, str):
            s = date_val.strip()
            # Формулу =DATE(...) или =IF(...) сюда не пускаем — скипаем
            if s.startswith('='):
                continue
            # Строковый формат "10/1/26" или "01.10.2026" или "2026-10-01"
            parts = re.findall(r'\d+', s)
            if len(parts) >= 3:
                if len(parts[0]) == 4:      # 2026-10-01
                    year_num = int(parts[0])
                    month_num = int(parts[1])
                    day_num = int(parts[2])
                else:                        # 10/1/26 или 01.10.2026
                    a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
                    if c < 100:
                        c += 2000
                    # Эвристика: если a > 12 — это день, иначе считаем м/д/г
                    if a > 12:
                        day_num, month_num, year_num = a, b, c
                    else:
                        month_num, day_num, year_num = a, b, c

        if not day_num or not month_num:
            continue
            
        if not (1 <= month_num <= 12 and 1 <= day_num <= 31):
            continue

        months_found.add(month_num)
        years_found.add(year_num)

        # 3. Проверяем колонки площадок
        for col_idx, location_name in location_cols.items():
            if location_name == selected_location_clean:
                if col_idx < len(row) and pd.notna(row[col_idx]):
                    doctor_name = str(row[col_idx]).strip()
                    print(f"   🔍 Строка {i}, день {day_num:02d}.{month_num:02d}, колонка {col_idx} ({location_name}): '{doctor_name}'")
                    
                    if doctor_name and len(doctor_name) > 2 and re.search(r'[а-яА-ЯёЁ]', doctor_name):
                        if day_num not in result:
                            result[day_num] = {}
                        result[day_num][(doctor_name, "АДМ")] = ""
                        print(f"      ✅ Добавлено: {doctor_name} на {day_num:02d}.{month_num:02d}")

    total_dejurstva = sum(len(v) for v in result.values())
    print(f"\n   📊 Обработано строк: {len(result)}")
    print(f"   ✅ Найдено дежурств администрации: {total_dejurstva}")
    return result, "АДМ", months_found, years_found