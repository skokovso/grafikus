# parsers/multisheet.py
"""Парсер для многостраничных Excel-файлов (Володарского)."""
import re
import pandas as pd
from pathlib import Path
from core.utils import normalize_time, find_month_in_text
from core.doctor_validator import is_valid_doctor_name
from core.config import DEPT_ABBR

def parse_excel_with_months(file_path: str | Path, selected_month: int, selected_year: int):
    print(f"\n   🔍 Парсим {Path(file_path).name} (многостраничный Excel)...")
    month_names_reverse = {
        1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель',
        5: 'май', 6: 'июнь', 7: 'июль', 8: 'август',
        9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'
    }
    month_name_ru = month_names_reverse.get(selected_month, '')
    if not month_name_ru:
        return {}, None, None

    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = [s.lower() for s in xl.sheet_names]
        if month_name_ru not in sheet_names:
            print(f"   ⚠️ Лист '{month_name_ru}' не найден в файле")
            return {}, None, None
        df = pd.read_excel(file_path, sheet_name=month_name_ru, header=None)
    except Exception as e:
        print(f"   ⚠️ Ошибка чтения Excel: {e}")
        return {}, None, None

    dept_keywords = {
        'Лаборатория ОАР': 'Лаборатория ОАР',
        'Лаборотория ОАР': 'Лаборатория ОАР',
        'ОАР': 'ОАР', 'Хирургия': 'Хирургия', 'Гнойная хирургия': 'Гнойная хирургия',
        'Приемное отделение': 'Приемное отделение', 'МХГ': 'МХГ',
        'Гинекология': 'Гинекология', 'Травматология': 'Травматология',
        'Рентген': 'Рентген', 'Неврология ОНМК': 'Неврология ОНМК', 'Терапевты': 'Терапевты',
    }

    result = {}
    months_found = {selected_month}
    years_found = {selected_year}
    doctors_found = 0

    i = 0
    while i < len(df):
        row = df.iloc[i].values
        row_str = ' '.join(str(v) for v in row if pd.notna(v))
        dept_name = None
        for keyword, dept in dept_keywords.items():
            if keyword in row_str:
                dept_name = dept
                break

        if dept_name:
            i += 1
            # Пропускаем строку с днями недели, если она есть сразу после названия отделения
            if i < len(df):
                row_check = df.iloc[i].values
                row_str_check = ' '.join(str(v) for v in row_check if pd.notna(v))
                if any(day in row_str_check for day in ['Сб', 'Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт']):
                    i += 1

            while i < len(df):
                row = df.iloc[i].values
                if len(row) == 0 or pd.isna(row[0]):
                    i += 1
                    continue
                
                first_cell = str(row[0]).strip()
                row_str_full = ' '.join(str(v) for v in row if pd.notna(v))
                
                # Если встретили новое отделение, прерываем цикл
                if any(keyword in row_str_full for keyword in dept_keywords.keys()):
                    break
                
                if not first_cell or 'ФИО' in first_cell or 'отпуск' in first_cell.lower() or 'отп' in first_cell.lower() or first_cell in ['х', 'Х', 'о', 'О']:
                    i += 1
                    continue

                if is_valid_doctor_name(first_cell):
                    doctor = first_cell
                    doctors_found += 1
                    for col_idx in range(1, len(row)):
                        if pd.notna(row[col_idx]):
                            val = str(row[col_idx]).strip()
                            if val and val not in ['', 'nan', 'None', 'х', 'Х', 'о', 'О']:
                                letter_match = re.search(r'([АРЭ])$', val)
                                if letter_match:
                                    letter = letter_match.group(1)
                                    time_part = val[:letter_match.start()].strip().rstrip('/')
                                    normalized = normalize_time(time_part)
                                    if normalized and normalized != '00-00':
                                        if col_idx not in result:
                                            result[col_idx] = {}
                                        result[col_idx][(doctor, dept_name)] = f"{normalized}{letter}"
                                else:
                                    normalized = normalize_time(val)
                                    if normalized and normalized != '00-00':
                                        if col_idx not in result:
                                            result[col_idx] = {}
                                        result[col_idx][(doctor, dept_name)] = normalized
                i += 1
            continue
        i += 1

    print(f"   👨‍⚕️ Найдено {doctors_found} врачей")
    return result, months_found, years_found