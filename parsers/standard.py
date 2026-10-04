# parsers/standard.py
"""Парсеры для стандартных файлов отделений и ответственных."""
import re
import pandas as pd
from pathlib import Path
from core.utils import read_excel_or_csv, extract_day_from_date, extract_month_from_date, extract_year_from_date, normalize_time, find_month_in_text
from core.doctor_validator import is_valid_doctor_name

def parse_standard_graph(file_path: str | Path, dept_name: str):
    print(f"\n   🔍 Парсим {Path(file_path).name}...")
    df = read_excel_or_csv(file_path)
    if df is None:
        return {}, dept_name, None, None

    date_row_idx = None
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
        if 'Дата' in row_str and re.search(r'\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4}', row_str):
            date_row_idx = i
            break
            
    if date_row_idx is None:
        for i, row in df.iterrows():
            row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
            if len(re.findall(r'\d+', row_str)) >= 10:
                date_row_idx = i
                break

    if date_row_idx is None:
        return {}, dept_name, None, None

    date_row = df.iloc[date_row_idx].values
    day_cols = []
    months_found = set()
    years_found = set()
    
    for i, val in enumerate(date_row):
        if pd.notna(val):
            day_num = extract_day_from_date(val)
            month_num = extract_month_from_date(val)
            year_num = extract_year_from_date(val)
            if month_num: months_found.add(month_num)
            if year_num: years_found.add(year_num)
            if day_num and 1 <= day_num <= 31:
                day_cols.append((i, day_num))

    # Поиск месяца в тексте заголовка (регистронезависимо)
    header_text = ' '.join(str(v) for row in df.iloc[:5].values for v in row if pd.notna(v))
    found_month = find_month_in_text(header_text)
    if found_month:
        months_found.add(found_month)

    header_row_idx = None
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
        if 'Дежурный врач' in row_str or 'дежурный' in row_str.lower():
            header_row_idx = i
            break

    if header_row_idx is None:
        return {}, dept_name, months_found if months_found else None, years_found if years_found else None

    result = {day: {} for _, day in day_cols}
    doctors_found = 0
    
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i].values
        if len(row) == 0 or pd.isna(row[0]) or str(row[0]).strip() == '':
            continue
        doctor = str(row[0]).strip()
        if not is_valid_doctor_name(doctor):
            continue
        if 'Заведующий' in doctor:
            break
        doctors_found += 1
        
        for col_idx, day_num in day_cols:
            if col_idx < len(row) and pd.notna(row[col_idx]):
                val = str(row[col_idx]).strip()
                if val and val not in ['', 'nan', 'None']:
                    if re.search(r'\d+[-/]\d+|\d+:\d+|[АРЭ]', val):
                        normalized = normalize_time(val)
                        result[day_num][(doctor, dept_name)] = normalized
                        
    print(f"   👨‍⚕️ Найдено {doctors_found} врачей")
    return result, dept_name, months_found if months_found else None, years_found if years_found else None

def parse_responsible_graph(file_path: str | Path, dept_name: str):
    # Ответственные парсятся почти так же, но с чуть более гибким поиском строки дат
    return parse_standard_graph(file_path, dept_name)