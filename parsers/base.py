# parsers/base.py
"""Базовые функции парсинга и определения отделений."""
import re
from pathlib import Path
import pandas as pd
from core.utils import read_excel_or_csv, find_month_in_text
from core.config import DEPT_ABBR, DEPARTMENT_ORDER

def get_dept_abbr(dept_name: str) -> str:
    """Возвращает сокращение названия отделения."""
    if dept_name in DEPT_ABBR:
        return DEPT_ABBR[dept_name]
    for key, abbr in DEPT_ABBR.items():
        if key.lower() in dept_name.lower() or dept_name.lower() in key.lower():
            return abbr
    words = dept_name.split()
    if len(words) >= 2:
        return ''.join(w[0].upper() for w in words[:2])
    return dept_name[:4].upper()

def get_dept_order(dept_name: str) -> int:
    """Возвращает порядок сортировки отделения."""
    for i, dept in enumerate(DEPARTMENT_ORDER):
        if dept.lower() in dept_name.lower() or dept_name.lower() in dept.lower():
            return i
    return len(DEPARTMENT_ORDER)

def detect_department(df: pd.DataFrame, file_path: str | Path) -> str:
    """Определяет отделение по содержимому файла."""
    file_stem = Path(file_path).stem.lower()
    
    # Для многостраничного Excel (Володарского)
    if 'график дежурств 2026' in file_stem or 'график дежурств' in file_stem:
        for i in range(min(20, len(df))):
            row = df.iloc[i].values
            row_str = ' '.join(str(v) for v in row if pd.notna(v))
            for keyword, dept in {
                'ОАР': 'ОАР', 'Хирургия': 'Хирургия', 'Приемное': 'Приемное отделение',
                'Гинекология': 'Гинекология', 'Травматология': 'Травматология',
                'Неврология': 'Неврология ОНМК', 'Терапевты': 'Терапевты',
            }.items():
                if keyword in row_str:
                    return dept
        return 'Володарского (сводный)'

    header_text = ''
    for i in range(min(10, len(df))):
        row = df.iloc[i].values
        row_str = ' '.join(str(v) for v in row if pd.notna(v))
        header_text += row_str + ' '
    header_text = header_text.lower()

    for dept in DEPARTMENT_ORDER:
        dept_clean = re.sub(r'[_\s]+', ' ', dept.lower())
        clean_filename = re.sub(r'графики?_дежурных?_врачей?_', '', file_stem)
        clean_filename = re.sub(r'[_\s]+', ' ', clean_filename).strip()
        if dept_clean in clean_filename or clean_filename in dept_clean:
            return dept

    dept_patterns = {
        'Ответственные по стационару': ['ответственн', 'ответсвтвенн', 'стационар'],
        'Приемное отделение': ['приемн', 'приёмн', 'приемного'],
        'ОАР': ['оар', 'реаниматолог', 'анестезиолог', 'анестезия'],
        'Хирургия': ['хирург', 'хирургическ'],
        'Гинекология': ['гинеколог', 'гинекологическ'],
        'Урология': ['уролог', 'урологическ'],
        'Травматология': ['травматолог', 'травматологическ'],
        'Травмпункт': ['травмпункт', 'травм пункт'],
    }
    for dept, patterns in dept_patterns.items():
        for pattern in patterns:
            if pattern in header_text:
                return dept
    return Path(file_path).stem.replace('_', ' ')