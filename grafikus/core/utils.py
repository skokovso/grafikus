# core/utils.py
"""Вспомогательные функции для парсинга и обработки данных."""
import re
from pathlib import Path
import pandas as pd
from .config import MONTHS_SEARCH

def normalize_time(time_str: str) -> str:
    """Нормализует строку времени в формат ЧЧ-ЧЧ или ЧЧ:ММ-ЧЧ:ММ."""
    if not time_str or time_str == '':
        return time_str
    time_str = time_str.strip()
    
    # Обработка формата "9\8" или "10\8"
    match = re.match(r'(\d{1,2})\s*[\\/]\s*(\d{1,2})', time_str)
    if match:
        h1, h2 = match.groups()
        return f"{int(h1):02d}-{int(h2):02d}"
        
    match = re.match(r'(\d{1,2}):(\d{2})\s*[-—]\s*(\d{1,2}):(\d{2})', time_str)
    if match:
        h1, m1, h2, m2 = match.groups()
        return f"{int(h1):02d}:{m1}-{int(h2):02d}:{m2}"
        
    match = re.match(r'(\d{1,2})\s*[/]\s*(\d{1,2})', time_str)
    if match:
        h1, h2 = match.groups()
        return f"{int(h1):02d}-{int(h2):02d}"
        
    match = re.match(r'(\d{1,2})\s*[-—]\s*(\d{1,2})', time_str)
    if match:
        h1, h2 = match.groups()
        return f"{int(h1):02d}-{int(h2):02d}"
        
    match = re.match(r'0-(\d{1,2})', time_str)
    if match:
        h2 = match.group(1)
        return f"00-{int(h2):02d}"
        
    return time_str

def extract_day_from_date(val) -> int | None:
    """Извлекает день из строки даты."""
    val_str = str(val).strip()
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', val_str)
    if match: return int(match.group(3))
    match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', val_str)
    if match: return int(match.group(1))
    match = re.search(r'(\d{2})/(\d{2})/(\d{4})', val_str)
    if match: return int(match.group(1))
    match = re.search(r'\b([1-9]|[12][0-9]|3[01])\b', val_str)
    if match: return int(match.group(1))
    return None

def extract_month_from_date(val) -> int | None:
    """Извлекает месяц из строки даты."""
    val_str = str(val).strip()
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', val_str)
    if match: return int(match.group(2))
    match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', val_str)
    if match: return int(match.group(2))
    match = re.search(r'(\d{2})/(\d{2})/(\d{4})', val_str)
    if match: return int(match.group(2))
    return None

def extract_year_from_date(val) -> int | None:
    """Извлекает год из строки даты."""
    val_str = str(val).strip()
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', val_str)
    if match: return int(match.group(1))
    match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', val_str)
    if match: return int(match.group(3))
    match = re.search(r'(\d{2})/(\d{2})/(\d{4})', val_str)
    if match: return int(match.group(3))
    return None

def find_month_in_text(text: str) -> int | None:
    """Ищет месяц в тексте (регистронезависимо)."""
    text_lower = text.lower()
    for month_name, month_num in MONTHS_SEARCH.items():
        if month_name in text_lower:
            return month_num
    return None

def read_excel_or_csv(file_path: str | Path) -> pd.DataFrame | None:
    """Читает Excel или CSV файл."""
    file_path = Path(file_path)
    if not file_path.exists():
        return None
    if file_path.suffix.lower() in ['.xlsx', '.xls']:
        try:
            df = pd.read_excel(file_path, header=None)
            return df
        except Exception as e:
            print(f"   ⚠️ Ошибка чтения Excel: {e}")
            return None
    elif file_path.suffix.lower() == '.csv':
        encodings = ['utf-8-sig', 'cp1251', 'latin-1']
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, sep=';', encoding=enc, header=None)
                return df
            except:
                continue
        return None
    return None