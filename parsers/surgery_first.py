# parsers/surgery.py
"""Парсеры для хирургии (PDF и TXT). ИСПРАВЛЕН БАГ С ПЕРЕНОСОМ ВТОРОГО ВРАЧА."""
import re
from pathlib import Path
from core.utils import normalize_time
from core.doctor_validator import is_valid_doctor_name

PDF_AVAILABLE = False
PDF_LIB = None
try:
    import pdfplumber
    PDF_AVAILABLE = True
    PDF_LIB = 'pdfplumber'
except ImportError:
    try:
        from pypdf import PdfReader
        PDF_AVAILABLE = True
        PDF_LIB = 'pypdf'
    except ImportError:
        pass

def _parse_surgery_text(text: str, dept_name: str):
    """Парсит текст из PDF/TXT (исправленная версия: все врачи дня остаются в своём дне)."""
    result = {}
    months_found = set()
    years_found = set()
    lines = text.split('\n')
    date_pattern = re.compile(r'(\d{2})/(\d{2})/(\d{4})')
    time_pattern = re.compile(r'(\d{2}):(\d{2})\s*[-—]\s*(\d{2}):(\d{2})')

    parsed_entries = []
    current_day = None

    for line in lines:
        line = line.strip()
        if not line or 'Утверждаю' in line or 'Хирургическое' in line or 'График' in line:
            continue

        date_match = date_pattern.search(line)
        if date_match:
            current_day = int(date_match.group(1))
            months_found.add(int(date_match.group(2)))
            years_found.add(int(date_match.group(3)))
            line = date_pattern.sub('', line).strip()

        if current_day is None or not line:
            continue

        time_match = time_pattern.search(line)
        if time_match:
            time_str = f"{time_match.group(1)}/{time_match.group(3)}"
            normalized = normalize_time(time_str)
            doctor_part = line[:time_match.start()].strip()
            doctor = ' '.join(doctor_part.split())
            if doctor and is_valid_doctor_name(doctor):
                parsed_entries.append((current_day, doctor, normalized))
        else:
            alt_time_match = re.search(r'(\d{2}:\d{2})\s*[-—]\s*(\d{2}:\d{2})', line)
            if alt_time_match:
                time_str = f"{alt_time_match.group(1)[:2]}/{alt_time_match.group(2)[:2]}"
                normalized = normalize_time(time_str)
                doctor_part = line[:alt_time_match.start()].strip()
                doctor = ' '.join(doctor_part.split())
                if doctor and is_valid_doctor_name(doctor):
                    parsed_entries.append((current_day, doctor, normalized))

    # ИСПРАВЛЕНИЕ: все врачи одного дня остаются в своём дне
    day_doctors = {}
    for day, doctor, time_str in parsed_entries:
        if day not in day_doctors:
            day_doctors[day] = []
        day_doctors[day].append((doctor, time_str))

    for day, doctors in day_doctors.items():
        if day not in result:
            result[day] = {}
        for doctor, time_str in doctors:
            result[day][(doctor, dept_name)] = time_str

    return result, months_found, years_found

def parse_surgery_pdf(file_path: str | Path, dept_name: str):
    print(f"\n   🔍 Парсим {Path(file_path).name} (PDF)...")
    if not PDF_AVAILABLE:
        print(f"   ⚠️ PDF-библиотека не установлена")
        return {}, "Хирургия", None, None

    result = {}
    months_found = set()
    years_found = set()
    dept_name = "Хирургия"

    try:
        if PDF_LIB == 'pdfplumber':
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        res, months, years = _parse_surgery_text(text, dept_name)
                        result.update(res)
                        months_found.update(months)
                        years_found.update(years)
        elif PDF_LIB == 'pypdf':
            reader = PdfReader(file_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    res, months, years = _parse_surgery_text(text, dept_name)
                    result.update(res)
                    months_found.update(months)
                    years_found.update(years)
    except Exception as e:
        print(f"   ⚠️ Ошибка чтения PDF: {e}")
        return {}, dept_name, None, None

    return result, dept_name, months_found, years_found

def parse_surgery_txt(file_path: str | Path, dept_name: str):
    print(f"\n   🔍 Парсим {Path(file_path).name} (TXT)...")
    encodings = ['utf-8-sig', 'cp1251', 'cp866', 'koi8-r', 'latin-1']
    lines = None
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                lines = f.readlines()
                break
        except UnicodeDecodeError:
            continue

    if lines is None:
        with open(file_path, 'rb') as f:
            raw = f.read()
            lines = raw.decode('utf-8', errors='ignore').splitlines()

    text = '\n'.join(lines)
    result, months, years = _parse_surgery_text(text, dept_name)
    return result, "Хирургия", months, years