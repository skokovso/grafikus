# parsers/surgery.py
"""Парсеры для хирургии (PDF и TXT). PDF парсится через извлечение таблиц."""
import os
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
    """Парсит текст хирургии из TXT (резервный вариант)."""
    result = {}
    months_found = set()
    years_found = set()
    lines = text.split('\n')
    date_pattern = re.compile(r'(\d{2})/(\d{2})/(\d{4})')
    time_pattern = re.compile(r'(\d{1,2}):(\d{2})\s*[-—]\s*(\d{1,2}):(\d{2})')

    parsed_entries = []
    current_day = None
    last_doctor = None

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
                last_doctor = None
            elif last_doctor and not doctor:
                parsed_entries.append((current_day, last_doctor, normalized))
                last_doctor = None
        else:
            alt_time_match = re.search(r'(\d{2}:\d{2})\s*[-—]\s*(\d{2}:\d{2})', line)
            if alt_time_match:
                time_str = f"{alt_time_match.group(1)[:2]}/{alt_time_match.group(2)[:2]}"
                normalized = normalize_time(time_str)
                doctor_part = line[:alt_time_match.start()].strip()
                doctor = ' '.join(doctor_part.split())
                if doctor and is_valid_doctor_name(doctor):
                    parsed_entries.append((current_day, doctor, normalized))
                    last_doctor = None
            else:
                doctor = ' '.join(line.split())
                if doctor and is_valid_doctor_name(doctor):
                    last_doctor = doctor

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


def _parse_surgery_table(table_data: list, dept_name: str):
    """
    Парсит таблицу, извлечённую из PDF через extract_tables().
    Обрабатывает сдвоенные ячейки: запоминаем последнюю дату и используем её для строк без даты.
    """
    result = {}
    months_found = set()
    years_found = set()
    
    date_pattern = re.compile(r'(\d{2})/(\d{2})/(\d{4})')
    time_pattern = re.compile(r'(\d{1,2}):(\d{2})\s*[-—]\s*(\d{1,2}):(\d{2})')
    
    last_day = None  # Запоминаем последнюю встреченную дату
    
    for row in table_data:
        if not row or len(row) < 2:
            continue
        
        # Ищем ячейку с датой (первая колонка)
        date_cell = row[0] if len(row) > 0 else None
        if date_cell and isinstance(date_cell, str):
            match = date_pattern.search(date_cell)
            if match:
                last_day = int(match.group(1))
                months_found.add(int(match.group(2)))
                years_found.add(int(match.group(3)))
        
        # Если даты нет в этой строке, используем последнюю запомненную
        if last_day is None:
            continue
        
        # Ищем врача и время в строке
        # Формат: [дата или пусто, врач, время]
        doctor_cell = row[1] if len(row) > 1 else None
        time_cell = row[2] if len(row) > 2 else None
        
        if not doctor_cell or not isinstance(doctor_cell, str):
            continue
        
        doctor = doctor_cell.strip()
        if not doctor or not is_valid_doctor_name(doctor):
            continue
        
        # Ищем время
        if time_cell and isinstance(time_cell, str):
            time_match = time_pattern.search(time_cell)
            if time_match:
                time_str = f"{time_match.group(1)}/{time_match.group(3)}"
                normalized = normalize_time(time_str)
                
                if last_day not in result:
                    result[last_day] = {}
                result[last_day][(doctor, dept_name)] = normalized
                print(f"      ✅ День {last_day:02d}: {doctor} → {normalized}")
        else:
            # Время может быть в той же ячейке, что и врач
            time_match = time_pattern.search(doctor_cell)
            if time_match:
                time_str = f"{time_match.group(1)}/{time_match.group(3)}"
                normalized = normalize_time(time_str)
                doctor_part = doctor_cell[:time_match.start()].strip()
                doctor = ' '.join(doctor_part.split())
                
                if doctor and is_valid_doctor_name(doctor):
                    if last_day not in result:
                        result[last_day] = {}
                    result[last_day][(doctor, dept_name)] = normalized
                    print(f"      ✅ День {last_day:02d}: {doctor} → {normalized}")
    
    return result, months_found, years_found


def parse_surgery_pdf(file_path: str | Path, dept_name: str):
    """Извлекает таблицы из PDF, сохраняет в Excel для отладки и парсит."""
    file_path = Path(file_path)
    print(f"\n   🔍 Парсим {file_path.name} (PDF→таблицы)...")
    
    if not PDF_AVAILABLE:
        print(f"   ⚠️ PDF-библиотека не установлена")
        return {}, "Хирургия", None, None

    result = {}
    months_found = set()
    years_found = set()
    dept_name = "Хирургия"
    
    all_tables = []

    try:
        if PDF_LIB == 'pdfplumber':
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    if tables:
                        print(f"   📄 Страница {page_num + 1}: найдено {len(tables)} таблиц")
                        for table_idx, table in enumerate(tables):
                            print(f"      Таблица {table_idx + 1}: {len(table)} строк")
                            all_tables.append({
                                'page': page_num + 1,
                                'table_idx': table_idx + 1,
                                'data': table
                            })
                            res, months, years = _parse_surgery_table(table, dept_name)
                            for day, doctors in res.items():
                                if day not in result:
                                    result[day] = {}
                                result[day].update(doctors)
                            months_found.update(months)
                            years_found.update(years)
                    else:
                        print(f"   ⚠️ Страница {page_num + 1}: таблицы не найдены, пробуем текст")
                        text = page.extract_text()
                        if text:
                            res, months, years = _parse_surgery_text(text, dept_name)
                            for day, doctors in res.items():
                                if day not in result:
                                    result[day] = {}
                                result[day].update(doctors)
                            months_found.update(months)
                            years_found.update(years)
        elif PDF_LIB == 'pypdf':
            from pypdf import PdfReader
            reader = PdfReader(file_path)
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    res, months, years = _parse_surgery_text(text, dept_name)
                    for day, doctors in res.items():
                        if day not in result:
                            result[day] = {}
                        result[day].update(doctors)
                    months_found.update(months)
                    years_found.update(years)
    except Exception as e:
        print(f"   ⚠️ Ошибка чтения PDF: {e}")
        import traceback
        traceback.print_exc()
        return {}, dept_name, None, None

    # Сохраняем извлечённые таблицы в Excel для отладки
    if all_tables:
        try:
            import pandas as pd
            excel_path = file_path.with_name(file_path.stem + '_extracted_tables.xlsx')
            
            rows = []
            for table_info in all_tables:
                page = table_info['page']
                table_idx = table_info['table_idx']
                table_data = table_info['data']
                
                rows.append([f"=== Страница {page}, Таблица {table_idx} ==="] + [''] * 10)
                
                for row in table_data:
                    cleaned_row = [str(cell) if cell is not None else '' for cell in row]
                    rows.append(cleaned_row)
                
                rows.append([''] * 11)
            
            df = pd.DataFrame(rows)
            df.to_excel(excel_path, index=False, header=False)
            print(f"   💾 Таблицы сохранены в Excel: {excel_path.name}")
        except Exception as e:
            print(f"   ⚠️ Ошибка сохранения Excel: {e}")

    print(f"   📊 Найдено {len(result)} дней с дежурствами")
    return result, dept_name, months_found, years_found


def parse_surgery_txt(file_path: str | Path, dept_name: str):
    """Парсит TXT-файл хирургии."""
    file_path = Path(file_path)
    print(f"\n   🔍 Парсим {file_path.name} (TXT)...")
    
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
    print(f"   📊 Найдено {len(result)} дней с дежурствами")
    return result, "Хирургия", months, years