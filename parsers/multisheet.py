# parsers/multisheet.py
"""Парсер для многостраничных Excel-файлов (Володарского).

Поддерживает:
- Чтение листа по названию месяца.
- Отделения внутри листа (Администрация, ОАР, Хирургия, ...).
- Известных врачей (KNOWN_DOCTORS) — на случай редких фамилий.
- Заливку ячеек: если ячейка с временем дежурства залита цветом
  (не белым), этот же врач дополнительно попадает в секцию
  «Ответственные по стационару» с тем же временем.
"""
import re
import pandas as pd
from pathlib import Path

from openpyxl import load_workbook

from core.utils import normalize_time, find_month_in_text
from core.doctor_validator import is_valid_doctor_name
from core.config import DEPT_ABBR


RESPONSIBLE_DEPT = "Ответственные по стационару"


def _load_colored_cells(file_path, sheet_name: str) -> set[tuple[int, int]]:
    """Возвращает множество 0-based (row, col) ячеек, залитых цветом.

    Игнорирует белый и «пустой» цвет. Результат сопоставляется с
    координатами pandas (тоже 0-based).
    """
    colored: set[tuple[int, int]] = set()
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            return colored
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                fill = cell.fill
                if fill is None or fill.fill_type is None:
                    continue
                start_color = fill.start_color
                rgb = getattr(start_color, 'rgb', None)
                if not isinstance(rgb, str):
                    continue
                rgb_up = rgb.upper()
                if rgb_up in ('00000000', 'FFFFFFFF'):
                    continue
                # openpyxl 1-based → pandas 0-based
                colored.add((cell.row - 1, cell.column - 1))
        wb.close()
    except Exception as e:
        print(f"   ⚠️ Не удалось прочитать цвета: {e}")
    return colored


def parse_excel_with_months(file_path, selected_month: int, selected_year: int):
    """Парсит многостраничный Excel (Володарского).

    Возвращает: (data, months_set, years_set).
    """
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

    # --- Карта закрашенных ячеек (0-based row, col) ---
    # Важно: read_only=True иногда не отдаёт fill корректно.
    # Если увидите 0 цветных — замените на read_only=False.
    colored_cells = _load_colored_cells(file_path, month_name_ru)
    if colored_cells:
        print(f"   🎨 Найдено цветных ячеек: {len(colored_cells)}")

    dept_keywords = {
        'Лаборатория ОАР': 'Лаборатория ОАР',
        'Лаборотория ОАР': 'Лаборатория ОАР',
        'ОАР': 'ОАР',
        'Хирургия': 'Хирургия',
        'Гнойная хирургия': 'Гнойная хирургия',
        'Приемное отделение': 'Приемное отделение',
        'МХГ': 'МХГ',
        'Гинекология': 'Гинекология',
        'Травматология': 'Травматология',
        'Рентген': 'Рентген',
        'Неврология ОНМК': 'Неврология ОНМК',
        'Терапевты': 'Терапевты',
    }

    # Известные врачи (могут не пройти is_valid_doctor_name)
    KNOWN_DOCTORS = ['Кодиров', 'Кодиров ИМ', 'Кодиров И.М.']

    # --- Определяем реальный год из заголовка листа ---
    # Ищем в первых 5 строках что-то вроде "График дежурств на октябрь 2026 года"
    import re as _re
    detected_year = None
    for _r in range(min(5, len(df))):
        _row = df.iloc[_r].values
        _text = ' '.join(str(v) for v in _row if pd.notna(v))
        _m = _re.search(r'\b(20\d{2})\b', _text)
        if _m:
            detected_year = int(_m.group(1))
            break

    if detected_year:
        years_found = {detected_year}
        print(f"   🗓️  Год из заголовка: {detected_year}")
    else:
        years_found = {selected_year}
        print(f"   🗓️  Год не найден в заголовке, используем выбранный: {selected_year}")

    result: dict[int, dict] = {}
    months_found = {selected_month}
    doctors_found = 0
    responsible_marked = 0  # сколько раз отметили ответственного по цвету

    i = 0
    while i < len(df):
        row = df.iloc[i].values
        row_str = ' '.join(str(v) for v in row if pd.notna(v))

        dept_name = None
        for keyword, dept in dept_keywords.items():
            if keyword in row_str:
                dept_name = dept
                break

        if not dept_name:
            i += 1
            continue

        # Явный пропуск блока «Администрация» — в этом файле не нужен
        if 'администрация' in row_str.lower():
            print(f"   ⏭️ Пропускаем блок: Администрация")
            i += 1
            while i < len(df):
                row_next = df.iloc[i].values
                row_str_next = ' '.join(str(v) for v in row_next if pd.notna(v))
                if any(keyword in row_str_next for keyword in dept_keywords.keys()):
                    break
                i += 1
            continue

        print(f"   🏥 Найдено отделение: {dept_name}")
        i += 1

        # Пропускаем строку с днями недели, если она сразу после названия
        if i < len(df):
            row_check = df.iloc[i].values
            row_str_check = ' '.join(str(v) for v in row_check if pd.notna(v))
            if any(day in row_str_check for day in ['Сб', 'Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт']):
                i += 1

        # Читаем врачей этого отделения
        while i < len(df):
            row = df.iloc[i].values
            if len(row) == 0 or pd.isna(row[0]):
                i += 1
                continue

            first_cell = str(row[0]).strip()
            row_str_full = ' '.join(str(v) for v in row if pd.notna(v))

            # Если встретили новое отделение — прерываем внутренний цикл
            if any(keyword in row_str_full for keyword in dept_keywords.keys()):
                break

            if not first_cell:
                i += 1
                continue
            if 'ФИО' in first_cell:
                i += 1
                continue
            if 'отпуск' in first_cell.lower() or 'отп' in first_cell.lower():
                i += 1
                continue
            if first_cell in ['х', 'Х', 'о', 'О']:
                i += 1
                continue

            # Проверяем, что это имя врача (или известный врач)
            is_doctor = is_valid_doctor_name(first_cell)
            if not is_doctor:
                for kd in KNOWN_DOCTORS:
                    if kd in first_cell:
                        is_doctor = True
                        break

            if is_doctor:
                doctor = first_cell
                doctors_found += 1

                for col_idx in range(1, len(row)):
                    if pd.isna(row[col_idx]):
                        continue
                    val = str(row[col_idx]).strip()
                    if not val or val in ['nan', 'None', 'х', 'Х', 'о', 'О']:
                        continue

                    # --- Нормализация значения (буква А/Р/Э в конце) ---
                    letter_match = re.search(r'([АРЭ])$', val)
                    if letter_match:
                        letter = letter_match.group(1)
                        time_part = val[:letter_match.start()].strip().rstrip('/')
                        normalized = normalize_time(time_part)
                        if normalized and normalized != '00-00':
                            value_to_write = f"{normalized}{letter}"
                        else:
                            value_to_write = None
                    else:
                        normalized = normalize_time(val)
                        if normalized and normalized != '00-00':
                            value_to_write = normalized
                        else:
                            value_to_write = None

                    if not value_to_write:
                        continue

                    # --- Основное дежурство ---
                    if col_idx not in result:
                        result[col_idx] = {}
                    result[col_idx][(doctor, dept_name)] = value_to_write

                    # --- Проверка заливки: этот же врач — ответственный ---
                    if (i, col_idx) in colored_cells:
                        result[col_idx][(doctor, RESPONSIBLE_DEPT)] = value_to_write
                        responsible_marked += 1
                        print(f"      🎨 {doctor} ({dept_name}) — ответственный, "
                              f"колонка {col_idx}, время {value_to_write}")

            i += 1

    print(f"   👨‍⚕️ Найдено врачей: {doctors_found}")
    if responsible_marked:
        print(f"   🎨 Отмечено ответственных по цвету: {responsible_marked}")

    return result, months_found, years_found