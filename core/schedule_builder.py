# core/schedule_builder.py
"""Основная логика сборки сводного графика из файлов."""
import calendar
import datetime
from pathlib import Path
import pandas as pd

from core.config import MONTHS_RU, MONTHS_RU_NOMINATIVE, MONTHS_RU_UPPER, DEPT_ABBR
from core.utils import read_excel_or_csv
from parsers.base import detect_department, get_dept_abbr, get_dept_order
from parsers.standard import parse_standard_graph, parse_responsible_graph
from parsers.oar import parse_oar_graph
from parsers.surgery import parse_surgery_pdf, parse_surgery_txt
from parsers.multisheet import parse_excel_with_months
from output.word_builder import save_to_word
from output.txt_builder import save_to_txt
from parsers.administration import parse_administration_graph

def build_master_schedule(input_folder, rukovoditel_text, ploshadka_text,
                          selected_month, selected_year):
    """Собирает сводный график из файлов в папке."""
    print(f"\n📊 Начинаем сборку сводного графика...")
    print(f" Папка: {input_folder}")
    print(f"📅 Месяц: {selected_month}, Год: {selected_year}")
    print("=" * 60)

    folder = Path(input_folder)
    output_folder = folder / "Сводный график"
    output_folder.mkdir(exist_ok=True)

    # Ищем все файлы
    all_files = []
    for ext in ['*.xlsx', '*.xls', '*.csv', '*.pdf', '*.txt']:
        all_files.extend(folder.glob(ext))
    all_files = [f for f in all_files if not f.name.startswith('~')]
    all_files = [f for f in all_files if '_extracted.txt' not in f.name]
    all_files = [f for f in all_files if 'сводный' not in f.name.lower()]

    if not all_files:
        return False, [], "❌ Не найдено ни одного файла с графиками!"

    print(f"\n📄 Найдено {len(all_files)} файлов:")
    for f in all_files:
        print(f"   - {f.name}")
    print("\n" + "=" * 60)

    all_data = {}
    doctors_by_dept = {}
    dept_order_map = {}
    files_with_errors = []

    for file_path in all_files:
        print(f"\n{'='*60}")
        print(f"📂 Обрабатываем: {file_path.name}")

        is_multisheet = False
        is_oar_file = 'ОАР' in file_path.name or 'оар' in file_path.name.lower()

        # Проверка на многостраничный Excel (Володарского)
        if file_path.suffix.lower() in ['.xlsx', '.xls']:
            try:
                xl = pd.ExcelFile(file_path)
                sheet_names = [s.lower() for s in xl.sheet_names]
                month_names = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                               'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
                if any(s in sheet_names for s in month_names) and not is_oar_file:
                    is_multisheet = True
            except:
                pass

        if is_multisheet:
            print(f"   📚 Определён как многостраничный Excel (Володарского)")
            data, months, years = parse_excel_with_months(file_path, selected_month, selected_year)
            if data:
                for day, doctors in data.items():
                    if day not in all_data:
                        all_data[day] = {}
                    all_data[day].update(doctors)
                    for doctor_key in doctors.keys():
                        dept_name = doctor_key[1]
                        if dept_name not in doctors_by_dept:
                            doctors_by_dept[dept_name] = []
                        if doctor_key not in doctors_by_dept[dept_name]:
                            doctors_by_dept[dept_name].append(doctor_key)
                        dept_order_map[dept_name] = get_dept_order(dept_name)
                print(f"   ✅ Обработано {len(data)} дней, {len(doctors_by_dept)} отделений")
                continue
            else:
                print(f"   ⚠️ Не удалось распарсить многостраничный Excel, пробуем стандартный парсер")

        # Стандартный парсер (Ломоносова)
        df_temp = read_excel_or_csv(file_path)
        if df_temp is not None:
            dept_name = detect_department(df_temp, file_path)
        else:
            dept_name = Path(file_path).stem.replace('_', ' ')

        print(f"   🏥 Отделение: {dept_name}")
        print(f"   📛 Сокращение: {get_dept_abbr(dept_name)}")

        # Выбор парсера
        if 'администрация' in file_path.name.lower():
            # Передаем выбранную площадку для фильтрации!
            data, detected_dept, months, years = parse_administration_graph(file_path, ploshadka_text)
        elif dept_name == 'ОАР' or 'ОАР' in file_path.name:
            data, detected_dept, months, years = parse_oar_graph(file_path, dept_name)
        elif dept_name == 'Ответственные по стационару' or 'ответственн' in dept_name.lower() or 'Ответственные' in file_path.name:
            data, detected_dept, months, years = parse_responsible_graph(file_path, dept_name)
        elif file_path.suffix.lower() == '.pdf':
            data, detected_dept, months, years = parse_surgery_pdf(file_path, dept_name)
        elif file_path.suffix.lower() == '.txt':
            data, detected_dept, months, years = parse_surgery_txt(file_path, dept_name)
            if not data:
                print(f"   ⚠️ Не удалось прочитать TXT, пропускаем")
                continue
        else:
            data, detected_dept, months, years = parse_standard_graph(file_path, dept_name)

        if detected_dept and detected_dept != dept_name:
            dept_name = detected_dept

        if not data:
            print(f"   ⚠️ Данные не найдены для {dept_name}")
            continue

        # Проверка месяца и года
        if months and selected_month not in months:
            print(f"   ⚠️ Месяц в файле ({months}) не соответствует выбранному ({selected_month})")
            files_with_errors.append((file_path.name, f"Месяц: {months}, ожидался: {selected_month}"))
            continue
        if years and selected_year not in years:
            print(f"   ⚠️ Год в файле ({years}) не соответствует выбранному ({selected_year})")
            files_with_errors.append((file_path.name, f"Год: {years}, ожидался: {selected_year}"))
            continue

        for day, doctors in data.items():
            if day not in all_data:
                all_data[day] = {}
            all_data[day].update(doctors)
            if dept_name not in doctors_by_dept:
                doctors_by_dept[dept_name] = []
            for doctor_key in doctors.keys():
                if doctor_key not in doctors_by_dept[dept_name]:
                    doctors_by_dept[dept_name].append(doctor_key)
        dept_order_map[dept_name] = get_dept_order(dept_name)
        print(f"   ✅ Обработано {len(data)} дней, {len(doctors_by_dept.get(dept_name, []))} врачей")

    # Проверка результатов
    if not all_data:
        month_name_ru = MONTHS_RU_NOMINATIVE.get(selected_month, '')
        if files_with_errors:
            msg = f"Данные за {month_name_ru} {selected_year} года не найдены.\n\nИсключённые файлы:\n"
            for fname, error in files_with_errors[:5]:
                msg += f"  • {fname}\n"
        else:
            msg = f"Данные за {month_name_ru} {selected_year} года не найдены.\nПроверьте выбранную папку и месяц."
        return False, files_with_errors, msg

    # Сортировка отделений
    sorted_depts = sorted(doctors_by_dept.keys(), key=lambda d: (dept_order_map.get(d, 999), d))

    print("\n" + "=" * 60)
    print(f"✅ Всего найдено {len(all_data)} дней с дежурствами")
    print(f"📊 Найдено отделений: {len(doctors_by_dept)}")
    for dept in sorted_depts:
        print(f"   {dept} ({get_dept_abbr(dept)}): {len(doctors_by_dept[dept])} врачей")

    # Сохранение TXT
    month_lower = MONTHS_RU.get(selected_month, 'августа')
    txt_file = output_folder / f"сводный_график_{month_lower}_{selected_year}.txt"
    save_to_txt(all_data, doctors_by_dept, sorted_depts, txt_file, selected_month, selected_year)

    # Сохранение Word
    docx_file = output_folder / f"сводный_график_{month_lower}_{selected_year}.docx"
    try:
        save_to_word(all_data, doctors_by_dept, sorted_depts, docx_file,
                     rukovoditel_text, ploshadka_text, selected_month, selected_year)
        print(f"💾 Word сохранён: {docx_file}")
    except PermissionError as e:
        return False, files_with_errors, str(e)

    print(f"\n✅ Готово!")
    return True, files_with_errors, ""