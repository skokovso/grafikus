# output/txt_builder.py
"""Сохранение сводного графика в текстовый формат (.txt) для быстрой проверки."""
import calendar
import datetime
from pathlib import Path
from core.config import MONTHS_RU, DEPT_ABBR
from core.utils import normalize_doctor_name_for_output, round_time_up_for_output

def get_dept_abbr_from_name(dept_name: str) -> str:
    if dept_name in DEPT_ABBR:
        return DEPT_ABBR[dept_name]
    for key, abbr in DEPT_ABBR.items():
        if key.lower() in dept_name.lower() or dept_name.lower() in key.lower():
            return abbr
    words = dept_name.split()
    if len(words) >= 2:
        return ''.join(w[0].upper() for w in words[:2])
    return dept_name[:4].upper()

def save_to_txt(all_data, doctors_by_dept, sorted_depts, output_file, month_num, year):
    """Сохраняет сводный график в TXT."""
    _, last_day = calendar.monthrange(year, month_num)
    weekday_names = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']
    weekdays = [weekday_names[datetime.date(year, month_num, day).weekday()] for day in range(1, last_day + 1)]

    output_lines = []
    for day in sorted(all_data.keys()):
        if 1 <= day <= last_day:
            day_name = weekdays[day-1]
            output_lines.append(f"{day:02d}.{month_num:02d} {day_name}")
            for dept in sorted_depts:
                dept_doctors = sorted(doctors_by_dept.get(dept, []), key=lambda x: x[0])
                abbr = get_dept_abbr_from_name(dept)
                is_travma_zav = (dept == 'Травматология (зав.)')

                for doctor, dept_key in dept_doctors:
                    if (doctor, dept_key) in all_data[day]:
                        value = all_data[day][(doctor, dept_key)]

                        # Нормализация ФИО для травматологии заведующего
                        out_doctor = normalize_doctor_name_for_output(doctor) if is_travma_zav else doctor

                        # Администрация выводится без времени
                        if dept == 'АДМ' or dept_key == 'АДМ':
                            output_lines.append(f"* {out_doctor} ({abbr})")
                            continue

                        # Парсим значение: если начинается с А/Р/Э — буква после скобки
                        letter = ''
                        time_part = value
                        if value and value[0] in 'АРЭ':
                            letter = value[0]
                            time_part = value[1:]

                        # Округление времени для травматологии заведующего
                        if is_travma_zav:
                            time_part = round_time_up_for_output(time_part)

                        if letter:
                            output_lines.append(f"  {out_doctor} ({abbr}){letter} {time_part}")
                        else:
                            output_lines.append(f"  {out_doctor} ({abbr}) {time_part}")
            output_lines.append("")

    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(output_lines))
        print(f"💾 TXT сохранён: {output_file}")
    except Exception as e:
        print(f"⚠️ Ошибка сохранения TXT: {e}")