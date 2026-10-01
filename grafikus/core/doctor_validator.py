# core/doctor_validator.py
"""Валидация имен врачей и отсеивание заголовков."""
import re

def is_valid_doctor_name(name: str) -> bool:
    """Проверяет, является ли строка именем врача, а не заголовком."""
    if not name or not isinstance(name, str):
        return False
    name = name.strip()
    if not name:
        return False
        
    # ПРОВЕРКА НА ЗАГОЛОВКИ (ИСПРАВЛЕННЫЙ БАГ С "КОД")
    headers = [
        'Дата', 'Дежурный врач', 'Сб', 'Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт',
        'График', 'Утверждаю', 'Заведующий', 'Ф.И.О.', 'цех', 'отделение',
        '2026', '2025', 'Август', 'Март', 'ФИО', '№', 'п/п', 'Код'
    ]
    for h in headers:
        # Проверяем точное совпадение или начало строки с пробелом
        if name == h or name.startswith(h + ' '):
            return False
            
    # Исключаем строки с датами
    if re.search(r'\d{4}-\d{2}-\d{2}', name):
        return False
        
    # Имя врача должно содержать буквы и быть не слишком длинным
    if len(name) > 50:
        return False
        
    # Должна быть хотя бы одна русская буква
    if not re.search(r'[а-яА-ЯёЁ]', name):
        return False
        
    return True