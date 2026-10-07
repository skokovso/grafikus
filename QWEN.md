# QWEN.md — контекст ассистента по проекту Grafikus (v8.1.1)

## Пользователь и стиль

- Автор проекта — врач/разработчик; интерфейс и сообщения — на русском.
- Код: полные файлы — только по явной просьбе; иначе точечные блоки
  «найди/замени» с точными отступами (4 пробела на уровень).
- Документы ведём скриптами-генераторами (`service/docs/make_docs.py`).
- Если даются два варианта одного блока — брать последний (самый полный).

## Текущее состояние (v8.1.1)

- Архитектура: см. `ARCHITECTURE.md`.
- Журнал сессий и потери: см. `DEVLOG.md`. Открытые задачи: см. `TODO.md`.
- Быстрый вход в контекст: см. `DIGEST.md`.
- Правила для ИИ: см. `AI_BRIEF.md`.

## Маркеры целостности core/

- `config.py`: `DEPT_ABBR` (включая `'АДМ'`), `DEPARTMENT_ORDER`
  (`'АДМ'` первым), `MONTHS_RU`, `MONTHS_RU_NOMINATIVE`,
  `MONTHS_RU_UPPER`, `MONTHS_SEARCH`.
- `utils.py`: `normalize_time`, `extract_day_from_date`,
  `extract_month_from_date`, `extract_year_from_date`,
  `find_month_in_text`, `read_excel_or_csv`.
- `doctor_validator.py`: `is_valid_doctor_name` — точное совпадение или
  `startswith(h + ' ')`; **никаких `in name` по списку headers**, иначе
  «Код» съест «Кодирова».
- `schedule_builder.py`: `build_master_schedule` возвращает
  `(success: bool, files_with_errors: list, error_message: str)`.

## Маркеры целостности parsers/

- `base.py`: `get_dept_abbr`, `get_dept_order`, `detect_department`.
- `administration.py`: `parse_administration_graph(file_path,
  selected_location)`. Ищет в шапке одну из четырёх площадок
  (`ломоносова, д.18`, `володарского, д.18`, `перинатальный центр`,
  `родильный дом`). Даты — `datetime` из pandas; проверка типа:
  `isinstance(date_val, (datetime, pd.Timestamp))`; **не** `pd.DatetimeTZDtype`.
  Значение — `''` (без времени).
- `standard.py`: `parse_standard_graph`, `parse_responsible_graph`.
  Ищет строку `Дата` с датами (ISO или ДД.ММ.ГГГГ), иначе — строку с
  ≥ 10 числами; шапка — `Дежурный врач`.
- `oar.py`: `parse_oar_graph`. Шапка — `Ф.И.О.`; время из предпоследней
  строки формата `\d{2}/\d{2}`; буква А/Р/Э **перед** временем.
- `surgery.py`: `parse_surgery_pdf` (extract_tables + `_parse_surgery_table`
  с `last_day`), `parse_surgery_txt`, `_parse_surgery_text` (построчный).
  Сохраняет `*_extracted_tables.xlsx` рядом с PDF для отладки.
- `multisheet.py`: `parse_excel_with_months`. Ключ листа — месяц; известные
  врачи (`KNOWN_DOCTORS`) — на случай, если `is_valid_doctor_name` не
  пропустил редкую фамилию.

## Маркеры целостности output/

- `word_builder.py`: `save_to_word`. 3 колонки, `w:tcBorders` = none,
  `FIRST_PAGE_MAX_ROWS=48`, `OTHER_PAGE_MAX_ROWS=59`,
  `distribute_blocks_with_limit`, `fill_column`. Формат блока:
  `ДД.ММ ДеньНедели` + строки `Фамилия (ОТД) время`. Для `dept == 'АДМ'`:
  `Фамилия (АДМ)` (без времени), `run.font.bold = True` при
  `"(АДМ)" in line`.
- `txt_builder.py`: `save_to_txt` — тот же формат, plain text. Для
  `dept == 'АДМ'`: `* Фамилия (АДМ)` (со звёздочкой, без времени).

## Маркеры целостности ui/

- `app.py`: класс `App(ctk.CTk)`; методы `create_widgets`, `set_status`,
  `clear_status`, `show_excluded_files`, `refresh_files_list`,
  `select_folder`, `on_status_click`, `open_output_folder`, `run`.
- `get_base_path()` — учитывает `sys.frozen` (PyInstaller).
- `load_spravochnik` / `save_spravochnik` — рядом с exe.
- Иконка: `assets/project_logo.ico` → `icon.ico` (fallback).

## Типовые ловушки

- `is_valid_doctor_name` с `if h in name` ломает «Кодирова». Только
  `name == h or name.startswith(h + ' ')`.
- PDF-хирургия: второй врач дня может «уехать» на следующий день при
  построчном парсинге. Спасает `extract_tables` + `last_day`.
- ОАР: буква смены должна быть **перед** временем, `А 16-09`.
- Администрация: значение **не** должно начинаться с `'А'`/`'Р'`/`'Э'`
  (иначе билдеры примут его за букву смены). Используем `''`.
- `pd.DatetimeTZDtype` — не тип значения, а dtype; в `isinstance` не
  работает на обычных `datetime`. Только `(datetime, pd.Timestamp)`.
- Excel-Володарского: многостраничность — по именам листов-месяцев;
  ОАР-файл исключение, у него свой парсер.
- `PermissionError` при `doc.save` — Word держит файл; ловим и сообщаем.
- PyInstaller: `sys._MEIPASS` в `sys.path`, иначе падают импорты.
- Перед коммитом — `python -m py_compile <изменённые файлы>`: `NameError`
  (забытый импорт) иначе всплывёт только в рантайме.

## Правила данных

- `data/` — вне git (реальные ФИО и графики).
- `spravochnik_*.txt` — рядом с exe; в git только примеры.
- Один файл = одно отделение (кроме многостраничного Excel).
- Word-файл не перезаписываем, если он открыт.

## Правила работы с документами

- Тексты правятся ТОЛЬКО в словаре `DOCS` файла
  `service/docs/make_docs.py`.
- Запуск: `python service/docs/make_docs.py` (перезапишет всё) или
  `python service/docs/make_docs.py --only-missing`.
- PDF и docx инструкции генерируются автоматически при запуске
  `make_docs.py` (Word-путь, fallback — fpdf2).
