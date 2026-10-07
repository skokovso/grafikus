# ARCHITECTURE — Grafikus (v8.1.1)

Автоматизация сбора разрозненных графиков дежурств (Excel, CSV, PDF, TXT)
в единый документ Word для утверждения главным врачом.

## 1. Слои

- **`main.py`** — точка входа. Готовит `sys.path` (важно для PyInstaller),
  настраивает CustomTkinter, импортирует `ui.app.App`.
- **`ui/app.py`** — интерфейс (CustomTkinter). Отображение, справочники,
  выбор папки/периода/руководителя/площадки, запуск оркестратора.
- **`core/`** — ядро без GUI:
  - `config.py` — константы, `DEPT_ABBR`, `DEPARTMENT_ORDER` (АДМ первым),
    месяцы.
  - `utils.py` — `normalize_time`, `extract_*_from_date`, `find_month_in_text`,
    `read_excel_or_csv`.
  - `doctor_validator.py` — `is_valid_doctor_name` (отсев заголовков).
  - `schedule_builder.py` — `build_master_schedule`: сканирует папку,
    определяет отделение, вызывает нужный парсер, агрегирует данные,
    проверяет месяц/год, отдаёт в `output`.
- **`parsers/`** — чтение файлов:
  - `base.py` — `get_dept_abbr`, `get_dept_order`, `detect_department`.
  - `administration.py` — `parse_administration_graph`: `Дата, день недели`
    + площадки в шапке; `datetime` из pandas/openpyxl; значение `''`.
  - `standard.py` — обычные Excel/CSV; `parse_standard_graph`,
    `parse_responsible_graph` (последний = тонкая обёртка).
  - `oar.py` — `parse_oar_graph` (буквы А/Р/Э, время из предпоследней строки).
  - `surgery.py` — `parse_surgery_pdf` (через `extract_tables`),
    `parse_surgery_txt`; сохраняет извлечённые таблицы в `_extracted_tables.xlsx`.
  - `multisheet.py` — `parse_excel_with_months` (Володарского: листы по месяцам).
- **`output/`** — запись результата:
  - `word_builder.py` — `save_to_word`: 3 колонки, шапка, страницы 48/59.
  - `txt_builder.py` — `save_to_txt`: быстрый дамп для проверки.

## 2. Поток данных

1. Пользователь выбирает папку, месяц, год, руководителя, площадку.
2. `build_master_schedule` собирает все `.xlsx/.xls/.csv/.pdf/.txt`
   (исключая файлы, начинающиеся с `~`, содержащие `_extracted.txt`,
   и файлы со словом «сводный»).
3. Для Excel-файлов проверяется многостраничность (листы по месяцам,
   кроме ОАР-файлов) — если да, идёт `parse_excel_with_months`.
4. Иначе `detect_department` → выбор парсера:
   - `администрация` в имени → `parse_administration_graph` (с площадкой);
   - `ОАР` в имени → `parse_oar_graph`;
   - `ответственн` в имени/отделении → `parse_responsible_graph`;
   - `.pdf` → `parse_surgery_pdf`;
   - `.txt` → `parse_surgery_txt`;
   - иначе → `parse_standard_graph`.
5. Проверка месяца/года. Несовпадение → в `files_with_errors`, данные не
   подмешиваются.
6. `all_data[day][(doctor, dept)] = time`.
7. Сортировка отделений по `dept_order_map` → `sorted_depts`
   (АДМ получает порядок 0 и идёт первым).
8. Сохранение TXT и Word в `<input_folder>/Сводный график/`.

## 3. Формат данных

- `dict[day: int, dict[(doctor: str, dept: str), time: str]]`.
- `time`: `ЧЧ-ЧЧ`, `ЧЧ:ММ-ЧЧ:ММ`, `<буква><время>` для ОАР (например,
  `А16-09` в `all_data`, при выводе — `А 16-09`), либо `''` для
  администрации.
- `doctors_by_dept[dept]` — список уникальных `(doctor, dept)`.

## 4. Печатная форма Word

- Шапка: `УТВЕРЖДАЮ`, должность, ФИО, дата (сегодня, в родительном падеже).
- Заголовок: `ГРАФИК ДЕЖУРСТВА ВРАЧЕЙ СТАЦИОНАРА ГОБУЗ МОКМЦ`, площадка,
  `НА <МЕСЯЦ> <ГОД> Г.`.
- Таблица 3×N без видимых границ (все `w:tcBorders` = `none`).
- Блок дня: `ДД.ММ ДеньНедели` жирным + строки `Фамилия (ОТД) время`.
- **Администрация**: строка `ФИО (АДМ)` (без времени), жирным; идёт
  первой после блока дня.
- Лимиты: 1-я страница 48 строк, остальные 59; `add_page_break` между
  страницами.
- `PermissionError` при сохранении → сообщение «закройте файл в Word».

## 5. Безопасность и приватность

- Входные данные с ФИО врачей хранятся в `data/` — вне git.
- Парсинг PDF локальный, никаких сетевых запросов.
- Справочники (`spravochnik_rukovoditeli.txt`, `spravochnik_ploshadki.txt`)
  лежат рядом с exe, в git попадают только примеры.

## 6. Миграция с монолита (для истории)

- `grafikus.py` (85 КБ) — старая монолитная версия (v7.x), оставлена в
  корне для справки и обратной совместимости некоторых функций.
- v8.0 разбила его на `core/`, `parsers/`, `output/`, `ui/`.
- `parsers/surgery_first.py` — предыдущая версия парсера хирургии, оставлена
  для истории (в импортах не используется).

## 7. Известные особенности

- PyInstaller: `main.py` добавляет `sys._MEIPASS` в `sys.path` — без этого
  падают импорты `core`/`parsers`/`ui` в exe.
- CustomTkinter в тёмной теме по умолчанию (`set_appearance_mode("dark")`).
- Word-документ может «молча» не перезаписаться, если открыт — ловим
  `PermissionError`.
- Многостраничный Excel определяется по именам листов-месяцев; ОАР-файл
  с таким именем не считается многостраничным (у него особый парсер).
- **Администрация**: в файле формулы `=DATE(...)`/`=IF(...)` в колонке A;
  `pandas.read_excel` возвращает уже вычисленные `datetime` (кэш Excel),
  если файл сохранён Excel'ем. `isinstance(date_val, (datetime,
  pd.Timestamp))` — единственно верная проверка (`pd.DatetimeTZDtype`
  ломается на обычных `datetime`).
