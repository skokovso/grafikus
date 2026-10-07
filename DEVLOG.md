# DEVLOG — журнал сессий разработки Grafikus

Фиксируем всё, что делали, включая неудачные ветки. Цель — не потерять
контекст и понимать, какие блоки кода в какой момент появлялись и пропадали.

## Исходная точка

- Монолит `grafikus.py` (~2000 строк, 85 КБ): всё в одном файле.
- Интерфейс на Tkinter, затем CustomTkinter.
- Парсеры Excel/PDF/TXT внутри монолита.

## Хронология (по тегам git)

- `start`, `restart` — первые версии.
- `Parsing surgery` — добавлен парсер хирургии.
- `icon` — иконка приложения.
- `ver 6_0 new gui` — переход на CustomTkinter.
- `ver 7 final` — стабилизация v7.0.
- `ver 7.1 word open error` — обработка `PermissionError` (Word).
- `instruction, readme` — первые документы.
- `ver 7.2 Volodarskogo File` — парсер многостраничного Excel.
- `ver 7.3 click open file` — клик по статусу открывает папку.
- `ver 7.4 V18 and L18` — правки парсинга хирургии.
- `v 8.0 Финальная версия рабочая перед переносом` — рефакторинг на модули.

## Сессия 2026-10-01 — v8.0

### Рефакторинг на модули

- `core/config.py` — DEPT_ABBR, DEPARTMENT_ORDER, MONTHS_RU*.
- `core/utils.py` — normalize_time, extract_*, find_month_in_text,
  read_excel_or_csv.
- `core/doctor_validator.py` — is_valid_doctor_name.
- `core/schedule_builder.py` — build_master_schedule.
- `parsers/base.py` — get_dept_abbr, get_dept_order, detect_department.
- `parsers/standard.py` — parse_standard_graph, parse_responsible_graph.
- `parsers/oar.py` — parse_oar_graph.
- `parsers/surgery.py` — parse_surgery_pdf (extract_tables), parse_surgery_txt.
- `parsers/multisheet.py` — parse_excel_with_months.
- `output/word_builder.py` — save_to_word.
- `output/txt_builder.py` — save_to_txt.
- `ui/app.py` — класс App на CustomTkinter.

### Исправленные баги

- **Кодиров и слово «Код».** В `is_valid_doctor_name` была проверка
  `if 'Код' in name: return False`, из-за чего терялись врачи с фамилией
  «Кодиров». Заменили на точное совпадение / `startswith(h + ' ')`.
- **PDF-хирургия и второй врач дня.** В старой `_parse_surgery_text`
  второй врач дня «переезжал» на следующий день (сдвиг из-за порядка
  строк). Причина — некорректная привязка дат. Решение:
  `pdfplumber.extract_tables()` + `_parse_surgery_table` с запоминанием
  последней даты (`last_day`). Старая версия сохранена в
  `parsers/surgery_first.py`.
- **ОАР: буква после времени.** В старом коде было `16-09А`, нужно `А 16-09`.
  Поправили в `parse_oar_graph`.
- **Регистр месяца.** Поиск `Октябрь/октябрь/ОКТЯБРЬ` — через
  `find_month_in_text` (lowercase).

## Сессия 2026-10-04 — упаковка для ИИ и документация

- Написан `service/docs/ai_pack.py` — упаковка проекта в `project_info/ai_pack/`.
- Написан `service/docs/make_docs.py` (первая версия, fpdf2-путь).
- Написан `service/docs/manuals_to_pdf.py` — fpdf2-рендер markdown.
- Написан `service/run.py` — запускатель служебных скриптов.
- Добавлены `AI_BRIEF.md`, `ARCHITECTURE.md`, `DIGEST.md`.

## Сессия 2026-10-06 — v8.1, Word-путь

- `make_docs.py` переписан по образцу sonolog:
  - каркас `_find_root` / `BASE_DIR` / `MANUALS_DIR` / `ASSETS_DIR`;
  - словарь `DOCS` (README, DIGEST, AI_BRIEF, ARCHITECTURE, CHANGELOG,
    DEVLOG, TODO, QWEN, USER_MANUAL);
  - `render_md_to_docx` (python-docx + логотип + авто-нумерация);
  - `render_md_to_pdf` (docx2pdf через Word → fallback fpdf2);
  - режимы `--only-missing` / обычный (перезапись).
- Исправлена двойная папка `project_info/manuals/manuals/`.
- Введена папка `assets/` с `project_logo.png` и `project_logo.ico`.

## Сессия 2026-10-07 — v8.1.1, администрация

### Что делали

- Диагностика `График_Администрация.xlsx`: файл содержит формулы
  `=DATE(...)` и `=IF(A4="", "", IF(MONTH(A4+1)=MONTH($A$4),A4+1, ""))`
  в колонке A, но Excel сохранил кэш — `pandas.read_excel` возвращает
  готовые `datetime`.

### Исправленные баги

- **`parse_administration_graph` возвращал 0 строк.** Проверка
  `isinstance(date_val, (pd.Timestamp, pd.DatetimeTZDtype))` —
  `pd.DatetimeTZDtype` это dtype-класс, а не тип значения; обычный
  `datetime` не проходит. Обе ветки (`if`/`elif isinstance(..., str))`
  проваливаются, `day_num` остаётся `None`. Исправлено на
  `isinstance(date_val, (datetime, pd.Timestamp))`, добавлен
  `from datetime import datetime`.
- **Артефакт «А дм»** в выводе: значение `'Адм'` интерпретировалось как
  буква смены `'А'` + время `'дм'` в `txt_builder`/`word_builder`.
  Заменено на пустую строку `''`, а билдеры теперь отдельно обрабатывают
  `dept == 'АДМ'`: без времени, со звёздочкой в TXT, жирным в Word.
- **Жирный шрифт** для администрации не срабатывал: проверка
  `" (АДМ) " in line` не находила `"(АДМ)"` без пробелов после `)`. Заменена
  на `"(АДМ)" in line`.

### Изменения

- `core/config.py`: `'АДМ'` первым в `DEPARTMENT_ORDER` — администрация
  идёт сразу после даты и дня недели.
- `parsers/administration.py`: значение `''` вместо `'Адм'`.
- `output/txt_builder.py`: для `dept == 'АДМ'` — `* ФИО (АДМ)` без времени.
- `output/word_builder.py`: для `dept == 'АДМ'` — `ФИО (АДМ)` без времени,
  жирным.

### Процессные выводы

- `pd.DatetimeTZDtype` — ловушка. Использовать только `datetime` +
  `pd.Timestamp` в `isinstance`.
- Значение в `all_data` не должно начинаться с `'А'`/`'Р'`/`'Э'`, если
  это не буква смены ОАР — иначе билдеры принимают его за букву.
- Перед коммитом — `python -m py_compile` по изменённым файлам (иначе
  `NameError` вылезет в момент парсинга через GUI).

## Неудачные ветки и потери (важно!)

- `parsers/surgery_first.py` — первая попытка исправить PDF-хирургию;
  оставлена в репозитории для истории (в импортах не используется).
- Монолит `grafikus.py` — после рефакторинга в v8.0 не удаляется: служит
  справочником и страховкой, пока новая модульная версия обкатывается.
- Дубли парсеров: `parse_responsible_graph` в `parsers/standard.py` —
  тонкая обёртка над `parse_standard_graph`, оставлена для совместимости.

## Процессные выводы

- Новый парсер — только в `parsers/`, монолит не трогаем.
- `is_valid_doctor_name` — святое, правим с осторожностью.
- Реальные ФИО — только в `data/`, никаких выгрузок в git.
- Документы генерируются скриптом, руками не правятся.
