# service/diagnostic/diag_month.py
"""Диагностика поиска месяца в тексте."""
import sys
from pathlib import Path

# service/diagnostic/diag_month.py → корень проекта
BASE_DIR = Path(__file__).resolve().parents[2]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from core.config import MONTHS_SEARCH
from core.utils import find_month_in_text


def main():
    print(f"Корень проекта: {BASE_DIR}")
    print()
    print("MONTHS_SEARCH =", MONTHS_SEARCH)
    print()

    text = "ГРАФИК работы врачей травматологического отделения ГОБУЗ МОКМЦ на ОКТЯБРЬ 2026 года"
    print(f"Текст: {text!r}")
    print(f"find_month_in_text → {find_month_in_text(text)!r}")
    print()

    tl = text.lower()
    print("Все совпадения ключей:")
    for name, num in MONTHS_SEARCH.items():
        if name in tl:
            print(f"  НАЙДЕНО: {name!r} → {num}")
    print()

    print("Проверка каждого ключа:")
    for name, num in MONTHS_SEARCH.items():
        print(f"  {name!r:>12} in text.lower() = {name in tl}")


if __name__ == "__main__":
    main()