# service/diagnostic/diag_travma.py
"""Диагностика формата графика травматологии (заведующего).

Запуск из корня проекта:
    python service/diagnostic/diag_travma.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_FILE = BASE_DIR / "data" / "input" / "_1_Графики дежурств от Отделений _ разбор" / "2026 10 Октябрь" / "Траваматология.xlsx"


def main() -> None:
    # Путь можно передать аргументом: python diag_travma.py <путь.xlsx>
    if len(sys.argv) > 1:
        fp = Path(sys.argv[1])
    else:
        fp = DEFAULT_FILE

    print(f"Корень проекта: {BASE_DIR}")
    print(f"Читаю файл:    {fp}")
    print(f"Существует:    {fp.exists()}")
    if not fp.exists():
        # Попробуем найти файл по имени во всём проекте
        candidates = list(BASE_DIR.rglob("Траваматология.xlsx"))
        if candidates:
            print(f"\nНашёл похожие файлы:")
            for c in candidates:
                print(f"  {c}")
            fp = candidates[0]
            print(f"\nБеру первый: {fp}")
        else:
            print("Файл не найден. Укажите путь аргументом.")
            return

    df = pd.read_excel(fp, header=None)
    print(f"\nФорма: {df.shape} (строк × колонок)")
    print("=" * 70)

    # Первые 12 строк — самое важное: шапка, дни недели, числа, врачи
    for i in range(min(12, len(df))):
        row = df.iloc[i].values
        print(f"\n--- row {i} ---")
        for j, v in enumerate(row[:35]):
            if pd.notna(v) and str(v).strip():
                print(f"  col {j:>2}: type={type(v).__name__:<10} "
                      f"repr={repr(v)[:70]}")

    print()
    print("=" * 70)
    print("Отдельно: колонка A (столбец ФИО)")
    print("=" * 70)
    for i in range(min(40, len(df))):
        v = df.iloc[i, 0]
        if pd.notna(v) and str(v).strip():
            print(f"  row {i:>2}: type={type(v).__name__:<10} "
                  f"repr={repr(v)[:70]}")


if __name__ == "__main__":
    main()