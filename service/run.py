#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Запускатель служебных скриптов из любых папок проекта.
Использование:  python service/run.py service/migrations/migrate_v3_9.py [аргументы]
Добавляет корень проекта в sys.path, чтобы скрипты видели боевые модули.
"""
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if len(sys.argv) < 2:
    print(__doc__)
    sys.exit(1)
target = (ROOT / sys.argv[1]).resolve()
if not target.exists():
    print(f"Файл не найден: {target}")
    sys.exit(1)
sys.path.insert(0, str(ROOT))
sys.argv = [str(target)] + sys.argv[2:]
runpy.run_path(str(target), run_name="__main__")