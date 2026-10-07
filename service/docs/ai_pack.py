#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ai_pack.py — собирает «упаковку» проекта для передачи в чат с ИИ.
Запуск:  python service/docs/ai_pack.py   (из любой папки проекта)
Выход:   project_info/ai_pack/ в корне проекта (корень ищется по .git)
Что создаётся:
PROJECT_TREE.txt        — структура проекта (с лимитами на списки)
PROJECT_INFO.md         — сводка: файлы + классы/функции + docstring (ядро!)
CODE_root.md            — код файлов из корня проекта
CODE_<папка>.md         — по одной на каждую папку с кодом (дробится по 500 КБ)
CONFIG_TEMPLATES.md     — конфиги, схемы .sql, текстовые шаблоны (секреты маскируются)
GIT_STATE.txt           — последние коммиты, теги, статус репозитория
MANIFEST.md             — что создано, размеры, порядок загрузки в чат
Игнорируются целиком: .git, __pycache__, venv, data/, uploads/, exports/,
personal/, transfer/, _old_v2/, ai_pack/ и т.п.; списки имён ограничены.
"""
from __future__ import annotations

import ast
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable

# ================================================================
# НАСТРОЙКИ
# ================================================================
def _find_root(start: Path) -> Path:
    """Корень проекта: ближайший предок с .git; иначе папка скрипта."""
    for p in [start, *start.parents]:
        if (p / ".git").exists():
            return p
    return start.parent


BASE_DIR = _find_root(Path(__file__).resolve())
OUT_DIR = BASE_DIR / "project_info" / "ai_pack"

# Папки и файлы, которые не трогаем
IGNORE_DIRS = {
    ".git", ".github", "__pycache__", ".pytest_cache", ".mypy_cache",
    "dist", "build", "venv", ".venv", "env", ".env",
    ".idea", ".vscode", "node_modules",
    "data", "var", "uploads", "exports", "qr_static", "personal",
    "transfer", "_old_v2", "ai_pack", "_project_dump",
    "SonologData",
}
IGNORE_GLOB_DIRS = ["templates_backup_*", "backup_*", "restore_my"]
IGNORE_FILES = {
    ".DS_Store", "Thumbs.db",
}
IGNORE_GLOBS = ["*.pyc", "*.pyo", "*.log", "*.tmp",
                "*.db", "*.sqlite", "*.sqlite3",
                "*legacy_dump*", "*legacy_backup*", "*.gz"]

# Расширения, которые включаем в CODE_*.md (текстовые исходники)
CODE_EXTS = {".py"}
CODE_SPECIAL = {".bat", ".sh", ".spec", ".html", ".css", ".js"}
# Расширения, которые включаем в CONFIG_TEMPLATES.md
CONFIG_EXTS = {".json", ".ini", ".cfg", ".toml", ".yaml", ".yml", ".bat", ".spec", ".sql"}
TEMPLATE_EXTS = {".txt", ".md"}
# Бинарные — только имена, без содержимого, и с лимитом списка
BINARY_EXTS = {".docx", ".xlsx", ".pdf", ".png", ".jpg", ".jpeg", ".ico",
               ".gif", ".zip", ".7z", ".rar", ".exe", ".dll", ".so"}

# Лимиты
MAX_OUTPUT_FILE_BYTES = 500 * 1024       # 500 КБ — дробить на части
MAX_SOURCE_FILE_BYTES = 200 * 1024       # 200 КБ — обрезать один файл
MAX_FILES_PER_TREE_DIR = 15              # в дереве больше — «... ещё N ...»
TREE_PREVIEW_HEAD = 5                    # показываем столько первых
TREE_PREVIEW_TAIL = 2                    # и столько последних
TEMPLATE_SAMPLES_PER_GROUP = 3           # образцов содержимого на семейство .txt/.md
MAX_LISTED_NAMES = 30                    # максимум имён в любом списке-перечислении

# Секреты, которые вырезаем из конфигов
SECRET_KEY_PATTERNS = [
    re.compile(r'("(?:password|passwd|pwd|token|secret|api_key|apikey)"\s*:\s*")([^"]+)(")', re.I),
    re.compile(r'((?:password|passwd|pwd|token|secret|api_key|apikey)\s*=\s*)(\S+)', re.I),
]

LANG_MAP = {"py": "python", "html": "html", "css": "css", "js": "javascript",
            "bat": "bat", "sh": "bash", "spec": "python",
            "json": "json", "ini": "ini", "toml": "toml",
            "yaml": "yaml", "yml": "yaml", "cfg": "ini", "sql": "sql"}

# ================================================================
# УТИЛИТЫ
# ================================================================
def is_ignored_dir(path: Path) -> bool:
    name = path.name
    if name in IGNORE_DIRS:
        return True
    for pat in IGNORE_GLOB_DIRS:
        if path.match(pat):
            return True
    return False


def is_ignored_file(path: Path) -> bool:
    if path.name in IGNORE_FILES:
        return True
    for pat in IGNORE_GLOBS:
        if path.match(pat):
            return True
    return False


def _under_ignored(path: Path, root: Path) -> bool:
    """Лежит ли файл внутри игнорируемой подпапки."""
    parts = path.relative_to(root).parts
    for i in range(1, len(parts)):
        if is_ignored_dir(root / Path(*parts[:i])):
            return True
    return False


def iter_project_files(root: Path) -> Iterable[Path]:
    """Все файлы проекта, кроме игнорируемых."""
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        parts = path.relative_to(root).parts
        skip = False
        for i in range(1, len(parts)):
            if is_ignored_dir(root / Path(*parts[:i])):
                skip = True
                break
        if skip:
            continue
        if is_ignored_file(path):
            continue
        yield path


def human_size(n: int) -> str:
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "Б" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} ТБ"


def read_text_safe(path: Path) -> str | None:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        try:
            return path.read_text(encoding="cp1251")
        except Exception:
            return None
    except Exception:
        return None


def mask_secrets(text: str) -> str:
    for pat in SECRET_KEY_PATTERNS:
        text = pat.sub(lambda m: m.group(1) + "***" + (m.group(3) if m.lastindex >= 3 else ""), text)
    return text


# ================================================================
# 1. PROJECT_TREE.txt
# ================================================================
def build_tree(root: Path) -> str:
    lines = [f"{root.name}/"]

    def walk(d: Path, prefix: str = ""):
        try:
            entries = sorted(d.iterdir(), key=lambda p: (p.is_file(), p.name.lower()))
        except PermissionError:
            return
        dirs = [e for e in entries if e.is_dir() and not is_ignored_dir(e)]
        files = [e for e in entries if e.is_file() and not is_ignored_file(e)]
        for i, sub in enumerate(dirs):
            is_last_dir = (i == len(dirs) - 1) and not files
            sub_files = [f for f in sub.rglob("*")
                         if f.is_file() and not is_ignored_file(f)
                         and not _under_ignored(f, root)]
            total_kb = sum(f.stat().st_size for f in sub_files) / 1024
            count = len(sub_files)
            marker = "└──" if is_last_dir else "├──"
            lines.append(f"{prefix}{marker} {sub.name}/"
                         f"{' ' * max(1, 30 - len(sub.name))}"
                         f"({count} файлов, {total_kb:.1f} КБ)")
            walk(sub, prefix + ("    " if is_last_dir else "│   "))
        if len(files) > MAX_FILES_PER_TREE_DIR:
            head = files[:TREE_PREVIEW_HEAD]
            tail = files[-TREE_PREVIEW_TAIL:]
            middle_count = len(files) - TREE_PREVIEW_HEAD - TREE_PREVIEW_TAIL
            shown = head + [None] + tail
        else:
            shown = files
            middle_count = 0
        for j, f in enumerate(shown):
            is_last = (j == len(shown) - 1)
            marker = "└──" if is_last else "├──"
            if f is None:
                lines.append(f"{prefix}├── ... ещё {middle_count} однотипных (имена опущены) ...")
                continue
            size = human_size(f.stat().st_size)
            lines.append(f"{prefix}{marker} {f.name}"
                         f"{' ' * max(1, 34 - len(f.name))}({size})")

    walk(root)
    lines.append("")
    lines.append(f"# Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"# Игнорируются: {', '.join(sorted(IGNORE_DIRS))}")
    return "\n".join(lines)


# ================================================================
# 2. CODE_<папка>.md
# ================================================================
def collect_code_groups(root: Path) -> dict[str, list[Path]]:
    groups: dict[str, list[Path]] = {}
    for path in iter_project_files(root):
        if path.suffix.lower() not in CODE_EXTS and path.suffix.lower() not in CODE_SPECIAL:
            continue
        rel = path.relative_to(root)
        key = "root" if len(rel.parts) == 1 else "/".join(rel.parts[:-1])
        groups.setdefault(key, []).append(path)
    return groups


def make_code_filename(folder_key: str) -> str:
    safe = folder_key.replace("/", "_").replace("\\", "_")
    return f"CODE_{safe}.md"


def render_code_file(path: Path, root: Path) -> str:
    rel = path.relative_to(root).as_posix()
    size = path.stat().st_size
    text = read_text_safe(path)
    if text is None:
        return (f"## Файл: {rel}\n"
                f"Размер: {human_size(size)}\n"
                f"(не удалось прочитать как текст)\n")
    if size > MAX_SOURCE_FILE_BYTES:
        text = text[:MAX_SOURCE_FILE_BYTES]
        truncated = f"\n# ... обрезано, всего {size} байт ..."
    else:
        truncated = ""
    lang = LANG_MAP.get(path.suffix.lstrip(".").lower(), "text")
    lines_count = text.count("\n") + 1
    return (f"## Файл: {rel}\n"
            f"Размер: {human_size(size)}, строк: ~{lines_count}\n"
            f"```{lang}\n{text}{truncated}\n```\n")


def build_code_md(folder_key: str, files: list[Path], root: Path) -> list[tuple[str, str]]:
    header = (f"# Код: {folder_key}\n"
              f"Файлов: {len(files)}\n"
              f"Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
              f"---\n")
    parts: list[str] = [header]
    current_size = len(header.encode("utf-8"))
    part_num = 1
    results: list[tuple[str, str]] = []
    for path in files:
        chunk = render_code_file(path, root)
        chunk_size = len(chunk.encode("utf-8"))
        if current_size + chunk_size > MAX_OUTPUT_FILE_BYTES and len(parts) > 1:
            base = make_code_filename(folder_key)
            if part_num == 1:
                results.append((base, "\n".join(parts)))
            else:
                results.append((f"{base[:-3]}_part{part_num}.md", "\n".join(parts)))
            part_num += 1
            parts = [header]
            current_size = len(header.encode("utf-8"))
        parts.append(chunk)
        current_size += chunk_size
    if len(parts) > 1:
        base = make_code_filename(folder_key)
        if part_num == 1:
            results.append((base, "\n".join(parts)))
        else:
            results.append((f"{base[:-3]}_part{part_num}.md", "\n".join(parts)))
    return results


# ================================================================
# 3. CONFIG_TEMPLATES.md
# ================================================================
def _append_template_block(lines: list[str], path: Path, root: Path):
    rel = path.relative_to(root).as_posix()
    size = human_size(path.stat().st_size)
    text = read_text_safe(path)
    lines.append(f"#### {rel}  ({size})")
    lines.append("")
    if text is None:
        lines.append("(не читается как текст)")
    else:
        lines.append("```")
        lines.append(text.rstrip())
        lines.append("```")
    lines.append("")


def build_config_templates(root: Path) -> str:
    lines = [
        "# Конфиги и шаблоны",
        "",
        f"Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "---",
        "",
    ]
    configs: list[Path] = []
    templates: list[Path] = []
    binaries: list[Path] = []
    for path in iter_project_files(root):
        ext = path.suffix.lower()
        if ext in CONFIG_EXTS:
            configs.append(path)
        elif ext in TEMPLATE_EXTS:
            templates.append(path)
        elif ext in BINARY_EXTS:
            binaries.append(path)

    lines.append("## Конфиги")
    lines.append("")
    if not configs:
        lines.append("(нет)")
    for path in configs:
        rel = path.relative_to(root).as_posix()
        text = read_text_safe(path)
        if text is None:
            lines.append(f"### {rel}\n(не читается как текст)\n")
            continue
        text = mask_secrets(text)
        size = human_size(path.stat().st_size)
        lang = LANG_MAP.get(path.suffix.lstrip(".").lower(), "text")
        lines.append(f"### {rel}  ({size})")
        lines.append("")
        lines.append(f"```{lang}")
        lines.append(text.rstrip())
        lines.append("```")
        lines.append("")

    lines.append("## Шаблоны и текст")
    lines.append("")
    if not templates:
        lines.append("(нет)")
    else:
        groups: dict[str, list[Path]] = {}
        for t in templates:
            base = re.sub(r'[_\-\s]*\d{4,}.*$', '', t.stem)
            base = re.sub(r'[_\-\s]*(copy|backup).*$', '', base, flags=re.I)
            groups.setdefault(base, []).append(t)
        for base, files in sorted(groups.items()):
            files_sorted = sorted(files)
            if len(files_sorted) > MAX_FILES_PER_TREE_DIR:
                sample = files_sorted[:TEMPLATE_SAMPLES_PER_GROUP]
                lines.append(f"### {base}.* — {len(files_sorted)} файлов "
                             f"(содержимое показано у {len(sample)})")
                lines.append("")
                for f in sample:
                    _append_template_block(lines, f, root)
                others = [f.name for f in files_sorted[TEMPLATE_SAMPLES_PER_GROUP:]]
                shown = others[:MAX_LISTED_NAMES]
                tail = (f" … и ещё {len(others) - len(shown)} (имена опущены)"
                        if len(others) > len(shown) else "")
                lines.append(f"*Остальные ({len(others)}):* {', '.join(shown)}{tail}")
                lines.append("")
            else:
                lines.append(f"### {base}.* — {len(files_sorted)} файлов")
                lines.append("")
                for f in files_sorted:
                    _append_template_block(lines, f, root)

    if binaries:
        lines.append("## Бинарные файлы (не включены, только имена)")
        lines.append("")
        shown = binaries[:MAX_LISTED_NAMES]
        for b in shown:
            rel = b.relative_to(root).as_posix()
            lines.append(f"- `{rel}` ({human_size(b.stat().st_size)})")
        if len(binaries) > len(shown):
            lines.append(f"- *… и ещё {len(binaries) - len(shown)} файлов — "
                         f"имена опущены, чтобы не раздувать пакет.*")
        lines.append("")
    return "\n".join(lines)


# ================================================================
# 4. PROJECT_INFO.md — ядро для ИИ
# ================================================================
def parse_python_file(path: Path) -> dict:
    result = {"module_doc": "", "classes": [], "functions": [], "error": None}
    text = read_text_safe(path)
    if text is None:
        result["error"] = "не читается как текст"
        return result
    try:
        tree = ast.parse(text)
    except SyntaxError as e:
        result["error"] = f"SyntaxError: {e}"
        return result
    result["module_doc"] = ast.get_docstring(tree) or ""
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            result["classes"].append({
                "name": node.name,
                "doc": (ast.get_docstring(node) or "").split("\n")[0].strip(),
                "methods": [n.name for n in node.body if isinstance(n, ast.FunctionDef)],
            })
        elif isinstance(node, ast.FunctionDef):
            result["functions"].append({
                "name": node.name,
                "doc": (ast.get_docstring(node) or "").split("\n")[0].strip(),
            })
    return result


def build_project_info(root: Path) -> str:
    lines = [
        "# PROJECT_INFO — сводка по проекту",
        "",
        f"Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Корень проекта: `{root.name}`",
        "",
        "> Это **ядро** для передачи проекта в чат с ИИ. Сначала покажи этот файл,",
        "> затем — GIT_STATE.txt и нужные CODE_*.md / CONFIG_TEMPLATES.md.",
        "",
        "---",
        "",
        "## Оглавление",
        "",
    ]
    py_files = []
    other_code = []
    for path in iter_project_files(root):
        ext = path.suffix.lower()
        rel = path.relative_to(root).as_posix()
        if ext == ".py":
            py_files.append((path, rel, parse_python_file(path)))
        elif ext in CODE_SPECIAL:
            other_code.append((path, rel))

    lines.append("### Python-файлы")
    lines.append("")
    for _, rel, info in py_files:
        size = human_size((root / rel).stat().st_size)
        lines.append(f"- `{rel}` ({size})")
    lines.append("")
    if other_code:
        lines.append("### Прочий код и разметка")
        lines.append("")
        for _, rel in other_code:
            size = human_size((root / rel).stat().st_size)
            lines.append(f"- `{rel}` ({size})")
        lines.append("")
    lines.append("---")
    lines.append("")

    lines.append("## Файлы с описанием (классы, функции)")
    lines.append("")
    for path, rel, info in py_files:
        size = human_size(path.stat().st_size)
        lines.append(f"### `{rel}`  ({size})")
        lines.append("")
        if info["error"]:
            lines.append(f"⚠️ {info['error']}")
            lines.append("")
            continue
        doc = info["module_doc"].strip().split("\n")[0] if info["module_doc"] else ""
        if doc:
            lines.append(f"**Назначение:** {doc}")
            lines.append("")
        if info["classes"]:
            lines.append("**Классы:**")
            lines.append("")
            for cls in info["classes"]:
                doc_c = f" — {cls['doc']}" if cls["doc"] else ""
                lines.append(f"- `{cls['name']}`{doc_c}")
                if cls["methods"]:
                    m = ", ".join(f"`{x}`" for x in cls["methods"][:20])
                    tail = " …" if len(cls["methods"]) > 20 else ""
                    lines.append(f"  - методы: {m}{tail}")
            lines.append("")
        if info["functions"]:
            lines.append("**Функции верхнего уровня:**")
            lines.append("")
            for fn in info["functions"]:
                doc_f = f" — {fn['doc']}" if fn["doc"] else ""
                lines.append(f"- `{fn['name']}`{doc_f}")
            lines.append("")
        lines.append("---")
        lines.append("")

    lines.append("## Карта по папкам")
    lines.append("")
    groups = collect_code_groups(root)
    for key in sorted(groups.keys()):
        lines.append(f"- **{key}/** → {len(groups[key])} файлов "
                     f"(см. `{make_code_filename(key)}`)")
    lines.append("")
    return "\n".join(lines)


# ================================================================
# 5. GIT_STATE.txt
# ================================================================
def build_git_state(root: Path) -> str | None:
    import subprocess
    if not (root / ".git").exists():
        return None
    out = []
    for args in (["log", "--oneline", "-15"], ["status", "--short"], ["tag", "-l"]):
        try:
            r = subprocess.run(
                ["git"] + args, cwd=root,
                capture_output=True, text=True, timeout=15,
                encoding="utf-8", errors="replace",
            )
            out.append(f"$ git {' '.join(args)}\n{r.stdout.strip()}\n")
        except Exception:
            return None
    return "\n".join(out)


# ================================================================
# 6. MANIFEST.md
# ================================================================
def build_manifest(root: Path, created: list[tuple[str, int]]) -> str:
    lines = [
        "# MANIFEST — что лежит в project_info/ai_pack/",
        "",
        f"Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Проект: `{root.name}`",
        "",
        "| Файл | Размер |",
        "|------|--------|",
    ]
    for name, size in sorted(created):
        lines.append(f"| `{name}` | {human_size(size)} |")
    lines.append("")
    lines.append("## Что грузить в чат")
    lines.append("")
    lines.append("1. **Всегда первым:** `PROJECT_INFO.md` — общая карта проекта.")
    lines.append("2. **Версии и статус:** `GIT_STATE.txt`.")
    lines.append("3. **По необходимости:** `CODE_<папка>.md` — код нужной папки.")
    lines.append("4. **Для настроек и схем:** `CONFIG_TEMPLATES.md`.")
    lines.append("5. **Структура:** `PROJECT_TREE.txt`.")
    lines.append("")
    lines.append("## Что НЕ включено (специально)")
    lines.append("")
    lines.append("- Каталоги данных: `data/`, `uploads/`, `exports/`, `personal/`,")
    lines.append("  `transfer/`, `_old_v2/` и подобные — целиком.")
    lines.append("- Списки имён длиннее 30 позиций урезаны с пометкой «имена опущены».")
    lines.append("- Бинарные файлы — только имена (до 30 штук).")
    lines.append("- `.git/`, `__pycache__/`, `dist/`, `build/`, `venv/`.")
    lines.append("- Секреты в конфигах заменены на `***`.")
    return "\n".join(lines)


# ================================================================
# MAIN
# ================================================================
def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for old in OUT_DIR.glob("*"):
        if old.is_file():
            old.unlink()
    print(f"📦 Собираю ai_pack/ в {OUT_DIR}")
    print()
    created: list[tuple[str, int]] = []

    tree = build_tree(BASE_DIR)
    p = OUT_DIR / "PROJECT_TREE.txt"
    p.write_text(tree, encoding="utf-8")
    created.append(("PROJECT_TREE.txt", p.stat().st_size))
    print(f"✅ PROJECT_TREE.txt  ({human_size(p.stat().st_size)})")

    gs = build_git_state(BASE_DIR)
    if gs:
        p = OUT_DIR / "GIT_STATE.txt"
        p.write_text(gs, encoding="utf-8")
        created.append(("GIT_STATE.txt", p.stat().st_size))
        print(f"✅ GIT_STATE.txt  ({human_size(p.stat().st_size)})")

    groups = collect_code_groups(BASE_DIR)
    for key in sorted(groups.keys()):
        files = groups[key]
        for name, content in build_code_md(key, files, BASE_DIR):
            out = OUT_DIR / name
            out.write_text(content, encoding="utf-8")
            created.append((name, out.stat().st_size))
            print(f"✅ {name}  ({human_size(out.stat().st_size)}, файлов: {len(files)})")

    cfg = build_config_templates(BASE_DIR)
    p = OUT_DIR / "CONFIG_TEMPLATES.md"
    p.write_text(cfg, encoding="utf-8")
    created.append(("CONFIG_TEMPLATES.md", p.stat().st_size))
    print(f"✅ CONFIG_TEMPLATES.md  ({human_size(p.stat().st_size)})")

    info = build_project_info(BASE_DIR)
    p = OUT_DIR / "PROJECT_INFO.md"
    p.write_text(info, encoding="utf-8")
    created.append(("PROJECT_INFO.md", p.stat().st_size))
    print(f"✅ PROJECT_INFO.md  ({human_size(p.stat().st_size)})")

    manifest = build_manifest(BASE_DIR, created)
    p = OUT_DIR / "MANIFEST.md"
    p.write_text(manifest, encoding="utf-8")
    print(f"✅ MANIFEST.md  ({human_size(p.stat().st_size)})")
    print()
    print(f"🎉 Готово. Файлов: {len(created) + 1}")
    print(f"📁 Смотри: {OUT_DIR}")


if __name__ == "__main__":
    main()