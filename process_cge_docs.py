#!/usr/bin/env python3
"""
Автоматизация правок XML для ЦГиЭ Документы.

Сценарий:
1) спрашивает у пользователя папку (например, 0126) — работает в ней;
2) удаляет все файлы, начинающиеся с «konvert» (рекурсивно);
3) в ON_NSCHFDOPPR* меняет СвПокуп/СвЮЛУч/@НаимОрг;
4) в ON_SCHET* меняет Покупатель/@Название и Покупатель/СвЮЛ/@Название;
5) переносит адрес ГрузПолуч → Грузополучатель по совпадению КПП (АдрТекст в двух местах).

Работает с двумя схемами структуры данных:
- старая: <дата>/УПД/Отправляемые и <дата>/Счет на оплату/Отправляемые;
- новая: <дата>/Отправляемые (оба типа файлов в одной папке).

Кодировка файлов сохраняется windows-1251. Вопрос только один: «Где продолжить работу?».
"""

from pathlib import Path
import re
import sys

TARGET_NAME = 'ФБУЗ &quot;Центр гигиены и эпидемиологии в городе Москве&quot;'

def ask_root(base: Path) -> Path:
    folder = input("Где продолжить работу? ").strip()
    if not folder:
        sys.exit("Не указана папка.")
    candidate = (base / folder).resolve()
    if not candidate.is_dir():
        sys.exit(f"Папка {candidate} не найдена.")
    return candidate

def delete_konvert(root: Path) -> int:
    removed = 0
    for path in root.rglob("*"):
        if path.is_file() and path.name.lower().startswith("konvert"):
            path.unlink()
            removed += 1
    return removed

def replace_nsch_names(upd_dir: Path) -> int:
    if not upd_dir:
        return 0
    pattern = re.compile(
        r'(<СвПокуп[\s\S]*?<СвЮЛУч[^>]*?НаимОрг=")[^"]*"[^>]*?(ИННЮЛ="\d+"\s+КПП="\d{9}")',
        re.DOTALL,
    )
    changed = 0
    for path in sorted(upd_dir.glob("ON_NSCHFDOPPR*.xml")):
        text = path.read_bytes().decode("windows-1251")
        new_text, n = pattern.subn(lambda m: f'{m.group(1)}{TARGET_NAME}" {m.group(2)}', text, count=1)
        if n and new_text != text:
            path.write_bytes(new_text.encode("windows-1251"))
            changed += 1
    return changed

def replace_schet_names(schet_dir: Path) -> int:
    if not schet_dir:
        return 0
    pat_buyer = re.compile(r'(<Покупатель[^>]*?Название=")([^"]*)(")')
    pat_legal = re.compile(
        r'(<Покупатель[\s\S]*?<СвЮЛ[^>]*?Название=")([^"]*)(")',
        re.DOTALL,
    )
    changed = 0
    for path in sorted(schet_dir.glob("ON_SCHET__*.xml")):
        text = path.read_bytes().decode("windows-1251")
        original = text
        text, n1 = pat_buyer.subn(lambda m: m.group(1) + TARGET_NAME + m.group(3), text, count=1)
        text, n2 = pat_legal.subn(lambda m: m.group(1) + TARGET_NAME + m.group(3), text, count=1)
        if text != original:
            path.write_bytes(text.encode("windows-1251"))
            changed += 1 if (n1 or n2) else 0
    return changed

def collect_kpp_to_addr(upd_dir: Path) -> dict:
    if not upd_dir:
        return {}
    kpp_re = re.compile(r'<ГрузПолуч[\s\S]*?<СвЮЛУч[^>]*?КПП="(\d{9})"', re.DOTALL)
    addr_re = re.compile(
        r'<ГрузПолуч[\s\S]*?<Адрес[\s\S]*?АдрИнф[^>]*?АдрТекст="([^"]+)"',
        re.DOTALL,
    )
    mapping = {}
    for path in sorted(upd_dir.glob("ON_NSCHFDOPPR*.xml")):
        text = path.read_bytes().decode("windows-1251")
        km = kpp_re.search(text)
        am = addr_re.search(text)
        if km and am:
            mapping[km.group(1)] = am.group(1)
    return mapping

def update_schet_addresses(schet_dir: Path, kpp_map: dict) -> int:
    if not schet_dir:
        return 0
    block_re = re.compile(r'(<Грузополучатель[\s\S]*?</Грузополучатель>)', re.DOTALL)
    kpp_re = re.compile(r'<СвЮЛ[^>]*?КПП="(\d{9})"')
    addr1_re = re.compile(r'(<Адрес[^>]*?АдрТекст=")([^"]*)(")')
    addr2_re = re.compile(r'(<АдрИно[^>]*?АдрТекст=")([^"]*)(")')
    changed = 0
    for path in sorted(schet_dir.glob("ON_SCHET__*.xml")):
        text = path.read_bytes().decode("windows-1251")
        block_m = block_re.search(text)
        if not block_m:
            continue
        block = block_m.group(1)
        kpp_m = kpp_re.search(block)
        if not kpp_m:
            continue
        new_addr = kpp_map.get(kpp_m.group(1))
        if not new_addr:
            continue
        new_block, n1 = addr1_re.subn(lambda m: m.group(1) + new_addr + m.group(3), block, count=1)
        new_block, n2 = addr2_re.subn(lambda m: m.group(1) + new_addr + m.group(3), new_block, count=1)
        if new_block != block:
            text = text.replace(block, new_block, 1)
            path.write_bytes(text.encode("windows-1251"))
            changed += 1 if (n1 or n2) else 0
    return changed

def pick_dirs(root: Path):
    upd_candidates = [root / "УПД" / "Отправляемые", root / "Отправляемые"]
    schet_candidates = [root / "Счет на оплату" / "Отправляемые", root / "Отправляемые"]
    upd_dir = next((p for p in upd_candidates if p.exists()), None)
    schet_dir = next((p for p in schet_candidates if p.exists()), None)
    return upd_dir, schet_dir

def main() -> None:
    base = Path(__file__).resolve().parent
    root = ask_root(base)

    removed = delete_konvert(root)

    upd_dir, schet_dir = pick_dirs(root)

    nsch_changed = replace_nsch_names(upd_dir)
    schet_changed = replace_schet_names(schet_dir)

    kpp_map = collect_kpp_to_addr(upd_dir)
    addr_changed = update_schet_addresses(schet_dir, kpp_map)

    print(f"Удалено konvert*: {removed}")
    print(f"Обновлено ON_NSCHFDOPPR: {nsch_changed}")
    print(f"Обновлено ON_SCHET (названия): {schet_changed}")
    print(f"Обновлено адресов Грузополучатель: {addr_changed}")

if __name__ == "__main__":
    main()
