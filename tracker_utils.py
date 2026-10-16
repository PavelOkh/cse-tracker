"""Input and durable local result helpers (no network)."""
import csv
import re
from pathlib import Path

HEADERS = {'накладная', 'накладные', 'номер накладной', 'waybill', 'waybill number', 'tracking no', 'tracking number'}


def parse_numbers(raw):
    parts = re.split(r'[\s,;|]+', raw.strip()) if raw.strip() else []
    return unique_numbers(parts)


def unique_numbers(items):
    result = []
    seen = set()
    for item in items:
        if item is None:
            continue
        value = str(item).strip().strip('`')
        if not value or value.lower() in HEADERS or re.fullmatch(r'[-:]+',value):
            continue
        if value not in seen:
            result.append(value)
            seen.add(value)
    return result


def excel_safe(value):
    if isinstance(value,str) and value.startswith(('=','+','-','@')):
        return "'" + value
    return value


def save_csv(rows, path):
    """Atomic CSV checkpoint; a failed update preserves the previous checkpoint."""
    path = Path(path)
    path.parent.mkdir(parents=True,exist_ok=True)
    columns = list(dict.fromkeys(k for row in rows for k in row))
    temporary = path.with_suffix(path.suffix+'.tmp')
    try:
        with temporary.open('w',encoding='utf-8-sig',newline='') as f:
            writer = csv.DictWriter(f,fieldnames=columns,delimiter=';')
            writer.writeheader()
            writer.writerows({k:excel_safe(v) for k,v in row.items()} for row in rows)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
