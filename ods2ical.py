#!/usr/bin/python3
# -*- coding: utf-8 -*-

import argparse
from datetime import timedelta
from icalendar import Calendar, Event, vText
from dateutil.parser import parse as dateparse
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from zoneinfo import ZoneInfo
import os


def get_text_from_element(elem):
    """Extracts all text recursively from an ODF element."""
    text_parts = []
    for node in elem.childNodes:
        if hasattr(node, "data"):  # text node
            text_parts.append(node.data)
        elif hasattr(node, "childNodes"):  # nested elements (like <span>)
            text_parts.append(get_text_from_element(node))
    return "".join(text_parts)


def read_ods(file_path, max_events=200):
    """Reads the first sheet of an ODS file (A–E: Datum, Uhrzeit, Typ, Thema, Alter)."""
    print(f"Reading ODS file: {file_path}")
    doc = load(file_path)
    rows = []
    sheet = doc.spreadsheet.getElementsByType(Table)[0]  # only first sheet

    for i, row in enumerate(sheet.getElementsByType(TableRow)):
        cells = []
        for cell in row.getElementsByType(TableCell):
            ps = cell.getElementsByType(P)
            cell_text = " ".join(get_text_from_element(p).strip() for p in ps).strip()
            cells.append(cell_text)
        if not any(cells):
            continue
        if len(cells) > 5:
            cells = cells[:5]
        rows.append(cells)
        if len(rows) >= max_events:
            break
    return rows


def read_events_from_rows(rows):
    """Parses rows (A–E) into event dicts."""
    results = []
    for i, row in enumerate(rows):
        if i == 0 and "Datum" in row[0]:
            continue
        if len(row) < 2 or not row[0].strip():
            continue

        datetimetxt = f"{row[0].strip()} {row[1].strip()}"
        time_str = row[1].strip().lower()
        is_open_end = time_str.lower().startswith("ab") if time_str else False
        has_time = bool(time_str)

        try:
            cleaned = datetimetxt.lower().replace("uhr", "").replace("ab", "").strip()
            start = dateparse(cleaned, dayfirst=True)
            if has_time:
                # Add timezone info for time-based events
                start = start.replace(tzinfo=ZoneInfo("Europe/Berlin"))
            end = None if is_open_end else (start + timedelta(hours=1) if has_time else None)
        except Exception as e:
            print(f"⚠️ Date parsing error in row {i}: '{datetimetxt}' -> {e}")
            continue

        category = row[2].strip() if len(row) > 2 else ''
        title = row[3].strip() if len(row) > 3 else 'Unbenannte Veranstaltung'
        age = row[4].strip() if len(row) > 4 else ''

        results.append({
            'begin': start,
            'end': end,
            'category': category,
            'title': title,
            'age': age,
            'open_end': is_open_end,
            'has_time': has_time,
        })
    return results


def generate_ical(events):
    """Generates a clean iCalendar file with timezone-aware times and proper categories."""
    cal = Calendar()
    cal.add('prodid', '-//suhler-sternfreunde.de//')
    cal.add('version', '2.0')

    for e in events:
        event = Event()
        try:
            if not e['has_time']:
                # All-day event
                event.add('dtstart', e['begin'].date())
            else:
                # Timed event with timezone
                event.add('dtstart', e['begin'])
                if not e['open_end'] and e['end']:
                    event.add('dtend', e['end'])

            event.add('summary', e['title'])
            event.add('location', vText("Sternwarte Suhl"))

            categories = []
            if e['category']:
                categories.append(vText(e['category']))
            if e['age']:
                categories.append(vText(f"ab {e['age']} Jahre"))

            if categories:
                # Add all as separate category values
                event.add('categories', categories)

            cal.add_component(event)
        except Exception as err:
            print(f"⚠️ Fehler beim Hinzufügen des Events '{e['title']}': {err}")
    return cal


def save_ical(cal, filename):
    """Saves iCalendar object to file."""
    content = cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"✅ iCal-Datei gespeichert: {filename}")


def main():
    parser = argparse.ArgumentParser(
        description='Convert an ODS working plan (A–E columns) to a timezone-correct iCalendar (.ics) file.',
        epilog='Example: ./ods2ical.py "Arbeitsplan 2026.ods"'
    )
    parser.add_argument('ods_file', help='Pfad zur ODS-Datei (Arbeitsplan)')
    parser.add_argument('--max', type=int, default=200, help='Maximale Anzahl an Terminen')
    args = parser.parse_args()

    ext = os.path.splitext(args.ods_file)[1].lower()
    if ext != '.ods':
        print("⚠️ Bitte eine ODS-Datei angeben (z. B. Arbeitsplan 2026.ods)")
        return

    rows = read_ods(args.ods_file, max_events=args.max)
    events = read_events_from_rows(rows)
    cal = generate_ical(events)
    ical_path = os.path.splitext(args.ods_file)[0] + ".ics"
    save_ical(cal, ical_path)


if __name__ == '__main__':
    main()
