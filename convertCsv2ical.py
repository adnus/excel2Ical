#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import argparse
from datetime import datetime, timedelta
from icalendar import Calendar, Event, vText
from dateutil.parser import parse as dateparse


def read_events_from_csv(sheet, max_events=50):
    results = []
    for i, row in enumerate(sheet):
        if not row or len(row) < 5 or not row[0].strip():
            print(f"Skipping empty or invalid row {i}: {row}")
            continue

        datetimetxt = f"{row[0].strip()} {row[1].strip()}"
        try:
            cleaned = datetimetxt.replace("Uhr", "").strip()
            start = dateparse(cleaned, dayfirst=True)
            end = start + timedelta(hours=1)
        except Exception as e:
            print(f"Date parsing error in row {i}: '{datetimetxt}' -> {e}")
            continue

        location = row[2].strip() if len(row) > 2 else ''
        title = row[3].strip() if len(row) > 3 else 'Untitled Event'
        age = row[4].strip() if len(row) > 4 else ''

        results.append({
            'begin': start,
            'end': end,
            'location': location,
            'title': title,
            'age': age
        })

        if len(results) >= max_events:
            break
    return results


def generate_ical(events):
    """Converts event dictionaries into an iCalendar file content."""
    cal = Calendar()
    cal.add('prodid', '-//suhler-sternfreunde.de//')
    cal.add('version', '2.0')

    for event_data in events:
        event = Event()
        try:
            event.add('dtstart', event_data['begin'])
            event.add('dtend', event_data['end'])
            event.add('summary', event_data['title'])
            event.add('location', event_data['location'])

            categories = []
            if 'Planetarium' in event_data['location']:
                categories.append("Planetariumsvorführung")
            if event_data['age']:
                categories.append(f"ab {event_data['age']} Jahre")

            if categories:
                # Note: "categories" is replaced with "description" for Plone compatibility
                event.add('description', ", ".join(categories))

            cal.add_component(event)
        except Exception as e:
            print(f"Error adding event: {e}")
    return cal


def save_ical(cal, filename):
    """Saves iCalendar object to file."""
    content = cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"iCal file saved to {filename}")


def main():
    parser = argparse.ArgumentParser(
        description='Convert a CSV file with event data to an iCalendar (.ics) file.',
        epilog='Example: ./convertCsv2ical.py sommer.2023.csv ";"'
    )
    parser.add_argument('csv_file', help='Path to the CSV input file')
    parser.add_argument('delimiter', help='Delimiter used in the CSV file (e.g. ";")')
    parser.add_argument('--max', type=int, default=50, help='Maximum number of events to read')
    args = parser.parse_args()

    with open(args.csv_file, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=args.delimiter)
        csv_data = list(reader)

    events = read_events_from_csv(csv_data, max_events=args.max)
    cal = generate_ical(events)
    save_ical(cal, args.csv_file + ".ics")


if __name__ == '__main__':
    main()
