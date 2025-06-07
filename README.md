# excel2Ical

import an excel csv-sheet and save as ical file

## Motivation

Add events to a plone website is painfully, if you do that manually event by event. Let's say we have an excel sheet with all the events of the current year. What if we could export that excel 
sheet as csv file, convert it to an ical file and import all the events in 10 sec to our plone site.

Fortunately, plone has a feature "ical import".

## Requirements

pip3 install icalendar
pip3 install python-dateutil

## Usage

Convert a CSV file with event data to an iCalendar (.ics) file.

positional arguments:
  csv_file    Path to the CSV input file
  delimiter   Delimiter used in the CSV file (e.g. ";")

options:
  -h, --help  show this help message and exit
  --max MAX   Maximum number of events to read

Example: ./convertCsv2ical.py sommer.2023.csv ";"


