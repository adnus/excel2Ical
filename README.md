# excel2Ical

import an excel csv-sheet and save as ical file

## Motivation

Add events to a plone website is painfully, if you do that manually event by event. Let's say we have an excel sheet with all the events of the current year. What if we could export that excel 
sheet as csv file, convert it to an ical file and import all the events in 10 sec to our plone site.

Fortunately, plone has a feature "ical import".

## Requirements

pip3 install icalendar

## Usage

In this initial version, you have to code some the parameters hard in the source.

The time2string format has to be define in the code, depending of the cell
content (cell 1 and cell 2).


