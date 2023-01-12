#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from datetime import datetime,timedelta
from icalendar import Calendar, Event, vText


def read_the_sheet(sheet="",max=50):
    results=[]
    for row in range(1,50):
      if sheet[row][0]:
          datetimetxt = sheet[row][0] + " " + sheet[row][1]
          try:
            eventbegin = datetime.strptime(datetimetxt,"%d/%m/%y %H:%M Uhr")
            eventend = eventbegin + timedelta(hours=1)
          except:
            eventbegin = datetimetxt
            eventend = datetimetxt  
          location = sheet[row][2]
          title = sheet[row][3]
          age = sheet[row][4]
          results.append({'begin': eventbegin, 'end': eventend, 'location': location,'title': title, 'age': age})
          if len(results)==max:
            break
    return results


def display(cal):
   return cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()
   #return cal.to_ical().decode("utf-8").strip()



import csv

if __name__=="__main__":
 
    csv_input="Tabelle.csv"
    with open(csv_input, 'r', newline='') as csv_file:
        csv_content = csv.reader(csv_file, delimiter=';')
        csv_content = list(csv_content)
    print(csv_content)
    csv_events = read_the_sheet(csv_content,max=3)
    cal = Calendar()
    cal.add('prodid', '-//suhler-sternfreunde.de//')
    cal.add('version', '2.0')
    for csv_event in csv_events:
      print(csv_event['begin'])
      event = Event()
      try:
        event.add('dtstart',csv_event['begin'])
        event.add('summary',csv_event['title'])
        event.add('dtend',csv_event['end'])
        event.add('location',csv_event['location'])
        event.add('description',"")
        categories=[]
        if 'Planetarium' in csv_event['location']:
          categories.append(vText("Planetariumsvorführung"))
        if csv_event['age']:
          categories.append(vText("ab {} Jahre".format(csv_event['age'])))
        if categories:
          #print(categories)
          #event.set_inline('categories',categories,encode=0)
          #event['categories']=vText(",".join(categories))
          # kategorien funktionieren nicht beim import in plone
          pass  
        cal.add_component(event)
      except Exception as error:
        print(error)
    print(display(cal))
    #print(cal)
    
    # write to file
    ical_output = "Tabelle.ics"
    with open(ical_output, 'w') as ics_file:
        ics_file.write(display(cal))
