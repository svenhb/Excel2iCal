# Excel2iCal
Exports excel data of birthdays and further information to iCal calender format for calendar import  
Events will be set as 'all day' with yearly repitition  
  
Exportiert Exceldaten von Geburtstagen und weitere Angaben ins iCal Format für den Import in Kalender  
Die Termine werden als ganztägig  mit jährlicher Wiederholung gesetzt  
  
  
### Screenshot of Excel2iCal.xlsm  
![Screenshot](Excel2iCal.png?raw=true "Screenshot") 
  
### Screenshot of one event after import into gmx  
![Screenshot](Excel2iCal_import2.png?raw=true "Screenshot") 
![Screenshot](Excel2iCal_import.png?raw=true "Screenshot") 
  
  
### Output result.ics  
```
BEGIN:VCALENDAR
CALSCALE:GREGORIAN
BEGIN:VEVENT
DTSTART;VALUE=DATE:19010120
DTEND;VALUE=DATE:19010120
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Opa 1 (Meier)
DESCRIPTION:Partner: Oma 1 (Meier)\nKinder: Anna, Beate, Carla\n
LOCATION:Friedhof 1
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19750718
DTEND;VALUE=DATE:19750718
RRULE:FREQ=YEARLY
SUMMARY:Todestag Opa 1 (Meier)
DESCRIPTION:Partner: Oma 1 (Meier)\nKinder: Anna, Beate, Carla\n
LOCATION:Friedhof 1
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19050219
DTEND;VALUE=DATE:19050219
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Oma 1 (Meier)
DESCRIPTION:Partner: Opa 1 (Meier)\nKinder: Anna, Beate, Carla\n
LOCATION:Friedhof 1
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19801013
DTEND;VALUE=DATE:19801013
RRULE:FREQ=YEARLY
SUMMARY:Todestag Oma 1 (Meier)
DESCRIPTION:Partner: Opa 1 (Meier)\nKinder: Anna, Beate, Carla\n
LOCATION:Friedhof 1
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19400313
DTEND;VALUE=DATE:19400313
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Anna Mustermann
DESCRIPTION:Partner: Max Mustermann\nEltern: Oma 1, Opa 1 Meier\nKinder: Ernst, Frida\n
LOCATION:
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:20100414
DTEND;VALUE=DATE:20100414
RRULE:FREQ=YEARLY
SUMMARY:Todestag Anna Mustermann
DESCRIPTION:Partner: Max Mustermann\nEltern: Oma 1, Opa 1 Meier\nKinder: Ernst, Frida\n
LOCATION:
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19420717
DTEND;VALUE=DATE:19420717
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Max Mustermann
DESCRIPTION:Partner: Anna Mustermann, geborene as Meier\nEltern: Oma 2, Opa 2 Mustermann\nKinder: Ernst, Frida\n
LOCATION:
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19760615
DTEND;VALUE=DATE:19760615
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Ernst Mustermann
DESCRIPTION:Eltern: Anna, Max Mustermann\n
LOCATION:
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19720716
DTEND;VALUE=DATE:19720716
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Frieda Neumann
DESCRIPTION:Partner: Gert Neumann\nEltern: Anna, Max Mustermann\nKinder: Heiko, Ida\n
LOCATION:
END:VEVENT
BEGIN:VEVENT
DTSTART;VALUE=DATE:19710808
DTEND;VALUE=DATE:19710808
RRULE:FREQ=YEARLY
SUMMARY:Geburtstag Gert Neumann
DESCRIPTION:Partner: Frida Neumann, geborene Mustermann\n
LOCATION:
END:VEVENT
END:VCALENDAR
```
