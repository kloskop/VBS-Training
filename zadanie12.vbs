option explicit

Dim AdamDate, FlightHour
AdamDate = Date
FlightHour=TimeSerial(06,30,45)
MsgBox "Adam's take off date: "&AdamDate&" - "&FlightHour
MsgBox "Adam's return date: "&DateAdd("d",10,AdamDate)