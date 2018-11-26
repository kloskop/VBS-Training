option explicit

Dim name, age, city
Const title = "Basic information"
name = InputBox("Enter your name: ", title, "your name")
age = InputBox("Enter your age: ", title, "your age")
city = InputBox("Enter your city: ", title, "your city")

MsgBox name & " is " & age & " years old and lives in " & city, 0, title

