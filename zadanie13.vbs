option explicit
'First task
' Dim number, grade1, grade2, grade3
Dim input1, input2, input3
'number = 9.333333333333 
'MsgBox FormatNumber(number,3)
'Second task
'grade1 = 90/100
'grade2 = 95/100
'grade3 = 96/100
'MsgBox FormatPercent(((grade1+grade2+grade3)/3),3)
'Third task
input1 = CInt(InputBox("Enter Math grade: ","MATH","Grade in %"))
input2 = CInt(InputBox("Enter Eng grade: ","ENGLISH","Grade in %"))
input3 = CInt(InputBox("Enter Science grade: ","SCIENCE","Grade in %"))
MsgBox "Average of your grades is: "&FormatPercent(((input1/100+input2/100+input3/100)/3),2)
