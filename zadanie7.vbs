option explicit
'First task
'Dim x
'For x = 1 To 10 Step -2
'MsgBox x,0,"Result"
'Next

Dim number, a, b
b = 0
number = CInt(InputBox("Enter a number between 1 and 25"))
If number < 1 and number > 25 Then
MsgBox "Your number is outside the range"
Else 
For a=1 To number
b = b+1
MsgBox "Number: "&b
Next
End If

	
	