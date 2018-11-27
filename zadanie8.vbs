option explicit
dim number, a, result
number = CInt(InputBox("Enter the number between 1 and 100: "))
If number < 1 and number > 100 Then
MsgBox "The number "&number&" is outside the range"
Else
For a=1 To number 
If Not result = a mod 10 = 0 Then
MsgBox "Number: "&a
End If
Next
End If