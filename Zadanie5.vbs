johnAge = CInt(InputBox("Enter John's Age ")) 
kerryAge = CInt(InputBox("Enter Kerry's Age ")) 

If johnAge > kerryAge Then
MsgBox "John is older than Kerry"
Elseif johnAge < kerryAge Then
MsgBox "Kerry is older than John"
Else
MsgBox "Kerry and John are at the same age"
End If
