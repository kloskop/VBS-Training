
option explicit

Dim name 
name = InputBox("Enter your Last Name: ")



Select Case name
  Case "Thompson"
    MsgBox "Hello "&name&", you are assigned as a President"
  Case "Rooter"
    MsgBox "Hello "&name&", you are assigned as a Sr. Vice President"
  Case "Cooper"
    MsgBox "Hello "&name&", you are assigned as a Vice President"
  Case "Parker"
    MsgBox "Hello "&name&", you are assigned as a Manager"
  Case else
    MsgBox "Hello "&name&", you are not part of the Managment Team"
End Select
