option explicit

Dim input1, input2, inputOperator, result
inputOperator = InputBox("Choose the math operation: add|sub|mtp|div|pow")
input1 = CInt(InputBox("Enter the first number: ")) 
input2 = CInt(InputBox("Enter the second number: "))


Function Add(num1,num2)
	Add = num1+num2
End Function

Function Substract(num1,num2)
	Substract=num1-num2
End Function

Function Multiple(num1,num2)
	Multiple=num1*num2
End Function

Function Division(num1,num2)
	Division=num1/num2
End Function

Function Power(num1,num2)
	Power=num1^num2
End Function

Sub Message(StrMsg, result)
	MsgBox "Your "&StrMsg&" result is "&result,0,"Result"
End Sub

Select Case inputOperator
Case "add"
	Call Message(inputOperator,Add(input1,input2))
Case "sub"
	Call Message(inputOperator,Substract(input1,input2))
Case "mtp"
	Call Message(inputOperator,Multiple(input1,input2))
Case "div"
	Call Message(inputOperator,Division(input1,input2))
Case "pow"
	Call Message(inputOperator,Power(input1,input2))
Case Else
	MsgBox "You've entered wrong operation's name"
	WScript.Quit 
End Select

