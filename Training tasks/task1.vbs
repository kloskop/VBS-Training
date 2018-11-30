'displaying numbers from 1 to 100 with divisibility checking

option explicit

dim n, result
result =""

For n=1 to 100
	If n mod 3 = 0 Then
		result=result+"Pop"
	End If
	If n mod 5 = 0 Then
		result=result+"Star"
	End If
	If n mod 5<>0 and n mod 3<>0 Then
		result=result+CStr(n)
	End If
	If n<>100 Then
	result=result+","
	End If
Next
WScript.Echo result