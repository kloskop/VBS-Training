Option explicit
dim n, input


function Factorial(num)
Factorial=1
	For n=1 to num
		Factorial=n*Factorial(n-1)
	Next
end function

input = CDbl(InputBox("Enter a number to get it's factorial value: ","Factorial","Your Number"))
WScript.Echo "Result: "&(Factorial(input))

'zmienic na for -1