'Cheking if number is a prime one
option explicit

dim input, n, isPrime
n=2
input = CDbl(InputBox("Enter a number: ","Cheking if Prime","Your number"))
'input=179426549 'test prime input
isPrime=true
Do while n<sqr(input) 'condition speeding up calculations
	if input mod n = 0 then
		isPrime=false
		exit do
	end if
	n=n+1
Loop
If isPrime then
	WScript.Echo input&" is Prime"
Else
	WScript.Echo input&" is not Prime"
End if

	