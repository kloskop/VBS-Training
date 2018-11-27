' Declare variables
Dim input1, input2, total
Const  SITE_TITLE = "www.GlobalETraining.com" 

'Getting the input from the user
input1 = InputBox("Enter the first number: ") 
input2 = InputBox("Enter the second number: ") 

total  = Add(CInt(input1),CInt(input2))

MsgBox "The sum of the two numbers : " & total, 0, SITE_TITLE

'A Function procedure -- can return a value. 
Function Add(num1, num2)
    sum = num1 + num2
Add = sum
End Function

