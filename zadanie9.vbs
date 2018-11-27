option explicit

Dim myArray, a, i
myArray = Array("Text0","Text1","Text2","Text3","Text4","Text5","Text6","Text7","Text8","Text9")
a=1
For  i=LBound(myArray) To UBound(myArray)
MsgBox "Element no. "&a&" is: "&myArray(a-1)
a=a+1
Next
