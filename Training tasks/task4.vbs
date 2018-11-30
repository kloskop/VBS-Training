Option explicit
Dim oFSO,textStream, line_no, line, splitted, columnNumber, result, n, m, ColumnLength
Const path = "C:\Users\NEX2ZUU\Desktop\Zadania z VBS\Address.txt"

'Call FunctionA
Call FunctionB

Sub FunctionA
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set textStream = oFSO.OpenTextFile(path,1)
	
	columnNumber=CInt(InputBox("Enter column's number:"))
	result=""
	
	Do Until textStream.atEndOfStream
	line_no=0
		line_no=line_no+1
		line = textStream.Readline
		splitted = split(line,"|")
		result=result+splitted(columnNumber-1)&vbCrLf
	Loop
	
	WScript.Echo result
	Set oFSO = Nothing
End Sub

Sub FunctionB
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set textStream = oFSO.OpenTextFile(path,1)
	
	result =""
	Do Until textStream.atEndOfStream
	line_no=0
	line_no=line_no+1
	line=textStream.Readline
	splitted=split(line,"|")
	ColumnLength=0
	For n=0 to UBound(splitted)
		Select Case n 
			Case 0
				ColumnLength=10
			Case 2
				ColumnLength=32
			Case 3 
				ColumnLength=25
			Case 4
				ColumnLength=25
			Case 5
				ColumnLength=25
			Case 10
				ColumnLength=4
			Case else
				ColumnLength=3
		End Select
		
		
		For m=0 to (columnLength-len(splitted(n))-1)
			splitted(n)=splitted(n)+" "
		Next
		result=result+splitted(n)
		If n = UBound(splitted) then
		result=result&vbCrLf
		end if
	Next
	
	Loop
	WScript.Echo result
	
End Sub
