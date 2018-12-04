Option explicit
Dim oFSO,textStream, line_no, line, splitted, columnNumber, result, n, m, ColumnLength, ColumnData
Const path = "C:\Users\NEX2ZUU\Desktop\Zadania z VBS\Address.txt"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set textStream = oFSO.OpenTextFile(path,1)

'Call FunctionA
Call FunctionB

Sub FunctionA
	
	
	columnNumber=CInt(InputBox("Enter column's number:"))
	result=""
	
	'Iterating through txt file line by line
	Do Until textStream.atEndOfStream 
	line_no=0
		line_no=line_no+1
		line = textStream.Readline
		splitted = split(line,"|")
		result=result+splitted(columnNumber-1)&vbCrLf
	Loop
	
	WScript.Echo result
	
End Sub

Sub FunctionB

	result =""
	'Iterating through txt file line by line
	Do Until textStream.atEndOfStream
	line_no=0
	line_no=line_no+1
	line=textStream.Readline
	splitted=split(line,"|")
	ColumnLength=0
	'Adjusting data display to fixed column lengths
	For n=0 to UBound(splitted)
		ColumnData = trim(splitted(n))
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
		'Adding space signs to ColumnData
		For m=0 to (columnLength-len(ColumnData)-1)
			ColumnData=ColumnData+" "
		Next
		result=result+ColumnData
		'Merging result into single stream
		If n = UBound(splitted) then
			result=result&vbCrLf
		end if
	Next
	
	Loop
	WScript.Echo result
	
	
End Sub

Set oFSO = Nothing
textStream.Close