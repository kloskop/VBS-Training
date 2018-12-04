option explicit

Dim oFSO, oFile, line_no, line, result, splitted, i, isEmpty
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile("C:\Users\NEX2ZUU\Desktop\Zadania z VBS\AddressChangeReport.csv",1)


Function splitLine(line) 

Dim regex
Set regex = CreateObject("vbscript.regexp")

regex.IgnoreCase = True
regex.Global = True

'Pattern = ",(?=([^"]*"[^"]*")*(?![^"]*"))"
regex.Pattern = ",(?=([^" & Chr(34) & "]*" & Chr(34) & "[^" & Chr(34) & "]*" & Chr(34) & ")*(?![^" & Chr(34) & "]*" & Chr(34) & "))"
    'regex.replaces will replace the commas outside quotes with semicolons and then the
    'Split function will split the result based on the semicollons
splitLine = Split(regex.Replace(line, ";"), ";")

End Function



result=""
Do until oFile.atEndOfStream
'isEmpty = 0
line_no=0
	line_no=line_no+1
	line = oFile.ReadLine
	splitted = splitLine(line)
		for i=0 to UBound(splitted)-1
			result=result&"|"&splitted(i)
		next
	result=result&vbCrLf
Loop

WScript.Echo result

Set oFSO = Nothing

'Do Until oFile.atEndOfStream
'line_no=0
'	line_no=line_no+1
'	line=oFile.ReadLine
'	splitted=splitLine(line)
'		for i=0 to UBound(splitted)
'		result=result&" "&splitted(i)
'		next
'Loop
'WScript.Echo result
oFile.Close