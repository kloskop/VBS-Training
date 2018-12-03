option explicit

Dim oFSO, oFile, line_no, line, result, splitted, i, isEmpty
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile("C:\Users\NEX2ZUU\Desktop\Zadania z VBS\TestPK1.csv",1)


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
isEmpty = 0
line_no=0
	line_no=line_no+1
	line = oFile.ReadLine
	splitted = split(line,",")
		for i=0 to UBound(splitted)-1
			If splitted(i)="" then
				isEmpty = isEmpty+1
			else
			end if
		next
	if isEmpty<>UBound(splitted) then
		result=result+replace(line,",","|")&vbCrlf
	else
	end if
	'WScript.Echo isEmpty&"-"&UBound(splitted)
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
