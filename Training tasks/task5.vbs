Option explicit

Dim intRow, intColumn, cell, objExcel, objWorkbook, objSheet,result,n
intColumn=1
result=""

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\NEX2ZUU\Desktop\Zadania z VBS\test.xls")
Set objSheet = objWorkbook.sheets(1)

For intRow=1 to objSheet.usedrange.rows.count 'counting rows containing data - CAREFUL WHEN EMPTY CELLS APPEARS BETWEEN OTHERS
Do Until objExcel.Cells(intRow,intColumn).Value="" 'Iteration through rows 
	cell = objExcel.Cells(intRow,intColumn).Value
	intColumn = intColumn+1
	result=result+CStr(cell)
	For n=1 to (10-len(CStr(cell)))
		result=result+" "
	Next
	
Loop
result=result&vbCrLf 
intColumn=1

Next
WScript.Echo result
objExcel.Quit

