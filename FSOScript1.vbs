Option explicit

Dim oFSO, oDrive, cDrives
Dim ListOfDrives
Set oFSO = CreateObject("Scripting.FileSystemObject")
' Creating folder oFSO.CreateFolder("C:\Users\NEX2ZUU\Desktop\Materialy\VBS Course\Dir1")

'Viewing list of drives
'Set cDrives = oFSO.Drives
'For Each oDrive in cDrives
'	MsgBox "Drive letter: "& oDrive.DriveLetter,0,"Drive on Your Computer: "
'	ListOfDrives = ListOfDrives&" "& oDrive.DriveLetter
'Next

'MsgBox "Number of Drives: "& cDrives.Count,0,"Drives On Your Computer:"
'MsgBox "Drive letters: "& ListOfDrives,0,"Drives On Your Computer:"

'Viewing drive's size
Set oDrive = oFSO.GetDrive("C:")
MsgBox "Available space: "&FormatNumber(((oDrive.AvailableSpace/1024)/1024)/1024,0)&" GB",0, "C Drive"
MsgBox "Free space: "&FormatNumber(((oDrive.FreeSpace/1024)/1024)/1024,0)&" GB",0, "C Drive"
MsgBox "Total Size: "&FormatNumber(((oDrive.TotalSize/1024)/1024)/1024,0)&" GB",0, "C Drive"