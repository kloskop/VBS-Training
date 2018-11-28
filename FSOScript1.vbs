Option explicit

Dim oFSO, oDrive, cDrives
Dim ListOfDrives
Set oFSO = CreateObject("Scripting.FileSystemObject")
' Creating folder oFSO.CreateFolder("C:\Users\NEX2ZUU\Desktop\Materialy\VBS Course\Dir1")

Set cDrives = oFSO.Drives
For Each oDrive in cDrives
	MsgBox "Drive letter: "& oDrive.Drive.DriveLetter,0,"Drive on Your Computer: "
	ListOfDrives = ListOfDrives&" "& oDrive.DriveLetter
Next

MsgBox "Number of Drives: "& oDrives.Count,0,"Drives On Your Computer:"
MsgBox "Drive letters: "& ListOfDrives,0,"Drives On Your Computer:"
