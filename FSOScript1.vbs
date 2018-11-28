Option explicit

Dim oFSO
'Dim oDrive, cDrives
Dim drive
'Dim ListOfDrives
Dim diskType

drive="C:\"

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

'Set oDrive = oFSO.GetDrive("C:")
'Viewing drive's size
'MsgBox "Available space: "&FormatNumber(((oDrive.AvailableSpace/1024)/1024)/1024,0)&" GB",0, "C Drive"
'MsgBox "Free space: "&FormatNumber(((oDrive.FreeSpace/1024)/1024)/1024,0)&" GB",0, "C Drive"
'MsgBox "Total Size: "&FormatNumber(((oDrive.TotalSize/1024)/1024)/1024,0)&" GB",0, "C Drive"

'Viewing drive's type
'Select Case oDrive.DriveType
'	Case 0: disktype = "Unknown"
'	Case 1: disktype = "Removable"
'	Case 2: disktype = "Fixed"
'	Case 3: disktype = "Network"
'	Case 4: disktype = "CD-ROM"
'	Case 5: disktype = "RAM Disk"
'End Select

'MsgBox "Drive Type: "& oDrive.DriveType,0,"Drive Type Number Returned by the drive object:"
'MsgBox "Drive Type: "& disktype,0,"Drive Type: (Use Friendly Strings)"
'MsgBox "FileSystem Type: "& oDrive.FileSystem,0,"Drive Information:"
'MsgBox "Volume Name: "& oDrive.VolumeName,0,"Drive Information: "

'Viewing Serial Number and isReady
'MsgBox "Serial Number: "& oDrive.SerialNumber,0,"Serial Number ("& oDrive.DriveLetter&": ):"
'MsgBox "Is the drive ready: "& oDrive.IsReady,0,"Drive Information: "

'Checking if the disk exists
If oFSO.DriveExists(drive)="True" Then
	MsgBox "We found the drive "&drive,0, "Result"
Else
	MsgBox "We did not find the drive "&drive,0, "Result"
End If




