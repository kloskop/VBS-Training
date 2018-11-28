Dim oFSO, sourceFolder, destFolder

Set oFSO=createobject("Scripting.FileSystemObject")

'Copying folder
'sourceFolder="C:\Users\NEX2ZUU\Desktop\Materialy\VBS Course\Dir1"
'destFolder="C:\Users\NEX2ZUU\Desktop\Materialy\VBS Course\Dir1-copy"

'oFSO.CopyFolder sourceFolder, destFolder, True
strFolder="C:\Users\NEX2ZUU\Desktop\Materialy\VBS Course\Dir1"

'Deleting folder
If oFSO.FolderExists(strFolder) Then
	oFSO.DeleteFolder(strFolder)
	MsgBox strFolder &" - is deleted now."
Else
	MsgBox strFolder &" - folder doesn't exist",0,"Alert!"
End If

'To move folder use .MoveFolder [source],[destination]

Set oFSO = Nothing

