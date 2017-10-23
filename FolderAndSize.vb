Sub Example2()
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("\\wp131\3DGenerationTeam")
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
'print folder name
Cells(i + 1, 1) = objSubFolder.Name
'print folder path
Cells(i + 1, 2) = objSubFolder.Path
'print folder size dont / 1073741824 if you want result in bytes
Cells(i + 1, 3) = objSubFolder.Size / 1073741824
'GB or byte depending on the above
Cells(i + 1, 4) = "GB"
i = i + 1
Next objSubFolder
End Sub