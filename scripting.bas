Attribute VB_Name = "Module2"
Sub UseFileSystemObject()
    
    Const FOLDERPATH = ""
    Dim fileSystemObject As Scripting.fileSystemObject
    Dim folderObject As Folder
    Dim file As file
    
    Dim selectedFile As file
    
    Set folderObject = fileSystemObject.GetFolder(FOLDERPATH)
    
    selectedFile = fileSystemObject.GetFile(FILEPATH)
    
    For Each file In folderObject.Files
        Debug.Print file.Name
        ' File attributes
        ' Check if the file is read-only
        If file.Attributes And ReadOnly Then
            Debug.Print "File is Read-only"
        End If
    Next
    
    
End Sub

