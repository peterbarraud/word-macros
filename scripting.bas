Attribute VB_Name = "scripts"
Sub readFileAttributes()
    Const FILEPATH = "C:\Users\barraud\Documents\work\vba-macros\readme.md"
    Dim readFile As file
    Dim fileSystemObject As scripting.fileSystemObject
    Set fileSystemObject = New scripting.fileSystemObject
    
    Set readFile = fileSystemObject.GetFile(FILEPATH)
    Debug.Print readFile.Attributes
    If readFile.Attributes And ReadOnly Then
        Debug.Print "Is ReadOnly"
    ElseIf readFile.Attributes And Compressed Then
        Debug.Print "Is Compressed"
    'etc
    End If
    

End Sub

Sub readFile()
    Const FILEPATH = "C:\Users\barraud\Documents\work\vba-macros\readme.md"
    Dim readFile As TextStream
    Dim fileSystemObject As scripting.fileSystemObject
    Set fileSystemObject = New scripting.fileSystemObject
    Set readFile = fileSystemObject.OpenTextFile(FILEPATH, ForReading)
    Do Until readFile.AtEndOfStream
        Debug.Print readFile.ReadLine
    Loop
    
    readFile.Close

End Sub

Sub writeFile()
    Const FILEPATH = "C:\Users\barraud\Documents\work\vba-macros\readme.log"
    Dim writeFile As TextStream
    Dim fileSystemObject As scripting.fileSystemObject
    Set fileSystemObject = New scripting.fileSystemObject
    Set writeFile = fileSystemObject.CreateTextFile(FILEPATH, False)
    
    writeFile.WriteLine "line 1"
    writeFile.WriteLine "line 2"
    
    writeFile.Close
    

End Sub

Sub iterateFilesInFolder()
    Const FOLDERPATH = "C:\Users\barraud\Documents\work\vba-macros"
    
    Dim fileSystemObject As scripting.fileSystemObject
    Dim folderObject As Folder
    Dim fileInFolder As file
    
    Set fileSystemObject = New scripting.fileSystemObject
    
    Set folderObject = fileSystemObject.GetFolder(FOLDERPATH)
    
    For Each fileInFolder In folderObject.Files
        Debug.Print fileInFolder.Name
    Next

End Sub

