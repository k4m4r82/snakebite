Attribute VB_Name = "modFs"
'Returns a boolean - True if the file exists
Public Function FExists(OrigFile As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    FExists = fs.fileexists(OrigFile)
End Function
'Returns a boolean - True if the folder exists
Public Function DirExists(OrigFile As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    DirExists = fs.folderexists(OrigFile)
End Function

