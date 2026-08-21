Attribute VB_Name = "m_FileIO"
' Module: m_FileIO - File I/O operations (read/write UTF-8, path resolution, folder creation)
Option Explicit

' ============================================================
'  File operations with UTF-8 encoding
' ============================================================

' Reads entire text file in UTF-8 encoding
Public Function ReadFileUTF8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                       ' adTypeText = 2
    stream.CharSet = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadFileUTF8 = stream.ReadText(-1)    ' adReadAll = -1
    stream.Close
    Set stream = Nothing
End Function

' Writes text to a file in UTF-8 encoding (without BOM).
' Overwrites existing file, creates missing folders.
Public Sub WriteFileUTF8(ByVal filePath As String, ByVal content As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folderPath As String
    folderPath = fso.GetParentFolderName(filePath)
    If folderPath <> "" Then
        EnsureFolderExists folderPath
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                       ' adTypeText = 2
    stream.CharSet = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2         ' adSaveCreateOverWrite = 2
    stream.Close
    Set stream = Nothing
    Set fso = Nothing
End Sub

' Resolves a path to absolute form. A relative path
' is resolved relative to the active Visio document folder.
Public Function ResolveFilePath(ByVal filePath As String) As String
    filePath = Trim$(filePath)
    
    ' Remove enclosing quotes
    If Left$(filePath, 1) = """" Then filePath = Mid$(filePath, 2)
    If Right$(filePath, 1) = """" Then filePath = Left$(filePath, Len(filePath) - 1)
    
    ' Check if the path is absolute (drive: or \\server\share)
    Dim isAbsolute As Boolean
    isAbsolute = (InStr(1, filePath, ":", vbTextCompare) > 0) Or (Left$(filePath, 2) = "\\")
    
    If Not isAbsolute Then
        Dim basePath As String
        basePath = ""
        On Error Resume Next
        basePath = ActiveDocument.path
        On Error GoTo 0
        If basePath = "" Then basePath = CurDir$
        
        If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
        filePath = basePath & filePath
    End If
    
    ResolveFilePath = filePath
End Function

' Recursively creates a folder with all nested subfolders.
Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        Set fso = Nothing
        Exit Sub
    End If
    
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If parent <> "" And Not fso.FolderExists(parent) Then
        EnsureFolderExists parent
    End If
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
End Sub

