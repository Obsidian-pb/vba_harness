VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_WriteTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_WriteTextFile - Tool: write text content to a file (UTF-8)
' Supports both complete overwrite of the file and appending to its end.
Option Explicit
Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "write_text_file"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Writes text content to a file in UTF-8 encoding. " & _
        "If the parameter ""append"" is false (default), the file is completely overwritten. " & _
        "If ""append"" is true, the content is added to the end of the file (if the file does not exist, it is created). " & _
        "Missing folders are created automatically. " & _
        "Use this tool when you need to save, write, create or append data to a text file."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """file_path"":{""type"":""string"",""description"":""Path to the text file (absolute or relative)""}," & _
            """content"":{""type"":""string"",""description"":""Text content to write to the file""}," & _
            """append"":{""type"":""boolean"",""description"":""If true, the content is appended to the end of the file; if false (default) the file is completely overwritten"",""default"":false}" & _
        "}," & _
        """required"":[""file_path"",""content""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim file_path As String
    Dim content As String
    Dim append As Boolean

    file_path = ExtractJsonString(argsJson, "file_path")
    content = ExtractJsonString(argsJson, "content")
    append = (StrComp(ExtractJsonString(argsJson, "append"), "true", vbTextCompare) = 0)
    
    If Trim$(file_path) = "" Then
        clsToolBase_Execute = MakeError("File path is not specified")
        Exit Function
    End If
    
    file_path = ResolveFilePath(file_path)
    
    Call WriteFileUTF8(file_path, content, append)
    
    Dim msg As String
    If append Then
        msg = "Data appended to file: " & file_path
    Else
        msg = "File successfully written: " & file_path
    End If
    clsToolBase_Execute = MakeOk(msg)
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = MakeError(Err.Description)
End Function



