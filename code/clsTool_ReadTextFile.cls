VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReadTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_ReadTextFile - Tool: read a text file and return its content
Option Explicit
Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "read_text_file"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Reads a text file in UTF-8 encoding and returns its entire content. " & _
        "Use this tool when the user asks to read, open or show the contents of a text file. " & _
        "The file path can be absolute or relative (relative to the current Visio document folder)."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """file_path"":{""type"":""string"",""description"":""Path to the text file (absolute or relative)""}" & _
        "}," & _
        """required"":[""file_path""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim file_path As String
    file_path = ExtractJsonString(argsJson, "file_path")
    
    If Trim$(file_path) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString("File path is not specified") & "}"
        Exit Function
    End If
    
    file_path = ResolveFilePath(file_path)
    
    If Dir$(file_path) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString("File not found: " & file_path) & "}"
        Exit Function
    End If
    
    Dim content As String
    content = ReadFileUTF8(file_path)
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(content) & "}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function


