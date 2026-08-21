VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReplaceCodeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_ReplaceCodeRange - Tool: replace VBA code in a specified line range of a module
Option Explicit
Option Base 0
Implements clsToolBase

' ============================================================
'  Tool: replace_code_range
'  Replaces VBA code in the specified line range of a given module.
'  Lines in the specified range are completely removed,
'  and new code is inserted in their place.
'  Resulting module size = [old size] + [new code lines] - [range size].
'
'  SUPPORT FOR EMPTY MODULES:
'    - If the module is empty (CountOfLines = 0), then start_line and end_line
'      must be equal to 1. In this case the new code is inserted
'      into the empty module via AddFromString.
'
'  Example: range [100:101] (2 lines), new code — 10 lines.
'  Result: lines 100-101 removed, 10 new lines inserted.
'  Final = [code before 100] + [10 new lines] + [code after 101].
'
'  UNDO SUPPORT (UndoScope):
'    - Before executing the replacement, the full module state is saved.
'    - After a successful replacement, UndoLast() can be called to revert.
'    - CanUndo shows whether undo is available.
'    - UndoDescription returns a text description of the last replacement.
' ============================================================

' ========== Undo data ==========
Private m_undo_available   As Boolean
Private m_undo_projectName As String
Private m_undo_moduleName  As String
Private m_undo_fullCode    As String
Private m_undo_startLine   As Long
Private m_undo_endLine     As Long
Private m_undo_newCode     As String

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "replace_code_range"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = _
        "Tool that replaces VBA code in the specified line range of a given module. " & _
        "Lines in the specified range are completely removed, and new code is inserted in their place. " & _
        "Parameters: 'module_name' — name of the module (required); " & _
        "'start_line' — start of the line range, integer >= 1 (required); " & _
        "'end_line' — end of the line range, integer >= start_line (required); " & _
        "'new_code' — new code to insert (required, can be empty string for simple deletion); " & _
        "'project_name' — VBA project name (if not specified, the active project is used). " & _
        "IMPORTANT: the resulting module size changes by the difference between new code lines and the removed range size. " & _
        "SUPPORT FOR EMPTY MODULES: if the module is empty (0 lines), start_line and end_line must be equal to 1 — " & _
        "in this case new code will be inserted into the empty module."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """project_name"":{""type"":""string"",""description"":""VBA project name. If not specified, the current active project is used.""}," & _
            """module_name"":{""type"":""string"",""description"":""Name of the module to replace code in (required).""}," & _
            """start_line"":{""type"":""integer"",""description"":""Start of line range (inclusive), >= 1. For empty module must be 1.""}," & _
            """end_line"":{""type"":""integer"",""description"":""End of line range (inclusive), >= start_line. For empty module must be 1.""}," & _
            """new_code"":{""type"":""string"",""description"":""New code to insert instead of the removed range. Can be empty string for deletion without insertion.""}" & _
        "}," & _
        """required"":[""project_name"",""module_name"",""start_line"",""end_line"",""new_code""]" & _
    "}"
End Property


' ============================================================
'  UndoScope: public methods for undo/check
' ============================================================

Public Property Get CanUndo() As Boolean
    CanUndo = m_undo_available
End Property

Public Property Get UndoDescription() As String
    If Not m_undo_available Then
        UndoDescription = "No undo operations available"
    Else
        UndoDescription = "Replaced range in module '" & m_undo_moduleName & _
                         "' of project '" & m_undo_projectName & _
                         "' (lines " & m_undo_startLine & "-" & m_undo_endLine & _
                         ", new code: " & CountLines(m_undo_newCode) & " lines)"
    End If
End Property

Public Function UndoLast() As Boolean
    If Not m_undo_available Then
        UndoLast = False
        Exit Function
    End If

    On Error GoTo UndoErrHandler

    Dim vbProj As Object
    If m_undo_projectName = "" Then
        Set vbProj = Application.Vbe.ActiveVBProject
    Else
        Set vbProj = GetVBProjectByName(m_undo_projectName)
    End If

    If vbProj Is Nothing Then
        UndoLast = False
        Exit Function
    End If

    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(m_undo_moduleName)
    On Error GoTo UndoErrHandler

    If vbComp Is Nothing Then
        UndoLast = False
        Exit Function
    End If

    Dim cm As Object
    Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then
        cm.DeleteLines 1, cm.CountOfLines
    End If

    If m_undo_fullCode <> "" Then
        cm.AddFromString m_undo_fullCode
    End If

    m_undo_available = False
    UndoLast = True
    Exit Function

UndoErrHandler:
    m_undo_available = False
    UndoLast = False
End Function

Public Sub ClearUndo()
    m_undo_available = False
    m_undo_projectName = ""
    m_undo_moduleName = ""
    m_undo_fullCode = ""
    m_undo_startLine = 0
    m_undo_endLine = 0
    m_undo_newCode = ""
End Sub

' ============================================================
'  Execute
' ============================================================
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    Dim project_name As String
    Dim module_name  As String
    Dim start_line   As Long
    Dim end_line     As Long
    Dim new_code     As String

    On Error GoTo ErrHandler

    ' Extract parameters from JSON
    project_name = ExtractJsonString(argsJson, "project_name")
    module_name = ExtractJsonString(argsJson, "module_name")
    new_code = ExtractJsonString(argsJson, "new_code")

    ' Numeric parameters — parse
    Dim startStr As String
    Dim endStr As String
    startStr = ExtractJsonString(argsJson, "start_line")
    endStr = ExtractJsonString(argsJson, "end_line")
    

    ' Check required fields
    If module_name = "" Then
        clsToolBase_Execute = MakeError("Module name (module_name) is not specified")
        Exit Function
    End If

    If startStr = "" Then
        clsToolBase_Execute = MakeError("Start line (start_line) is not specified")
        Exit Function
    End If

    If endStr = "" Then
        clsToolBase_Execute = MakeError("End line (end_line) is not specified")
        Exit Function
    End If

    ' Convert to numbers
    On Error Resume Next
    start_line = CLng(startStr)
    end_line = CLng(endStr)
    On Error GoTo ErrHandler

    If start_line < 1 Then
        clsToolBase_Execute = MakeError("start_line must be >= 1, got: " & startStr)
        Exit Function
    End If

    If end_line < start_line Then
        clsToolBase_Execute = MakeError("end_line (" & endStr & ") must be >= start_line (" & startStr & ")")
        Exit Function
    End If

    ' Determine the project
    Dim vbProj As Object
    If project_name = "" Then
        Set vbProj = Application.Vbe.ActiveVBProject
    Else
        Set vbProj = GetVBProjectByName(project_name)
    End If
    
    If vbProj Is Nothing Then
        clsToolBase_Execute = MakeError("Project not found")
        Exit Function
    End If

    ' Check if the target project is the current harness project
    If vbProj.name = HARNESS_NAME Then
        clsToolBase_Execute = MakeError("Cannot make changes to project '" & vbProj.name & "'")
        Exit Function
    End If

    ' Find the module
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(module_name)
    On Error GoTo 0

    If vbComp Is Nothing Then
        clsToolBase_Execute = MakeError("Module '" & module_name & "' not found in project")
        Exit Function
    End If

    Dim cm As Object
    Set cm = vbComp.CodeModule

    Dim totalLines As Long
    totalLines = cm.CountOfLines

    ' ============================================================
    '  BRANCH: EMPTY MODULE (totalLines = 0)
    ' ============================================================
    If totalLines = 0 Then
        ' For empty module, start_line and end_line must be 1
        If start_line <> 1 Or end_line <> 1 Then
            clsToolBase_Execute = MakeError("Module '" & module_name & "' is empty. To insert code into an empty module, start_line and end_line must both be 1")
            Exit Function
        End If

        ' Save state for Undo
        m_undo_available = True
        If project_name = "" Then
            m_undo_projectName = vbProj.name
        Else
            m_undo_projectName = project_name
        End If
        m_undo_moduleName = module_name
        m_undo_fullCode = ""   ' module was empty
        m_undo_startLine = start_line
        m_undo_endLine = end_line
        m_undo_newCode = new_code

        ' Insert new code into empty module
        If new_code <> "" Then
            cm.AddFromString new_code
        End If

        ' Build response
        Dim newLineCount_empty As Long
        newLineCount_empty = CountLines(new_code)
        If new_code = "" Then newLineCount_empty = 0

        Dim resultMsg_empty As String
        resultMsg_empty = "Successfully inserted " & newLineCount_empty & " line(s) of new code into empty module '" & module_name & "'."

        clsToolBase_Execute = MakeOk(resultMsg_empty)
        Exit Function
    End If

    ' ============================================================
    '  NON-EMPTY MODULE: standard logic
    ' ============================================================

    ' Check that the range is within module boundaries
    If start_line > totalLines Then
        clsToolBase_Execute = MakeError("start_line (" & startStr & ") exceeds module boundaries (total lines: " & totalLines & ")")
        Exit Function
    End If

    ' Adjust end_line if it exceeds boundaries
    If end_line > totalLines Then
        end_line = totalLines
    End If

    ' Get full module code (for Undo)
    Dim fullCode As String
    fullCode = cm.Lines(1, totalLines)

    ' ============================================================
    '  SAVE STATE FOR UNDO
    ' ============================================================
    m_undo_available = True
    If project_name = "" Then
        m_undo_projectName = vbProj.name
    Else
        m_undo_projectName = project_name
    End If
    m_undo_moduleName = module_name
    m_undo_fullCode = fullCode
    m_undo_startLine = start_line
    m_undo_endLine = end_line
    m_undo_newCode = new_code

    ' ============================================================
    '  EXECUTE THE REPLACEMENT
    ' ============================================================
    Dim rangeSize As Long
    rangeSize = end_line - start_line + 1

    ' First delete the old range
    On Error Resume Next
    cm.DeleteLines start_line, rangeSize
    If Err.Number <> 0 Then
        Dim errDel As Long: errDel = Err.Number
        Dim errDescDel As String: errDescDel = Err.Description
        On Error GoTo ErrHandler
        Err.Raise errDel, , "Error deleting lines " & start_line & "-" & end_line & ": " & errDescDel
    End If
    On Error GoTo ErrHandler

    ' Insert new code at the same position
    If new_code <> "" Then
        cm.InsertLines start_line, new_code
    End If

    ' ============================================================
    '  BUILD RESPONSE
    ' ============================================================
    Dim newLineCount As Long
    newLineCount = CountLines(new_code)
    If new_code = "" Then newLineCount = 0

    Dim resultMsg As String
    resultMsg = "Range [" & start_line & ":" & end_line & "] (" & rangeSize & " line(s)) " & _
               "in module '" & module_name & "' successfully replaced. " & _
               "Inserted " & newLineCount & " line(s) of new code. " & _
               "Module size changed by " & (newLineCount - rangeSize) & " line(s)."

    clsToolBase_Execute = MakeOk(resultMsg)
    Exit Function

ErrHandler:
    clsToolBase_Execute = MakeError("Error: " & Err.Description)
End Function

' ============================================================
'  HELPER FUNCTIONS
' ============================================================

' Count lines in text (aligns with CountLines logic in clsTool_ReplaceCode)
Private Function CountLines(ByVal text As String) As Long
    If text = "" Then
        CountLines = 0
        Exit Function
    End If

    Dim normalized As String
    normalized = Replace(text, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim i As Long
    Dim lineCount As Long
    lineCount = 1
    For i = 1 To Len(normalized)
        If Mid$(normalized, i, 1) = vbLf Then
            lineCount = lineCount + 1
        End If
    Next i

    CountLines = lineCount
End Function

Private Function MakeError(ByVal msg As String) As String
    MakeError = "{""status"":""error"",""message"":""" & EscapeJsonString(msg) & """}"
End Function

Private Function MakeOk(ByVal msg As String) As String
    MakeOk = "{""status"":""ok"",""message"":""" & EscapeJsonString(msg) & """}"
End Function

' Minimal escaping: quotes and backslashes
Private Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function
