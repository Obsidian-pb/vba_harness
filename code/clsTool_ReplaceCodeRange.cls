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
'  Example: range [100:101] (2 lines), new code - 10 lines.
'  Result: lines 100-101 removed, 10 new lines inserted.
'  Final = [code before 100] + [10 new lines] + [code after 101].
'
'  OUT-OF-RANGE HANDLING:
'    - By default (strict mode) an end_line that exceeds the total number
'      of module lines is treated as an error: the module is left unchanged
'      and an error response is returned. This makes the caller aware of
'      an incorrect range instead of silently replacing a larger range.
'    - Automatic reduction of end_line to the last line is performed ONLY
'      when the explicit mode "clamp_range":true is passed.
'
'  ============================================================
'  UNDO MODEL: SINGLE LAST CHANGE (not a multi-step history)
'  ============================================================
'  This tool keeps ONLY ONE undoable operation - the most recent
'  replacement performed by THIS instance of the tool.
'
'    - The undo slot is a single record: project_name, module_name,
'      the full saved module code, the replaced range and the new code.
'    - A new replacement COMPLETELY OVERWRITES the previous undo state.
'      After two replacements in a row, only the last one can be undone;
'      the first one is lost forever.
'    - ClearUndo() discards the saved state without touching the module.
'    - UndoLast() can be called only once per saved state: after a
'      SUCCESSFUL rollback the undo slot is cleared.
'    - CanUndo reports whether an undoable state is present.
'    - UndoDescription returns a text description of the saved state.
'
'  ERROR HANDLING IN UndoLast (HI-03):
'    - The undo state is cleared ONLY after a fully successful restoration.
'    - If UndoLast fails for any reason (project/module not found, runtime
'      error during delete/insert), the saved data is KEPT in the undo slot,
'      m_undo_available remains True, and the diagnostic message is exposed
'      through the UndoLastError property. The caller can then retry UndoLast
'      or inspect the reason of the failure without losing the snapshot.
'
'  IMPORTANT LIMITATION (external modifications):
'    Undo restores the saved module content as-is. The stored snapshot is
'    taken right before the replacement, so Undo CANNOT detect whether the
'    module was changed by another means (other tools, manual edits, another
'    instance of this tool) AFTER the replacement was made. In that case
'    UndoLast() will OVERWRITE those external modifications with the saved
'    snapshot. Treat Undo as "restore the exact previous snapshot", not as
'    a merge with the current module state.
' ============================================================

' ========== Undo data (single slot for the last change) ==========
Private m_undo_available   As Boolean ' True if the single undo slot holds a saved state
Private m_undo_projectName As String  ' Project of the saved snapshot
Private m_undo_moduleName  As String  ' Module of the saved snapshot
Private m_undo_fullCode    As String  ' Full module code BEFORE the replacement
Private m_undo_startLine   As Long    ' Replaced range start
Private m_undo_endLine     As Long    ' Replaced range end
Private m_undo_newCode     As String  ' Code inserted by the last replacement
Private m_undoLastError    As String  ' Diagnostic message of the last failed UndoLast attempt

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "replace_code_range"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = _
        "Tool that replaces VBA code in the specified line range of a given module. " & _
        "Lines in the specified range are completely removed, and new code is inserted in their place. " & _
        "Parameters: 'module_name' - name of the module (required); " & _
        "'start_line' - start of the line range, integer >= 1 (required); " & _
        "'end_line' - end of the line range, integer >= start_line (required); " & _
        "'new_code' - new code to insert (required, can be empty string for simple deletion); " & _
        "'project_name' - VBA project name (if not specified, the active project is used). " & _
        "'clamp_range' - optional boolean mode; if end_line exceeds the module bounds this mode " & _
        "automatically reduces end_line to the last line (default: an error is returned instead). " & _
        "IMPORTANT: the resulting module size changes by the difference between new code lines and the removed range size. " & _
        "SUPPORT FOR EMPTY MODULES: if the module is empty (0 lines), start_line and end_line must be equal to 1 - " & _
        "in this case new code will be inserted into the empty module. " & _
        "UNDO MODEL: single last change - this tool keeps ONLY ONE undoable operation " & _
        "(the most recent replacement performed by this instance); a new replacement overwrites " & _
        "the previous undo state. Undo restores the exact snapshot saved before the replacement " & _
        "and does not detect or merge external modifications made by other tools or manual edits. " & _
        "On a failed UndoLast the saved undo state is KEPT and the failure reason is available " & _
        "via the UndoLastError property."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """project_name"":{""type"":""string"",""description"":""VBA project name. If not specified, the current active project is used.""}," & _
            """module_name"":{""type"":""string"",""description"":""Name of the module to replace code in (required).""}," & _
            """start_line"":{""type"":""integer"",""description"":""Start of line range (inclusive), >= 1. For empty module must be 1.""}," & _
            """end_line"":{""type"":""integer"",""description"":""End of line range (inclusive), >= start_line. For empty module must be 1. Must not exceed the module size unless clamp_range is true.""}," & _
            """new_code"":{""type"":""string"",""description"":""New code to insert instead of the removed range. Can be empty string for deletion without insertion.""}," & _
            """clamp_range"":{""type"":""boolean"",""description"":""Optional. If true, end_line that exceeds the module boundaries is automatically reduced to the last line. Default false: an error is returned instead.""}" & _
        "}," & _
        """required"":[""project_name"",""module_name"",""start_line"",""end_line"",""new_code""]" & _
    "}"
End Property


' ============================================================
'  UndoScope: public methods for undo/check
'  Model: SINGLE LAST CHANGE (one undo slot per instance)
' ============================================================

' True only if this instance has a saved state from its last replacement.
Public Property Get CanUndo() As Boolean
    CanUndo = m_undo_available
End Property

' Human-readable description of the saved undo state (the last change).
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

' Diagnostic message of the LAST FAILED UndoLast attempt.
' Returns an empty string if the last UndoLast call succeeded (or none was made).
Public Property Get UndoLastError() As String
    UndoLastError = m_undoLastError
End Property

' Undo the last saved change of this instance.
' IMPORTANT:
'   - Restores the exact snapshot saved before the replacement.
'   - Does NOT detect whether the module was changed by external means after the change.
'   - On SUCCESS the single undo slot is cleared (m_undo_available = False).
'   - On FAILURE the undo state is PRESERVED: the saved data stays available,
'     the description is exposed via UndoLastError, and the caller may retry.
Public Function UndoLast() As Boolean
    m_undoLastError = ""

    If Not m_undo_available Then
        m_undoLastError = "No undo operations available"
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
        m_undoLastError = "Undo failed: VBA project '" & m_undo_projectName & "' was not found. Saved undo state is preserved."
        UndoLast = False
        Exit Function
    End If

    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(m_undo_moduleName)
    On Error GoTo UndoErrHandler

    If vbComp Is Nothing Then
        m_undoLastError = "Undo failed: module '" & m_undo_moduleName & _
                          "' was not found in project '" & m_undo_projectName & _
                          "'. Saved undo state is preserved."
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

    ' Full restoration succeeded: the single undo slot is now consumed.
    m_undo_available = False
    m_undoLastError = ""
    UndoLast = True
    Exit Function

UndoErrHandler:
    ' HI-03: on any error DURING restoration we must NOT discard the undo
    ' data. The module may be in an intermediate state, so the snapshot is
    ' kept (m_undo_available stays True) and the diagnostic message is stored.
    m_undoLastError = "Undo failed: " & Err.Description & _
                      " (Error " & Err.Number & "). Saved undo state is preserved."
    UndoLast = False
End Function

' Clears the single undo slot without modifying the module.
Public Sub ClearUndo()
    m_undo_available = False
    m_undo_projectName = ""
    m_undo_moduleName = ""
    m_undo_fullCode = ""
    m_undo_startLine = 0
    m_undo_endLine = 0
    m_undo_newCode = ""
    m_undoLastError = ""
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
    
    ' Numeric parameters - parse
    Dim startStr As String
    Dim endStr As String
    startStr = ExtractJsonString(argsJson, "start_line")
    endStr = ExtractJsonString(argsJson, "end_line")
    
    ' Optional mode: clamp_range - allow automatic reduction of end_line
    ' to the last line when it exceeds the module boundaries.
    ' By default (false) an out-of-range end_line is an error.
    Dim clamp_range As Boolean
    Dim clampStr As String
    clampStr = LCase$(Trim$(ExtractJsonString(argsJson, "clamp_range")))
    clamp_range = (clampStr = "true" Or clampStr = "1" Or clampStr = "yes")

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

        ' Save the only undo slot (previous state overwritten)
        m_undo_available = True
        m_undoLastError = ""
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

    ' Out-of-range end_line handling:
    ' - By default (clamp_range=false) this is strictly an error: the module
    '   is NOT changed and the caller gets an explicit error message.
    ' - Automatic reduction to the last line is performed ONLY in the
    '   explicitly requested clamp_range mode.
    If end_line > totalLines Then
        If Not clamp_range Then
            clsToolBase_Execute = MakeError("end_line (" & endStr & ") exceeds module boundaries (total lines: " & totalLines & "). " & _
                                            "Use ""clamp_range"":true to allow automatic reduction to the last line.")
            Exit Function
        End If
        end_line = totalLines
    End If

    ' Get full module code (for Undo)
    Dim fullCode As String
    fullCode = cm.Lines(1, totalLines)

    ' ============================================================
    '  SAVE STATE FOR UNDO (single slot - replaces previous state)
    ' ============================================================
    m_undo_available = True
    m_undoLastError = ""
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

    ' First delete the selected range
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
               "Module size changed by " & (newLineCount - rangeSize) & " line(s). " & _
               "Undo (single last change) is available via UndoLast; on failure " & _
               "the saved state is preserved and UndoLastError returns the reason."

    clsToolBase_Execute = MakeOk(resultMsg)
    Exit Function

ErrHandler:
    clsToolBase_Execute = MakeError("Error: " & Err.Description)
End Function

' ============================================================
'  HELPER FUNCTIONS
' ============================================================

' Count lines in text
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

