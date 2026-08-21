VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_GetSelectedCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_GetSelectedCode - Tool: get selected code from active VBA editor window
Option Explicit
Implements clsToolBase

' ============================================================
'  Tool: get_selected_code
'  Returns the selected VBA program code from the active editor window.
'  If no code is selected (just cursor), returns the full code
'  of the active module/form/class entirely.
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "get_selected_code"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = _
        "The tool returns the selected VBA program code from the active editor window. " & _
        "If no code is selected (just cursor), it returns the full code of the active module/form/class entirely. " & _
        "Parameters are not required."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{""type"":""object"",""properties"":{},""required"":[]}"
End Property


Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    Dim cm As Object
    Dim cp As Object
    Dim startLine As Long, startCol As Long
    Dim endLine As Long, endCol As Long
    Dim lineCount As Long
    Dim resultCode As String
    
    On Error GoTo ErrHandler
    
    ' Get the active code window
'    Application.Vbe.MainWindow.SetFocus
    On Error Resume Next
    Set cp = Application.Vbe.ActiveCodePane
'    LLM_chat_html.Show
    On Error GoTo 0
    
    If cp Is Nothing Then
        clsToolBase_Execute = "{""status"":""error"",""message"":""No active editor window""}"
        Exit Function
    End If
    
    Set cm = cp.CodeModule
    
    ' Get selection coordinates
    cp.GetSelection startLine, startCol, endLine, endCol
    
    ' If nothing is selected (just cursor) — get the entire code of the active module
    If startLine = endLine And startCol = endCol Then
        If cm.CountOfLines > 0 Then
            resultCode = cm.Lines(1, cm.CountOfLines)
        Else
            resultCode = ""
        End If
    Else
        ' If a range is selected — extract it
        lineCount = endLine - startLine + 1
        resultCode = cm.Lines(startLine, lineCount)
    End If
    
    ' Build the response
    Dim responseMsg As String
    responseMsg = "Code from module '" & cm.parent.name & "':" & vbCrLf & vbCrLf & resultCode
       
    clsToolBase_Execute = "{""status"":""ok"",""message"":""" & responseMsg & """}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""Error: " & Err.Description & """}"
End Function

