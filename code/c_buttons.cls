VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module: c_buttons - Class for handling toolbar button click events
Option Explicit
'Class for handling button clicks

Public WithEvents ComButOpenChat As Office.CommandBarButton          'LLM Chat
Attribute ComButOpenChat.VB_VarHelpID = -1
Public WithEvents ComButLLMSettings As Office.CommandBarButton       'LLM Settings
Attribute ComButLLMSettings.VB_VarHelpID = -1



Private Sub Class_Initialize()
    Set ComButOpenChat = Application.Vbe.CommandBars(TOLLBAR_NAME).Controls("Chat")
    Set ComButLLMSettings = Application.Vbe.CommandBars(TOLLBAR_NAME).Controls("Settings")
End Sub

Private Sub Class_Terminate()
    Set ComButOpenChat = Nothing
    Set ComButLLMSettings = Nothing
End Sub



Private Sub ComButOpenChat_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
' ERR-02: log Err.Number/Err.Description and notify the user instead of
' silently swallowing the error when jumping to the tail label.
    On Error GoTo ErrHandler

    LLM_chat_html.Show
    
    Exit Sub

ErrHandler:
    Log "[c_buttons.ComButOpenChat_Click] Error " & Err.Number & ": " & Err.Description
    MsgBox "Could not open Chat." & vbCrLf & _
           "Error: " & Err.Number & " — " & Err.Description, _
           vbExclamation + vbOKOnly, "LLM Harness"
    On Error GoTo 0
End Sub

Private Sub ComButLLMSettings_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
' ERR-02: log Err.Number/Err.Description and notify the user instead of
' silently swallowing the error when jumping to the tail label.
    On Error GoTo ErrHandler

    LLM_config.Show
    
    Exit Sub

ErrHandler:
    Log "[c_buttons.ComButLLMSettings_Click] Error " & Err.Number & ": " & Err.Description
    MsgBox "Could not open Settings." & vbCrLf & _
           "Error: " & Err.Number & " — " & Err.Description, _
           vbExclamation + vbOKOnly, "LLM Harness"
    On Error GoTo 0
End Sub
