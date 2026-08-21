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
'---Enable error handling - to prevent class disposal when trying to access nothing
    On Error GoTo Tail
    
    LLM_chat_html.Show

    
Tail:
    '---Exit the handler procedure
End Sub

Private Sub ComButLLMSettings_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'---Enable error handling - to prevent class disposal when trying to access nothing
    On Error GoTo Tail
    
    LLM_config.Show
    
    
Tail:
    '---Exit the handler procedure
End Sub
