VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Module: ThisDocument - Document-level event handlers (commented out by default)
Option Explicit


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    If ThisDocument.VBProject.name = HARNESS_NAME Then
        AddTB_LLM
    End If
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    If ThisDocument.VBProject.name = HARNESS_NAME Then
        RemoveTB_LLM
    End If
End Sub
