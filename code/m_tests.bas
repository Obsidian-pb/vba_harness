Attribute VB_Name = "m_tests"
' diffirent tests

Public Sub TestSaveProfile()
' Save (or update) a profile
SaveLLMProfile "smart", "https://api.selectel.ru/aig/v1/chat/completions", _
               "sk-sl-v1-***", "deepseek/deepseek-v4-pro-0813", "Пиши все комментарии к коду VBA на английском"
End Sub


Public Sub TestLoadProfile()
' Get a profile object
Dim p As Object
    Set p = GetLLMProfile("default")
    If Not p Is Nothing Then Debug.Print p("model")
End Sub


Public Sub TEstGetProfileFields()
' Get a profile field by field
    Dim u As String, k As String, m As String, s As String
    If GetLLMProfileParams("default", u, k, m, s) Then Debug.Print m
End Sub


Public Sub TestProfilesList()
Dim names As Collection
Dim i As Long

    Set names = GetProfileNames()
    
    If names.Count = 0 Then
        MsgBox "No profiles have been saved yet."
    Else
        For i = 1 To names.Count
            Debug.Print i & ": " & names(i)
        Next i
End If

End Sub
