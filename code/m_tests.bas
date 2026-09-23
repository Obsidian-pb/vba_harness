Attribute VB_Name = "m_tests"
' diffirent tests

Public Sub TestSaveProfile()
' Save (or update) a profile
SaveLLMProfile "default", "https://api.selectel.ru/aig/v1/chat/completions", _
               "sk-sl-v1-***", "deepseek/deepseek-v4-flash-0731", _
               "80000", "0.8", "Пиши все комментарии к коду VBA на английском"
End Sub


Public Sub TestLoadProfile()
' Get a profile object
Dim p As Object
    Set p = GetLLMProfile("default")
    If Not p Is Nothing Then Debug.Print p("model")
End Sub


Public Sub TEstGetProfileFields()
' Get a profile field by field
    Dim u As String, k As String, m As String, mt As String, t As String, s As String
    If GetLLMProfileParams("smart", u, k, m, mt, t, s) Then
        Debug.Print u
        Debug.Print k
        Debug.Print m
        Debug.Print mt
        Debug.Print t
        Debug.Print s
    End If
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
