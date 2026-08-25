Attribute VB_Name = "m_code"
' Module: m_code - Functions for accessing VBA project code (get module/selected code, project search)
Option Explicit

' ==============================================================================
'  MODULE FOR ACCESSING VBA PROJECT CODE
' ==============================================================================
'  Functions for getting module code, selected text, and project search.
'  Refactored: eliminated duplication, replaced On Error with checks.
' ==============================================================================

' --- Unified function for getting module code ---
' If projectName = "" — the active project is used.
' If moduleName = "" — the code of the active module is returned.
Public Function GetModuleCode(Optional ByVal moduleName As String = "", _
                              Optional ByVal projectName As String = "") As String
    Dim cm As Object
    Set cm = GetCodeModule(moduleName, projectName)
    If cm Is Nothing Then Exit Function
    If cm.CountOfLines > 0 Then
        GetModuleCode = cm.Lines(1, cm.CountOfLines)
    End If
End Function

' --- Get CodeModule by module and project name ---
Private Function GetCodeModule(ByVal moduleName As String, _
                               ByVal projectName As String) As Object
    Dim vbProj As Object
    Dim vbComp As Object
    
    ' Determine the project
    If projectName = "" Then
        Set vbProj = Application.Vbe.ActiveVBProject
    Else
        Set vbProj = GetVBProjectByName(projectName)
        If vbProj Is Nothing Then
            Err.Raise 9, , "Project '" & projectName & "' not found"
        End If
    End If
    
    ' If module name is not specified — return the active CodePane
    If moduleName = "" Then
        On Error Resume Next
        Set GetCodeModule = Application.Vbe.ActiveCodePane.CodeModule
        On Error GoTo 0
        Exit Function
    End If
    
    ' Search for the component by name
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(moduleName)
    On Error GoTo 0
    If vbComp Is Nothing Then
        Err.Raise 9, , "Module '" & moduleName & "' does not exist" & _
                      IIf(projectName <> "", " in project '" & projectName & "'", "")
    End If
    Set GetCodeModule = vbComp.CodeModule
End Function

' --- Get selected code (or current line if no selection) ---
Public Function GetSelectedCode() As String
    Dim cp As Object
    Dim cm As Object
    Dim startLine As Long, startCol As Long
    Dim endLine As Long, endCol As Long
    
    On Error Resume Next
    Set cp = Application.Vbe.ActiveCodePane
    On Error GoTo 0
    If cp Is Nothing Then Exit Function
    
    Set cm = cp.CodeModule
    cp.GetSelection startLine, startCol, endLine, endCol
    
    If startLine = endLine And startCol = endCol Then
        ' Cursor without selection — return the entire module
        If cm.CountOfLines > 0 Then
            GetSelectedCode = cm.Lines(1, cm.CountOfLines)
        End If
    Else
        ' Selected range
        GetSelectedCode = cm.Lines(startLine, endLine - startLine + 1)
    End If
End Function

' --- Find project by name (case-insensitive) ---
Public Function GetVBProjectByName(ByVal projectName As String) As Object
    Dim vbProj As Object
    Dim i As Long
    
    For i = 1 To Application.Vbe.VBProjects.Count
        Set vbProj = Application.Vbe.VBProjects(i)
        If StrComp(vbProj.name, projectName, vbTextCompare) = 0 Then
            Set GetVBProjectByName = vbProj
            Exit Function
        End If
    Next i
    ' Not found — return Nothing
End Function

' ==============================================================================
'  COMPATIBILITY WITH OLD CODE (wrappers)
' ==============================================================================
Public Function GetVbaCode(Optional moduleName As String = "") As String
    GetVbaCode = GetModuleCode(moduleName)
End Function

Public Function GetVbaCodeProj(ByVal project_name, Optional moduleName As String = "") As String
    GetVbaCodeProj = GetModuleCode(moduleName, project_name)
End Function

