VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_GetProjecCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_GetProjecCode - Tool: get VBA module/project code by name
Option Explicit
Implements clsToolBase

' The tool provides the model with information about the VBA program code of the entire project or a specific module with an explicitly specified name

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "get_module_vba_code"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "The tool provides the model with information about the VBA program code of the entire project or a specific module with an explicitly specified name. " & _
            "Call it when you need to view the entire project as a whole or the program code of one specific module."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """project_name"":{""type"":""string"",""description"":""If you need to refer to a specific VBA project by name, you can specify it here. If not specified, the current project will be used.""}," & _
            """module_name"":{""type"":""string"",""description"":""If you need to find out the VBA code of a specific module, you can specify it here.""}" & _
        "}," & _
        """required"":[""project_name""]" & _
    "}"
End Property


Private Function clsToolBase_Execute(ByVal argsJson As String) As String
Dim project_name As String
Dim module_name As String
Dim code As String

    On Error GoTo ErrHandler
    
    project_name = ExtractJsonString(argsJson, "project_name")
    module_name = ExtractJsonString(argsJson, "module_name")
    
    If project_name = "" Then
        ' If the project name is not specified, assume the current project
        project_name = Application.Vbe.ActiveVBProject.name
    End If
    
    If project_name = "" Then
        ' The active project could not be determined
        clsToolBase_Execute = MakeError("The VBA project name has not been specified.")
        Exit Function
    End If

    ' Find the project by name
    If module_name = "" Then
        code = GetAllProjectCode(project_name)
        clsToolBase_Execute = MakeOk("Full VBA project code is: " & Chr(10) & Chr(10) & code)
    Else
        code = GetVbaCodeProj(project_name, module_name)
        clsToolBase_Execute = MakeOk("Module '" & module_name & "' VBA code is: " & Chr(10) & Chr(10) & code)
    End If
    
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = MakeError(Err.Description)
End Function



Function GetAllProjectCode(projectName As String) As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim result As String
    Dim separator As String
    
    separator = String(60, "=") & vbCrLf
    
    Set vbProj = GetVBProjectByName(projectName)
    
    If vbProj Is Nothing Then
        result = "Project not found"
        GetAllProjectCode = result
        Exit Function
    End If
    
    result = "Project: " & vbProj.name & vbCrLf & separator
    
    For Each vbComp In vbProj.VBComponents
        ' Skip components without code (e.g., empty forms)
        If vbComp.CodeModule.CountOfLines > 0 Then
            result = result & vbCrLf
            result = result & separator
            result = result & "' MODULE: " & vbComp.name & _
                     "  [type: " & ComponentTypeName(vbComp.Type) & "]" & vbCrLf
            result = result & separator
            result = result & GetVbaCode(vbComp.name) & vbCrLf
        End If
    Next
    
    GetAllProjectCode = result
End Function

Private Function ComponentTypeName(compType As Long) As String
    Select Case compType
        Case 1:   ComponentTypeName = "Module"      ' vbext_ct_StdModule
        Case 2:   ComponentTypeName = "Class"       ' vbext_ct_ClassModule
        Case 3:   ComponentTypeName = "Form"        ' vbext_ct_MSForm
        Case 100: ComponentTypeName = "Document"    ' vbext_ct_Document
        Case Else: ComponentTypeName = "Type=" & compType
    End Select
End Function

