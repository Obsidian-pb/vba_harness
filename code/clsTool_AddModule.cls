VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_AddModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_AddModule - Tool: add a new VBA component (module, class, or form) to a project
Option Explicit
Implements clsToolBase

' ============================================================
'  Tool: add_module
'  Adds a new component (module, class, form) to a VBA project
' ============================================================

' VBA component types
Private Const vbext_ct_StdModule  As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm     As Long = 3

' ========== Undo data ==========
Private m_undo_available As Boolean ' True if there is something to roll back
Private m_undo_vbProj As Object ' The project where the component was created
Private m_undo_compName As String ' Name of the created component
Private m_undo_compType As Long ' Type of the created component
Private m_undo_code As String ' The code that was inserted (if any)



Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "add_module"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = _
        "The tool creates a new component (module, class or form) in a VBA project. " & _
        "Parameters: 'project_name' - project name (if not specified, the active one is used); " & _
        "'module_name' - name of the new component (required); " & _
        "'module_type' - component type: 'module' (standard module), 'class' (class) or 'form' (user form); " & _
        "'code' - optional initial code to insert into the new module."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """project_name"":{""type"":""string"",""description"":""VBA project name. If not specified, the current active project is used.""}," & _
            """module_name"":{""type"":""string"",""description"":""Name of the component to create (required).""}," & _
            """module_type"":{""type"":""string"",""description"":""Component type: module / class / form""}," & _
            """code"":{""type"":""string"",""description"":""Initial program code to insert into the new module (optional).""}" & _
        "}," & _
        """required"":[""project_name"",""module_name""]" & _
    "}"
End Property

Public Property Get clsToolBase_ChatString() As String
    clsToolBase_ChatString = "Adding a new module"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    Dim project_name As String
    Dim module_name  As String
    Dim module_type  As String
    Dim code         As String

    On Error GoTo ErrHandler

    ' Extract parameters from JSON
    project_name = ExtractJsonString(argsJson, "project_name")
    module_name = ExtractJsonString(argsJson, "module_name")
    module_type = ExtractJsonString(argsJson, "module_type")
    code = ExtractJsonString(argsJson, "code")

    ' Check required field
    If module_name = "" Then
        clsToolBase_Execute = MakeError("Module name (module_name) not specified")
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
    
    ' Check if the modified project is the current Harness project
    If vbProj.name = HARNESS_NAME Then
        clsToolBase_Execute = MakeError("Cannot make changes to project '" & vbProj.name & "'")
        Exit Function
    End If

    ' Check if a component with this name already exists
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents(module_name)
    On Error GoTo 0
    If Not existingComp Is Nothing Then
        clsToolBase_Execute = MakeError("Component with name '" & module_name & "' already exists in the project")
        Exit Function
    End If

    ' Determine the component type
    Dim compType As Long
    Select Case LCase$(Trim$(module_type))
        Case "class", "classmodule", "cls"
            compType = vbext_ct_ClassModule
        Case "form", "userform", "user_form"
            compType = vbext_ct_MSForm
        Case Else   ' Default - standard module
            compType = vbext_ct_StdModule
    End Select

    ' Create a new component
    Dim newComp As Object
    Set newComp = vbProj.VBComponents.Add(compType)
    newComp.name = module_name

    ' If code is provided - insert it into the module
    If code <> "" Then
        Dim cm As Object
        Set cm = newComp.CodeModule
        cm.AddFromString code
    End If

    ' Build success response
    Dim typeName As String
    Select Case compType
        Case vbext_ct_StdModule:   typeName = "Module"
        Case vbext_ct_ClassModule: typeName = "Class"
        Case vbext_ct_MSForm:      typeName = "Form"
    End Select

    Dim resultMsg As String
    resultMsg = "Component '" & module_name & "' (type: " & typeName & ") successfully created in project '" & vbProj.name & "'."
    If code <> "" Then
        resultMsg = resultMsg & " Code inserted (" & Len(code) & " characters)."
    End If

    clsToolBase_Execute = MakeOk(resultMsg)
    Exit Function

ErrHandler:
    clsToolBase_Execute = MakeError("Error: " & Err.Description)
End Function

