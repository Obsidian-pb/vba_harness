VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ListProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_ListProjects - Tool: list all open VBA projects and their components (structure only)
Option Explicit
Implements clsToolBase

' ============================================================
'  Tool: list_vba_projects
'  Returns a list of all open documents (VBA projects)
'  and the modules they contain (names and types only, no code).
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "list_vba_projects"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = _
        "The tool returns a list of all open documents (VBA projects) " & _
        "and the components they contain (modules, classes, forms) - only structure " & _
        "and names. The module code itself is NOT read and NOT passed to the agent. " & _
        "Call this tool when you need to know which documents are open " & _
        "and which VBA modules are in them, before requesting the code of a specific module."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{}," & _
        """required"":[]" & _
    "}"
End Property


Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    Dim result As String
    Dim separator As String
    Dim totalProjects As Long
    Dim totalComponents As Long
    
    On Error GoTo ErrHandler
    
    separator = String(60, "=") & vbCrLf
    result = "Structure of VBA projects of open documents:" & vbCrLf & separator
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Long, j As Long
    
    ' Iterate through all VBA projects in the editor
    For i = 1 To Application.Vbe.VBProjects.Count
        Set vbProj = Application.Vbe.VBProjects(i)
        
        ' If the project is not protected
        If vbProj.Protection = 0 Then
        
            If vbProj.name = HARNESS_NAME Then GoTo NextIteration
            
            totalProjects = totalProjects + 1
            
            ' Determine the document type (by project name)
            Dim docType As String
            docType = GetDocumentType(vbProj)
            
            ' Project header
            result = result & vbCrLf
            result = result & "Document: " & vbProj.name
            If docType <> "" Then
                result = result & "  [" & docType & "]"
            End If
            result = result & vbCrLf
            result = result & String(40, "-") & vbCrLf
            
            ' Iterate through VBA project components
            Dim compCount As Long
            compCount = 0
            
            For Each vbComp In vbProj.VBComponents
                compCount = compCount + 1
                totalComponents = totalComponents + 1
                
                ' Component type
                Dim compTypeName As String
                compTypeName = ComponentTypeName(vbComp.Type)
                
                ' Number of lines of code (just to show the module is not empty)
                Dim lineCount As Long
                lineCount = vbComp.CodeModule.CountOfLines
                
                ' Build the component information string
                result = result & "  " & compCount & ". " & compTypeName & _
                         " """ & vbComp.name & """"
                
                ' If the module is not empty - show the number of lines
                If lineCount > 0 Then
                    result = result & "  (" & lineCount & " lines)"
                Else
                    result = result & "  (empty)"
                End If
                
                ' Show the component file type (forms may have a designer)
                If vbComp.Type = 3 Then ' vbext_ct_MSForm
                    result = result & " [Form Designer]"
                End If
                
                result = result & vbCrLf
            Next vbComp
            
            ' If the project has no components
            If compCount = 0 Then
                result = result & "  (no components)" & vbCrLf
            End If
        
        End If
        
NextIteration:
    Next i
    
    ' If there are no projects at all
    If totalProjects = 0 Then
        result = result & "No open VBA projects." & vbCrLf
    Else
        ' Summary statistics
        result = result & separator
        result = result & "Total: " & totalProjects & " project(s), " & _
                 totalComponents & " component(s)" & vbCrLf
    End If
    
    ' The message is escaped via MakeOk (JsonString), so quotes, backslashes
    ' and newlines in project/module names cannot break the JSON.
    clsToolBase_Execute = MakeOk(result)
    Exit Function

ErrHandler:
    clsToolBase_Execute = MakeError("Error: " & Err.Description)
End Function


' ============================================================
'  HELPER FUNCTIONS
' ============================================================

' Determines the host application type of the project by its properties
Private Function GetDocumentType(ByVal vbProj As Object) As String
    Dim filename As String
    
    On Error Resume Next
    filename = vbProj.filename
    On Error GoTo 0
    
    If filename <> "" Then
        Dim ext As String
        ext = LCase$(Right$(filename, 5))
        
        Select Case True
            Case Right$(ext, 4) = ".xls" Or Right$(ext, 5) = ".xlsx" Or _
                 Right$(ext, 5) = ".xlsm" Or Right$(ext, 5) = ".xlsb"
                GetDocumentType = "Excel Workbook"
            
            Case Right$(ext, 4) = ".doc" Or Right$(ext, 5) = ".docx" Or _
                 Right$(ext, 5) = ".docm"
                GetDocumentType = "Word Document"
            
            Case Right$(ext, 4) = ".vsd" Or Right$(ext, 5) = ".vsdx" Or _
                 Right$(ext, 5) = ".vsdm" Or Right$(ext, 4) = ".vss" Or _
                 Right$(ext, 4) = ".vst"
                GetDocumentType = "Visio Document"
            
            Case Right$(ext, 4) = ".ppt" Or Right$(ext, 5) = ".pptx" Or _
                 Right$(ext, 5) = ".pptm"
                GetDocumentType = "PowerPoint Presentation"
            
            Case Right$(ext, 4) = ".acc" Or Right$(ext, 5) = ".accdb"
                GetDocumentType = "Access Database"
            
            Case Else
                GetDocumentType = "Document (" & ext & ")"
        End Select
    Else
        ' If file name is not available - could be a new document
        GetDocumentType = "New Document"
    End If
End Function

' Returns a human-readable name of the component type
Private Function ComponentTypeName(ByVal compType As Long) As String
    Select Case compType
        Case 1:   ComponentTypeName = "Module"       ' vbext_ct_StdModule
        Case 2:   ComponentTypeName = "Class"        ' vbext_ct_ClassModule
        Case 3:   ComponentTypeName = "Form"         ' vbext_ct_MSForm
        Case 100: ComponentTypeName = "Document"     ' vbext_ct_Document (ThisDocument, Sheet1, etc.)
        Case Else: ComponentTypeName = "Unknown (Type=" & compType & ")"
    End Select
End Function
