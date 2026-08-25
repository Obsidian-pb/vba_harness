VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHarnessConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsHarnessConsole - Console-based agent runner (debugging via Immediate Window)
Option Explicit

Public WithEvents agent As clsHarness
Attribute agent.VB_VarHelpID = -1


Public Sub RunAgent(ByVal prompt As String)


Dim llm_api_url As String
Dim llm_model_id As String
Dim llm_model_fast_id As String
Dim llm_api_key As String
Dim llm_system_prompt As String

    ' Get LLM settings
    On Error Resume Next
    llm_api_url = CStr(GetSettingFromRegistry(REG_LLM_API_URL, DEF_LLM_API_URL))
    llm_model_id = CStr(GetSettingFromRegistry(REG_LLM_MODEL_ID, DEF_LLM_MODEL_ID))
    llm_model_fast_id = CStr(GetSettingFromRegistry(REG_LLM_MODEL_FAST_ID, DEF_LLM_MODEL_FAST_ID))
    llm_api_key = CStr(GetSettingFromRegistry(REG_LLM_API_KEY, ""))
    llm_system_prompt = CStr(GetSettingFromRegistry(REG_LLM_SYSTEM_PROMPT, ""))
    
    On Error GoTo 0
    
    Set agent = New clsHarness
    agent.Init.Set_LLM llm_api_url, llm_model_id, llm_api_key, llm_system_prompt
    
    
    agent.AgentLoop prompt
End Sub

Private Sub agent_OnAnswer(ByVal s As String)
    Debug.Print "[Answer]" & s
End Sub

Private Sub agent_OnReasoning(ByVal s As String)
    Debug.Print "[Reasoning]" & s
End Sub

Private Sub agent_OnSystem(ByVal s As String)
    Debug.Print "[System]" & s
End Sub

Private Sub agent_OnThinking()
    Debug.Print "[Thinking...]"
End Sub

Private Sub agent_OnToolCalling(ByVal s As String)
    Debug.Print "[Tool call]" & s
End Sub
