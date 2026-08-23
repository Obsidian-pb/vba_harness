Attribute VB_Name = "m_json"
' Module: m_json - JSON utility functions (serialization, deserialization, escaping, extraction)
Option Explicit

' Converts a Collection of JSON strings into a JSON array [...].
Public Function CollectionToJsonArray(ByVal col As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    For i = 1 To col.Count
        result = result & col.Item(i)
        If i < col.Count Then result = result & ","
    Next i
    result = result & "]"
    CollectionToJsonArray = result
End Function

' Wraps a string in JSON quotes with proper escaping.
Public Function JsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr(34), "\" & Chr(34))
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonString = """" & s & """"
End Function

' Universal extractor for a string field from a JSON object.
' Universal extractor for a field value from a JSON object.
' Supports:
'   - quoted strings:    "name":"value"  > "value"
'   - unquoted numbers:  "x1":8.27      > "8.27"
'   - booleans and null: "flag":true     > "true"
' Returns the string representation of the value.
Public Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    Dim searchKey As String, pos As Long, startPos As Long, endPos As Long
    searchKey = """" & key & """:"
    pos = InStr(1, json, searchKey, vbTextCompare)
    If pos = 0 Then ExtractJsonString = "": Exit Function

    startPos = pos + Len(searchKey)
    ' Skip spaces and tabs after the colon
    Do While startPos <= Len(json)
        Dim c As String
        c = Mid$(json, startPos, 1)
        If c <> " " And c <> vbTab Then Exit Do
        startPos = startPos + 1
    Loop

    If startPos > Len(json) Then
        ExtractJsonString = ""
        Exit Function
    End If

    ' Check for null (with protection against out-of-bounds)
    If startPos + 3 <= Len(json) And LCase$(Mid$(json, startPos, 4)) = "null" Then
        ExtractJsonString = ""
        Exit Function
    End If

    Dim firstChar As String
    firstChar = Mid$(json, startPos, 1)

    ' === CASE 1: Value in quotes (string) ===
    ' Process as before: look for the closing quote, handling escaped characters
    If firstChar = """" Then
        startPos = startPos + 1
        endPos = startPos
        Do While endPos <= Len(json)
            If Mid$(json, endPos, 1) = """" And Mid$(json, endPos - 1, 1) <> "\" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJsonString = JsonUnescape(Mid$(json, startPos, endPos - startPos))
        Exit Function
    End If

    ' === CASE 2: Value without quotes (number, true, false) ===
    ' Look for the end of the value: comma, closing brace } or ], or end of string
    endPos = startPos
    Do While endPos <= Len(json)
        Dim ch As String
        ch = Mid$(json, endPos, 1)
        If ch = "," Or ch = "}" Or ch = "]" Then Exit Do
        endPos = endPos + 1
    Loop

    ' Trim any surrounding whitespace
    ExtractJsonString = Trim$(Mid$(json, startPos, endPos - startPos))
End Function

' Unescapes JSON strings (reverse escaping).
Public Function JsonUnescape(ByVal s As String) As String
    ' Temporary marker — using a rare character sequence that definitely won't appear in the code
    Const MARKER As String = "§§BS§§"
    
    ' 1. First, protect escaped slashes: \\ > marker
    s = Replace(s, "\\", MARKER)
    
    ' 2. Now safely process the remaining escape sequences
    s = Replace(s, "\" & Chr(34), Chr(34))   ' \" > "
    s = Replace(s, "\n", vbCrLf)             ' \n > vbCrLf
    s = Replace(s, "\r", vbCr)               ' \r > vbCr
    s = Replace(s, "\t", vbTab)              ' \t > vbTab
    
    ' 3. Restore the marker back to a single backslash
    s = Replace(s, MARKER, "\")
    
    JsonUnescape = s
End Function


