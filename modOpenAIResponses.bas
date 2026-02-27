Attribute VB_Name = "modOpenAIResponses"



'=== modOpenAIResponses (標準モジュール) に貼る ===
Option Explicit

Private Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/responses"
Private Const OPENAI_MODEL As String = "gpt-4.1-mini" '必要なら gpt-5.2 等に変更可 :contentReference[oaicite:1]{index=1}

Public Function OpenAI_BuildDraft(ByVal systemInstructions As String, ByVal userInput As String) As String
    On Error GoTo EH

    Dim apiKey As String
    apiKey = GetOpenAIApiKey()

    Dim body As String
    body = "{""model"":""" & JsonEsc(OPENAI_MODEL) & """," & _
           """instructions"":""" & JsonEsc(systemInstructions) & """," & _
           """input"":""" & JsonEsc(userInput) & """}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", OPENAI_ENDPOINT, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.Send body

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 513, "OpenAI_BuildDraft", "HTTP " & http.Status & ": " & Left$(http.ResponseText, 500)
    End If

    OpenAI_BuildDraft = JsonGetOutputText(http.ResponseText) ' response.output_text 相当を抜く :contentReference[oaicite:2]{index=2}
    Exit Function

EH:
    OpenAI_BuildDraft = "#ERR " & Err.Number & ": " & Err.Description
End Function

Private Function GetOpenAIApiKey() As String
    ' 名前定義 OPENAI_API_KEY のセルにキー文字列を入れている前提（前に作ったやつ）
    GetOpenAIApiKey = CStr(ThisWorkbook.Names("OPENAI_API_KEY").RefersToRange.value)
    GetOpenAIApiKey = Trim$(GetOpenAIApiKey)
    If Len(GetOpenAIApiKey) = 0 Then Err.Raise vbObjectError + 514, "GetOpenAIApiKey", "OPENAI_API_KEY が空です"
End Function

Private Function JsonEsc(ByVal s As String) As String
    ' JSON文字列用の最小エスケープ
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEsc = s
End Function

Private Function JsonGetOutputText(ByVal json As String) As String
    ' Responses API: output[] -> message -> content[] -> output_text -> text を抜く（最小実装）
    Dim p As Long, q As Long, k As String

    ' まず output_text ブロックを探す
    p = InStr(1, json, """type"": ""output_text""", vbTextCompare)
    If p = 0 Then p = InStr(1, json, """type"":""output_text""", vbTextCompare)
    If p = 0 Then
        JsonGetOutputText = ""
        Exit Function
    End If

    ' その近くの "text":"...." を探す
    k = """text"": """
    q = InStr(p, json, k, vbTextCompare)
    If q = 0 Then
        k = """text"":"""
        q = InStr(p, json, k, vbTextCompare)
        If q = 0 Then
            JsonGetOutputText = ""
            Exit Function
        End If
    End If
    q = q + Len(k)

    ' 終端の " を探す（単純版：エスケープ未対応、まずは疎通確認用）
    p = InStr(q, json, """", vbBinaryCompare)
    If p = 0 Then
        JsonGetOutputText = ""
        Exit Function
    End If

    JsonGetOutputText = Mid$(json, q, p - q)
    JsonGetOutputText = JsonUnescape(JsonGetOutputText)
End Function








Private Function JsonUnescape(ByVal s As String) As String
    Dim i As Long, n As Long, hex4 As String, ch As String
    Dim out As String

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            ch = Mid$(s, i + 1, 1)
            Select Case ch
                Case "n": out = out & vbCrLf: i = i + 2
                Case "r": i = i + 2
                Case "t": out = out & vbTab: i = i + 2
                Case """": out = out & """": i = i + 2
                Case "\": out = out & "\": i = i + 2
                Case "u"
                    If i + 5 <= Len(s) Then
                        hex4 = Mid$(s, i + 2, 4)
                        If hex4 Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                            n = CLng("&H" & hex4)
                            out = out & ChrW$(n)
                            i = i + 6
                        Else
                            out = out & "\u": i = i + 2
                        End If
                    Else
                        out = out & "\u": i = i + 2
                    End If
                Case Else
                    out = out & "\" & ch: i = i + 2
            End Select
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    JsonUnescape = out
End Function

