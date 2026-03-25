Attribute VB_Name = "modOpenAIResponses"



'=== modOpenAIResponses (讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ) 縺ｫ雋ｼ繧・===
Option Explicit

Private Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/responses"
Private Const OPENAI_MODEL As String = "gpt-4.1-mini" '蠢・ｦ√↑繧・gpt-5.2 遲峨↓螟画峩蜿ｯ :contentReference[oaicite:1]{index=1}

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

    OpenAI_BuildDraft = JsonGetOutputText(http.ResponseText) ' response.output_text 逶ｸ蠖薙ｒ謚懊￥ :contentReference[oaicite:2]{index=2}
    Exit Function

EH:
    OpenAI_BuildDraft = "#ERR " & Err.Number & ": " & Err.Description
End Function

Private Function GetOpenAIApiKey() As String
    ' 蜷榊燕螳夂ｾｩ OPENAI_API_KEY 縺ｮ繧ｻ繝ｫ縺ｫ繧ｭ繝ｼ譁・ｭ怜・繧貞・繧後※縺・ｋ蜑肴署・亥燕縺ｫ菴懊▲縺溘ｄ縺､・・
    GetOpenAIApiKey = CStr(ThisWorkbook.names("OPENAI_API_KEY").RefersToRange.value)
    GetOpenAIApiKey = Trim$(GetOpenAIApiKey)
    If Len(GetOpenAIApiKey) = 0 Then Err.Raise vbObjectError + 514, "GetOpenAIApiKey", "OPENAI_API_KEY 縺檎ｩｺ縺ｧ縺・
End Function

Private Function JsonEsc(ByVal s As String) As String
    ' JSON譁・ｭ怜・逕ｨ縺ｮ譛蟆上お繧ｹ繧ｱ繝ｼ繝・
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEsc = s
End Function

Private Function JsonGetOutputText(ByVal json As String) As String
    ' Responses API: output[] -> message -> content[] -> output_text -> text 繧呈栢縺擾ｼ域怙蟆丞ｮ溯｣・ｼ・
    Dim p As Long, q As Long, k As String

    ' 縺ｾ縺・output_text 繝悶Ο繝・け繧呈爾縺・
    p = InStr(1, json, """type"": ""output_text""", vbTextCompare)
    If p = 0 Then p = InStr(1, json, """type"":""output_text""", vbTextCompare)
    If p = 0 Then
        JsonGetOutputText = ""
        Exit Function
    End If

    ' 縺昴・霑代￥縺ｮ "text":"...." 繧呈爾縺・
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

    ' 邨らｫｯ縺ｮ " 繧呈爾縺呻ｼ亥腰邏皮沿・壹お繧ｹ繧ｱ繝ｼ繝玲悴蟇ｾ蠢懊√∪縺壹・逍朱夂｢ｺ隱咲畑・・
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




Public Function GenerateBasicPlanNarrative(ByVal planStructure As Object) As Object
    Dim systemInstructions As String
    Dim userInput As String
    Dim draft As Object

    systemInstructions = BuildBasicSystemPrompt()
    userInput = BuildBasicUserPrompt(planStructure)

    Set draft = CreateObject("Scripting.Dictionary")
    draft("PlanText") = OpenAI_BuildDraft(systemInstructions, userInput)
    draft("MonitoringText") = OpenAI_BuildDraft(systemInstructions, userInput & vbCrLf & "[task] monitoring")

    Set GenerateBasicPlanNarrative = draft
End Function

Private Function BuildBasicSystemPrompt() As String
    BuildBasicSystemPrompt = _
        "縺ゅ↑縺溘・蛹ｻ逋りｨ育判譖ｸ縺ｮ譁・ｫ菴懈・蟆ら畑繧｢繧ｷ繧ｹ繧ｿ繝ｳ繝医〒縺吶・ & _
        "蛻､螳壹・險ｺ譁ｭ繝ｻ謨ｰ蛟､隗｣驥医・陦後ｏ縺壹∝・蜉帶ｸ医∩縺ｮ讒矩蛹悶ョ繝ｼ繧ｿ繧定・辟ｶ譁・↓謨ｴ蠖｢縺励※縺上□縺輔＞縲・
End Function

Private Function BuildBasicUserPrompt(ByVal planStructure As Object) As String
    Dim k As Variant
    Dim lines As Collection
    Dim i As Long
    Dim arr() As String

    Set lines = New Collection
    lines.Add "莉･荳九・VBA蛻､螳壽ｸ医∩繝・・繧ｿ縺ｧ縺吶ょ愛譁ｭ繧定ｿｽ蜉縺帙★譁・ｫ蛹悶・縺ｿ陦後▲縺ｦ縺上□縺輔＞縲・

    For Each k In planStructure.keys
        lines.Add CStr(k) & ": " & CStr(planStructure(k))
    Next k

    ReDim arr(1 To lines.count)
    For i = 1 To lines.count
        arr(i) = lines(i)
    Next i

    BuildBasicUserPrompt = Join(arr, vbCrLf)
End Function

