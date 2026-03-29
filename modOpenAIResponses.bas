Attribute VB_Name = "modOpenAIResponses"



'=== modOpenAIResponses (標準モジュール) に貼る ===
Option Explicit

Private Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/responses"
Private Const OPENAI_MODEL As String = "gpt-4.1"

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
    GetOpenAIApiKey = CStr(ThisWorkbook.names("OPENAI_API_KEY").RefersToRange.value)
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




' 変化・課題を評価値比較からAI生成（2回目以降のみ呼ばれる）
Public Function GenerateChangeAndIssue(ByVal planStructure As Object, ByVal prevSnap As Object) As Object
    On Error GoTo EH
    Dim raw As String
    raw = OpenAI_BuildDraft(BuildChangeAndIssueSystemPrompt(), BuildChangeAndIssueUserPrompt(planStructure, prevSnap))
    Set GenerateChangeAndIssue = ParseChangeAndIssueLines(raw)
    Exit Function
EH:
    Err.Clear
End Function

Private Function BuildChangeAndIssueSystemPrompt() As String
    Dim s As String
    ' ブロック1：役割・出力形式
    s = "あなたは個別機能訓練計画書の実施後対応欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【出力形式】必ず以下の2行のみを出力してください。" & _
        "変化: （内容）" & vbLf & _
        "課題: （内容）" & vbLf
    ' ブロック2：記述ルール・制約
    s = s & _
        "【変化の記述】前回と今回の評価値を比較し、改善・維持・悪化を具体的な数値とともに記述する。" & _
        "変化の内容と程度を明示する。2文程度。" & _
        "【課題の記述】改善が不十分な項目・残存している制限・今後注意すべき点を記述する。" & _
        "目標達成に向けて継続が必要な理由を具体的に示す。2文程度。" & _
        "【専門用語の禁止】MMT・ROM・TUG等の医療用語は使わず平易な表現で記述する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。" & _
        "【言語】日本語のみ。箇条書き不可。"
    BuildChangeAndIssueSystemPrompt = s
End Function

Private Function BuildChangeAndIssueUserPrompt(ByVal planStructure As Object, ByVal prevSnap As Object) As String
    Dim lines As Collection
    Set lines = New Collection

    ' 今回の訓練目標・評価（planStructureから）
    lines.Add "【今回の訓練目標】"
    If planStructure.exists("Activity_Long") Then lines.Add "長期目標: " & CStr(planStructure("Activity_Long"))
    If planStructure.exists("Activity_Short") Then lines.Add "短期目標: " & CStr(planStructure("Activity_Short"))
    If planStructure.exists("MainCause") Then lines.Add "主因: " & CStr(planStructure("MainCause"))
    If planStructure.exists("MMT_TargetMuscle_Score") Then lines.Add "筋力スコア: " & CStr(planStructure("MMT_TargetMuscle_Score"))

    ' 前回の評価値（EV_XXXXの最新保存行から）
    lines.Add ""
    lines.Add "【前回（" & CStr(prevSnap("EvalDate")) & "）の評価値】"
    AddCompField lines, prevSnap, "BITotal", "日常生活動作合計点"
    AddCompField lines, prevSnap, "Test_TUG_sec", "立ち上がりから歩行テスト（秒）"
    AddCompField lines, prevSnap, "Test_10MWalk_sec", "10m歩行テスト（秒）"
    AddCompField lines, prevSnap, "Test_Grip_R_kg", "握力右（kg）"
    AddCompField lines, prevSnap, "Test_Grip_L_kg", "握力左（kg）"
    AddCompField lines, prevSnap, "Test_5xSitStand_sec", "5回立ち座りテスト（秒）"

    Dim arr() As String
    ReDim arr(1 To lines.count)
    Dim i As Long
    For i = 1 To lines.count
        arr(i) = lines(i)
    Next i
    BuildChangeAndIssueUserPrompt = Join(arr, vbCrLf)
End Function

Private Sub AddCompField(ByVal lines As Collection, ByVal d As Object, ByVal key As String, ByVal label As String)
    If d.exists(key) Then lines.Add label & ": " & CStr(d(key))
End Sub

Private Function ParseChangeAndIssueLines(ByVal raw As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim lines() As String
    lines = Split(raw, vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(Replace(lines(i), vbCr, ""))
        Dim cp As Long: cp = InStr(1, ln, ": ", vbBinaryCompare)
        If cp = 0 Then GoTo NextCILine
        Dim lbl As String: lbl = Trim$(Left$(ln, cp - 1))
        Dim val As String: val = Trim$(Mid$(ln, cp + 2))
        If LenB(val) = 0 Then GoTo NextCILine
        Select Case lbl
            Case "変化": d("Change") = val
            Case "課題": d("Issue") = val
        End Select
NextCILine:
    Next i
    Set ParseChangeAndIssueLines = d
End Function

Public Function GenerateBasicPlanNarrative(ByVal planStructure As Object) As Object
    Dim userInput As String
    Dim draft As Object

    userInput = BuildBasicUserPrompt(planStructure)

    Set draft = CreateObject("Scripting.Dictionary")
    draft("MonitoringText") = OpenAI_BuildDraft(BuildMonitoringSystemPrompt(), userInput)

    ' 目標文6項目をAI生成して追加
    Dim goalRaw As String
    goalRaw = OpenAI_BuildDraft(BuildGoalSystemPrompt(), userInput)
    Dim goals As Object
    Set goals = ParseGoalLines(goalRaw)
    Dim gk As Variant
    For Each gk In goals.keys
        draft(CStr(gk)) = goals(CStr(gk))
    Next gk

    ' プログラム内容①?⑤をAI生成して追加
    Dim programRaw As String
    programRaw = OpenAI_BuildDraft(BuildProgramsSystemPrompt(), userInput)
    Dim programs As Object
    Set programs = ParseProgramLines(programRaw)
    Dim pk As Variant
    For Each pk In programs.keys
        draft(CStr(pk)) = programs(CStr(pk))
    Next pk

    ' 利用者本人・家族が実施することをAI生成して追加
    draft("HomeExercise") = OpenAI_BuildDraft(BuildHomeExerciseSystemPrompt(), userInput)

    Set GenerateBasicPlanNarrative = draft
End Function

' 目標文6項目用システムプロンプト（行継続24回制限のため3ブロックに分割）
Private Function BuildGoalSystemPrompt() As String
    Dim s As String

    ' ブロック1：役割・出力形式
    s = "あなたは個別機能訓練計画書の目標欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【出力形式】必ず以下の6行のみを出力してください。" & _
        "各行はラベルで始まり、コロンと半角スペースの後に目標文が続きます。" & _
        "機能長期: （文）" & vbLf & _
        "機能短期: （文）" & vbLf & _
        "活動長期: （文）" & vbLf & _
        "活動短期: （文）" & vbLf & _
        "参加長期: （文）" & vbLf & _
        "参加短期: （文）" & vbLf

    ' ブロック2：文構造・専門用語禁止
    s = s & _
        "【文の構造（全項目共通・最重要）】" & _
        "必ず「〔現状や理由〕のため、〔達成したいこと〕を目標とする。」という構造で1文にまとめること。" & _
        "例：「足を横に開く動きの筋力が不足しているため、日常的な歩行動作で安定した力を発揮できるようになることを目標とする。」" & _
        "【専門用語の禁止（絶対厳守）】" & _
        "・「MMT」「ROM」「随意収縮」等の医療専門用語は一切使用禁止。" & _
        "・MMT_TargetMuscle_Scoreは内部参照のみ。スコア数値は文中に出力しない。" & _
        "・筋名は平易な動作表現に置き換える。例：股外転は足を横に開く動き、背屈はつま先を上げる動き、膝伸展は膝を伸ばす動き、腸腰筋は足を前に上げる動き。" & _
        "・TrunkROM_LimitedValuesがある場合も角度数値は書かず前屈みの動きが制限されているのように表現する。"

    ' ブロック3：各目標記述ルール・制約
    s = s & _
        "【機能目標の記述ルール】" & _
        "・短期（1ヶ月）：現状の問題を理由として、安定した動作の獲得など中間段階の到達像を示す。" & _
        "・長期（3ヶ月）：現状の問題を理由として、目標動作に必要な力や動きが十分に回復した最終状態を示す。" & _
        "・短期と長期は明確に段階が異なる内容にすること（同じ内容の繰り返し禁止）。" & _
        "【活動・参加目標の記述ルール】" & _
        "・活動目標：Activity_Longの動作が困難な現状を理由として、安全に行えるようになる到達像を記述する。短期は中間段階、長期は自立した状態。" & _
        "・参加目標：活動目標達成後の生活範囲・社会参加の拡大をNeedPatient/NeedFamilyの希望を踏まえて記述する。" & _
        "・MainCauseが疼痛なら痛みを理由に、麻痺なら動きの弱さを理由に、困難度なら動作のふらつき・不安定さを理由に記述する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。判断・推測・補完は禁止。" & _
        "【言語】日本語のみ。各1文。箇条書き不可。"

    BuildGoalSystemPrompt = s
End Function

' AI出力をパースして目標Dictionaryに変換
Private Function ParseGoalLines(ByVal raw As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim lines() As String
    lines = Split(raw, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(Replace(lines(i), vbCr, ""))
        If LenB(ln) = 0 Then GoTo NextGoalLine

        ' ": " または "：" で分割
        Dim cp As Long
        cp = InStr(1, ln, ": ", vbBinaryCompare)
        If cp = 0 Then cp = InStr(1, ln, "：", vbBinaryCompare)
        If cp = 0 Then GoTo NextGoalLine

        Dim lbl As String
        Dim val As String
        lbl = Trim$(Left$(ln, cp - 1))
        val = Trim$(Mid$(ln, cp + 2))
        If InStr(1, ln, "：", vbBinaryCompare) > 0 And cp = InStr(1, ln, "：", vbBinaryCompare) Then
            val = Trim$(Mid$(ln, cp + 1)) ' 全角コロンは1文字
        End If
        If LenB(val) = 0 Then GoTo NextGoalLine

        Select Case lbl
            Case "機能長期": d("Function_Long") = val
            Case "機能短期": d("Function_Short") = val
            Case "活動長期": d("Activity_Long") = val
            Case "活動短期": d("Activity_Short") = val
            Case "参加長期": d("Participation_Long") = val
            Case "参加短期": d("Participation_Short") = val
        End Select
NextGoalLine:
    Next i

    Set ParseGoalLines = d
End Function

' 利用者本人・家族が実施すること用システムプロンプト
Private Function BuildHomeExerciseSystemPrompt() As String
    Dim s As String
    s = "あなたは個別機能訓練計画書の自主練習欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【記述内容】サービス利用時間外に利用者本人または家族が自宅で安全に実施できる運動を1?2文で記述する。" & _
        "MMT_TargetMuscleの筋を意識した動作で、道具不要・安全・簡単なものにすること。" & _
        "「〇〇（目的）のため、△△（具体的な動作）を1日〇回程度行う。」という形式で記述する。"
    s = s & _
        "【専門用語の禁止】MMT・ROM等の医療用語は使わず平易な動作表現で記述する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。判断・推測・補完は禁止。" & _
        "【言語】日本語のみ。箇条書き不可。"
    BuildHomeExerciseSystemPrompt = s
End Function

' プログラム内容①?⑤用システムプロンプト（行継続24回制限のため2ブロックに分割）
Private Function BuildProgramsSystemPrompt() As String
    Dim s As String

    ' ブロック1：役割・出力形式・各項目の役割定義
    s = "あなたは個別機能訓練計画書のプログラム内容欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【出力形式】必ず以下の5行のみを出力してください。" & _
        "①: （内容）" & vbLf & _
        "②: （内容）" & vbLf & _
        "③: （内容）" & vbLf & _
        "④: （内容）" & vbLf & _
        "⑤: （内容）" & vbLf

    ' ブロック2：各項目の役割（目的レベルで異なる内容にする）
    s = s & _
        "【各項目の役割（必ず守ること）】" & _
        "①: 対象筋への抵抗運動による筋力強化（MMT_TargetMuscleの筋に負荷をかける運動）。" & _
        "②: 対象筋および関連筋のストレッチ・柔軟性改善（動きの制限を緩める運動）。" & _
        "③: バランス・協調性の訓練（安定した姿勢保持や重心移動の練習）。" & _
        "④: Activity_Longの動作そのものの練習（目標動作を段階的に繰り返す）。" & _
        "⑤: 日常生活への応用・定着（NeedPatientの希望に沿った実生活での活用練習）。"

    ' ブロック3：文形式・禁止事項・制約
    s = s & _
        "【文形式】「〇〇を目的に、△△を実施する。」という1文。各項目で目的も運動内容も異なること。" & _
        "【専門用語の禁止】MMT・ROM等の医療用語は使わず平易な動作表現で記述する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。判断・推測・補完は禁止。" & _
        "【言語】日本語のみ。各1文。箇条書き不可。"

    BuildProgramsSystemPrompt = s
End Function

' AI出力をパースしてプログラムDictionaryに変換
Private Function ParseProgramLines(ByVal raw As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim lines() As String
    lines = Split(raw, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(Replace(lines(i), vbCr, ""))
        If LenB(ln) = 0 Then GoTo NextProgLine

        Dim cp As Long
        cp = InStr(1, ln, ": ", vbBinaryCompare)
        If cp = 0 Then GoTo NextProgLine

        Dim lbl As String
        Dim val As String
        lbl = Trim$(Left$(ln, cp - 1))
        val = Trim$(Mid$(ln, cp + 2))
        If LenB(val) = 0 Then GoTo NextProgLine

        Select Case lbl
            Case "①": d("Program1Content") = val
            Case "②": d("Program2Content") = val
            Case "③": d("Program3Content") = val
            Case "④": d("Program4Content") = val
            Case "⑤": d("Program5Content") = val
        End Select
NextProgLine:
    Next i

    Set ParseProgramLines = d
End Function

' プログラム内容欄（何を目的に・何をする）用システムプロンプト
Private Function BuildPlanTextSystemPrompt() As String
    BuildPlanTextSystemPrompt = _
        "あなたは個別機能訓練計画書のプログラム内容欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【出力形式】「〇〇（機能・活動目標）を目的に、△△（具体的なアプローチ）を実施する。」という形式で2?4文程度。" & _
        "【記述内容】主因（MainCause）・機能目標・活動目標・対象筋（MMT_TargetMuscle）・MMTスコアを根拠に、" & _
        "具体的な運動内容・負荷設定の方向性・動作練習の対象を明示する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。判断・推測・補完は禁止。" & _
        "【言語】日本語のみ。箇条書き不可。連続した文章で記述する。"
End Function

' モニタリング欄（変化・課題の観察ポイント）用システムプロンプト
Private Function BuildMonitoringSystemPrompt() As String
    BuildMonitoringSystemPrompt = _
        "あなたは個別機能訓練計画書のモニタリング欄を記述する専門アシスタントです。" & _
        "以下のルールを厳守してください。" & _
        "【出力形式】「〇〇（目標）に向けて、△△（観察・確認すべき変化）を評価する。」という形式で2?3文程度。" & _
        "【記述内容】短期・長期目標の達成に向けて確認すべき身体機能の変化、動作の変化、" & _
        "リスク管理上の観察点（疼痛・疲労・バイタル等）を具体的に記述する。" & _
        "【制約】提供されたデータにない情報は一切追加しない。判断・推測・補完は禁止。" & _
        "【言語】日本語のみ。箇条書き不可。連続した文章で記述する。"
End Function

Private Function BuildBasicUserPrompt(ByVal planStructure As Object) As String
    Dim lines As Collection
    Dim i As Long
    Dim arr() As String

    Set lines = New Collection
    lines.Add "【VBA判定済みデータ（追加判断禁止）】"

    ' 臨床的に意味のあるキーのみ渡す（VBA内部概念語・生成済み目標文は除外）
    Dim clinicalKeys As Variant
    clinicalKeys = Array( _
        "MainCause", _
        "Activity_Long", _
        "MMT_TargetMuscle", _
        "MMT_TargetMuscle_Score", _
        "NeedPatient", _
        "NeedFamily", _
        "TrunkROM_LimitedValues", _
        "EvalTestCriticalFindings" _
    )

    Dim k As Variant
    For Each k In clinicalKeys
        If planStructure.exists(CStr(k)) Then
            Dim v As String
            v = Trim$(CStr(planStructure(CStr(k))))
            If LenB(v) > 0 Then
                lines.Add CStr(k) & ": " & v
            End If
        End If
    Next k

    If planStructure.exists("EvalTestNoteRaw") Then
        Dim note As String
        note = Trim$(CStr(planStructure("EvalTestNoteRaw")))
        If LenB(note) > 0 Then
            lines.Add "[参考：評価メモ原文（補完に使用禁止）] " & note
        End If
    End If

    ReDim arr(1 To lines.count)
    For i = 1 To lines.count
        arr(i) = lines(i)
    Next i

    BuildBasicUserPrompt = Join(arr, vbCrLf)
End Function

