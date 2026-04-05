Attribute VB_Name = "modLifeFuncMapping"

Option Explicit

Private Const CAT_BI As String = "BI"
Private Const CAT_IADL As String = "IADL"
Private Const CAT_KYO As String = "KYO"

Private Const FALLBACK_TEXT As String = ""
Private Const FALLBACK_NUM As Long = -1

' ==================== Public API: BI ====================

' BIキー+点数からWordレベル文字列へ変換
Public Function LFM_BIWordLevel(ByVal biKey As String, ByVal scoreValue As Variant, Optional ByVal Fallback As String = FALLBACK_TEXT) As String
    Dim scoreText As String
    scoreText = Trim$(CStr(scoreValue))

    Select Case UCase$(Trim$(biKey))
        Case "BI_0"
            LFM_BIWordLevel = BILevel_3(scoreText, "全介助", "一部介助", "自立", Fallback)
        Case "BI_1"
            Select Case scoreText
                Case "0": LFM_BIWordLevel = "全介助"
                Case "5": LFM_BIWordLevel = "座れるが移れない"
                Case "10": LFM_BIWordLevel = "監視下"
                Case "15": LFM_BIWordLevel = "自立"
                Case Else: LFM_BIWordLevel = Fallback
            End Select
        Case "BI_2"
            LFM_BIWordLevel = BILevel_2(scoreText, "全介助", "自立", Fallback)
        Case "BI_3"
            LFM_BIWordLevel = BILevel_3(scoreText, "全介助", "一部介助", "自立", Fallback)
        Case "BI_4"
            LFM_BIWordLevel = BILevel_2(scoreText, "全介助", "自立", Fallback)
        Case "BI_5"
            Select Case scoreText
                Case "0": LFM_BIWordLevel = "全介助"
                Case "5": LFM_BIWordLevel = "車椅子操作が可能"
                Case "10": LFM_BIWordLevel = "歩行器等"
                Case "15": LFM_BIWordLevel = "自立"
                Case Else: LFM_BIWordLevel = Fallback
            End Select
        Case "BI_6", "BI_7", "BI_8", "BI_9"
            LFM_BIWordLevel = BILevel_3(scoreText, "全介助", "一部介助", "自立", Fallback)
        Case Else
            LFM_BIWordLevel = Fallback
    End Select
End Function

' Word項目名からBI情報を返す
' 戻り値: Array(found As Boolean, Category, WordItem, SourceKey, ControlName, UILabel, ScoreCandidatesCsv)
Public Function LFM_FindBIByWordItem(ByVal wordItem As String) As Variant
    Dim itemNorm As String
    itemNorm = NormalizeWordItem(wordItem)

    Select Case itemNorm
        Case "食事"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "食事", "BI_0", "cmbBI_0", "摂食", "0,5,10")
        Case "椅子とベッド間の移乗"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "椅子とベッド間の移乗", "BI_1", "cmbBI_1", "車いす-ベッド移乗", "0,5,10,15")
        Case "整容"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "整容", "BI_2", "cmbBI_2", "整容", "0,5")
        Case "トイレ動作"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "トイレ動作", "BI_3", "cmbBI_3", "トイレ動作", "0,5,10")
        Case "入浴"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "入浴", "BI_4", "cmbBI_4", "入浴", "0,5")
        Case "平地歩行"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "平地歩行", "BI_5", "cmbBI_5", "歩行/車いす移動", "0,5,10,15")
        Case "階段昇降"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "階段昇降", "BI_6", "cmbBI_6", "階段昇降", "0,5,10")
        Case "更衣"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "更衣", "BI_7", "cmbBI_7", "更衣", "0,5,10")
        Case "排便コントロール"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "排便コントロール", "BI_8", "cmbBI_8", "排便コントロール", "0,5,10")
        Case "排尿コントロール"
            LFM_FindBIByWordItem = Array(True, CAT_BI, "排尿コントロール", "BI_9", "cmbBI_9", "排尿コントロール", "0,5,10")
        Case Else
            LFM_FindBIByWordItem = Array(False, CAT_BI, "", "", "", "", "")
    End Select
End Function

' BI 10項目テーブル
' 列: Category, WordItem, SourceKey, ControlName, UILabel, ScoreCandidatesCsv, MatchType, Note
Public Function LFM_GetBITable() As Variant
    Dim rows As Variant
    rows = Array( _
        Array(CAT_BI, "食事", "BI_0", "cmbBI_0", "摂食", "0,5,10", "文言差", "摂食⇔食事"), _
        Array(CAT_BI, "椅子とベッド間の移乗", "BI_1", "cmbBI_1", "車いす-ベッド移乗", "0,5,10,15", "文言差", "文言差吸収"), _
        Array(CAT_BI, "整容", "BI_2", "cmbBI_2", "整容", "0,5", "完全一致", "0点は全介助固定"), _
        Array(CAT_BI, "トイレ動作", "BI_3", "cmbBI_3", "トイレ動作", "0,5,10", "完全一致", ""), _
        Array(CAT_BI, "入浴", "BI_4", "cmbBI_4", "入浴", "0,5", "完全一致", "0点は全介助固定"), _
        Array(CAT_BI, "平地歩行", "BI_5", "cmbBI_5", "歩行/車いす移動", "0,5,10,15", "概念差", "歩行/車いす移動を平地歩行へ対応"), _
        Array(CAT_BI, "階段昇降", "BI_6", "cmbBI_6", "階段昇降", "0,5,10", "完全一致", ""), _
        Array(CAT_BI, "更衣", "BI_7", "cmbBI_7", "更衣", "0,5,10", "完全一致", ""), _
        Array(CAT_BI, "排便コントロール", "BI_8", "cmbBI_8", "排便コントロール", "0,5,10", "完全一致", ""), _
        Array(CAT_BI, "排尿コントロール", "BI_9", "cmbBI_9", "排尿コントロール", "0,5,10", "完全一致", "") _
    )

    LFM_GetBITable = rows
End Function

' ==================== Public API: IADL ====================

' IADL文字列を 2 / 1 / 0 へ正規化
Public Function LFM_NormalizeIADLLevel(ByVal levelText As String, Optional ByVal Fallback As Long = FALLBACK_NUM) As Long
    LFM_NormalizeIADLLevel = NormalizeAssistLevel(levelText, Fallback)
End Function

' Word項目名からIADL情報を返す
' 戻り値: Array(found As Boolean, Category, WordItem, SourceKey, ControlName)
Public Function LFM_FindIADLByWordItem(ByVal wordItem As String) As Variant
    Select Case NormalizeWordItem(wordItem)
        Case "調理"
            LFM_FindIADLByWordItem = Array(True, CAT_IADL, "調理", "IADL_0", "cmbIADL_0")
        Case "洗濯"
            LFM_FindIADLByWordItem = Array(True, CAT_IADL, "洗濯", "IADL_1", "cmbIADL_1")
        Case "掃除"
            LFM_FindIADLByWordItem = Array(True, CAT_IADL, "掃除", "IADL_2", "cmbIADL_2")
        Case Else
            LFM_FindIADLByWordItem = Array(False, CAT_IADL, "", "", "")
    End Select
End Function

' IADL 3項目テーブル
' 列: Category, WordItem, SourceKey, ControlName, ValueCandidatesCsv, Note
Public Function LFM_GetIADLTable() As Variant
    LFM_GetIADLTable = Array( _
        Array(CAT_IADL, "調理", "IADL_0", "cmbIADL_0", "自立,見守り（監視下）,一部介助,全介助", ""), _
        Array(CAT_IADL, "洗濯", "IADL_1", "cmbIADL_1", "自立,見守り（監視下）,一部介助,全介助", ""), _
        Array(CAT_IADL, "掃除", "IADL_2", "cmbIADL_2", "自立,見守り（監視下）,一部介助,全介助", "") _
    )
End Function

' ==================== Public API: 起居動作 ====================

' 起居文字列を文言吸収付きで 2 / 1 / 0 へ正規化
Public Function LFM_NormalizeKyoLevel(ByVal levelText As String, Optional ByVal Fallback As Long = FALLBACK_NUM) As Long
    Dim canon As String
    canon = CanonicalizeLevelText(levelText)
    LFM_NormalizeKyoLevel = NormalizeAssistLevel(canon, Fallback)
End Function

' 起居項目名の文言吸収（座位保持→座位、立位保持→立位）
Public Function LFM_CanonicalizeKyoItem(ByVal itemName As String) As String
    Dim s As String
    s = NormalizeWordItem(itemName)

    Select Case s
        Case "座位保持": LFM_CanonicalizeKyoItem = "座位"
        Case "立位保持": LFM_CanonicalizeKyoItem = "立位"
        Case Else: LFM_CanonicalizeKyoItem = s
    End Select
End Function

' Word項目名から起居情報を返す
' 戻り値: Array(found As Boolean, Category, WordItem, SourceKey, ControlName)
Public Function LFM_FindKyoByWordItem(ByVal wordItem As String) As Variant
    Select Case LFM_CanonicalizeKyoItem(wordItem)
        Case "寝返り"
            LFM_FindKyoByWordItem = Array(True, CAT_KYO, "寝返り", "Kyo_Roll", "cmbKyo_Roll")
        Case "起き上がり"
            LFM_FindKyoByWordItem = Array(True, CAT_KYO, "起き上がり", "Kyo_SitUp", "cmbKyo_SitUp")
        Case "座位"
            LFM_FindKyoByWordItem = Array(True, CAT_KYO, "座位", "Kyo_SitHold", "cmbKyo_SitHold")
        Case "立ち上がり"
            LFM_FindKyoByWordItem = Array(True, CAT_KYO, "立ち上がり", "Kyo_StandUp", "unnamed-right-of-label")
        Case "立位"
            LFM_FindKyoByWordItem = Array(True, CAT_KYO, "立位", "Kyo_StandHold", "unnamed-right-of-label")
        Case Else
            LFM_FindKyoByWordItem = Array(False, CAT_KYO, "", "", "")
    End Select
End Function

' 起居 5項目テーブル
' 列: Category, WordItem, SourceKey, ControlName, ValueCandidatesCsv, AliasNote
Public Function LFM_GetKyoTable() As Variant
    LFM_GetKyoTable = Array( _
        Array(CAT_KYO, "寝返り", "Kyo_Roll", "cmbKyo_Roll", "自立,見守り（監視下）,一部介助,全介助", ""), _
        Array(CAT_KYO, "起き上がり", "Kyo_SitUp", "cmbKyo_SitUp", "自立,見守り（監視下）,一部介助,全介助", ""), _
        Array(CAT_KYO, "座位", "Kyo_SitHold", "cmbKyo_SitHold", "自立,見守り（監視下）,一部介助,全介助", "座位保持→座位"), _
        Array(CAT_KYO, "立ち上がり", "Kyo_StandUp", "unnamed-right-of-label", "自立,見守り（監視下）,一部介助,全介助", ""), _
        Array(CAT_KYO, "立位", "Kyo_StandHold", "unnamed-right-of-label", "自立,見守り（監視下）,一部介助,全介助", "立位保持→立位") _
    )
End Function

' ==================== Public API: 共通 ====================

' Word項目名からカテゴリを返す（BI / IADL / KYO / ""）
Public Function LFM_CategoryByWordItem(ByVal wordItem As String) As String
    Dim v As Variant

    v = LFM_FindBIByWordItem(wordItem)
    If CBool(v(0)) Then
        LFM_CategoryByWordItem = CAT_BI
        Exit Function
    End If

    v = LFM_FindIADLByWordItem(wordItem)
    If CBool(v(0)) Then
        LFM_CategoryByWordItem = CAT_IADL
        Exit Function
    End If

    v = LFM_FindKyoByWordItem(wordItem)
    If CBool(v(0)) Then
        LFM_CategoryByWordItem = CAT_KYO
        Exit Function
    End If

    LFM_CategoryByWordItem = FALLBACK_TEXT
End Function

' 共通: 介助レベル正規化（IADL/KYOで共用）
Public Function LFM_NormalizeAssistLevel(ByVal levelText As String, Optional ByVal Fallback As Long = FALLBACK_NUM) As Long
    LFM_NormalizeAssistLevel = NormalizeAssistLevel(levelText, Fallback)
End Function

' デバッグ: 全テーブルと主要変換を出力
Public Sub LFM_DebugDump()
    Debug.Print "=== [LFM Debug Dump] ==="
    Debug.Print "[Category] 食事=" & LFM_CategoryByWordItem("食事")
    Debug.Print "[Category] 調理=" & LFM_CategoryByWordItem("調理")
    Debug.Print "[Category] 座位保持=" & LFM_CategoryByWordItem("座位保持")

    Debug.Print "[BI] BI_0:10 => " & LFM_BIWordLevel("BI_0", 10, "?")
    Debug.Print "[BI] BI_1:5  => " & LFM_BIWordLevel("BI_1", 5, "?")
    Debug.Print "[IADL] 見守り（監視下） => " & LFM_NormalizeIADLLevel("見守り（監視下）", -1)
    Debug.Print "[KYO] 見守り => " & LFM_NormalizeKyoLevel("見守り", -1)
    Debug.Print "[KYO Item] 立位保持 => " & LFM_CanonicalizeKyoItem("立位保持")

    DumpRows "BI", LFM_GetBITable()
    DumpRows "IADL", LFM_GetIADLTable()
    DumpRows "KYO", LFM_GetKyoTable()
    Debug.Print "=== [/LFM Debug Dump] ==="
End Sub

' ==================== Private ====================

Private Function NormalizeAssistLevel(ByVal levelText As String, ByVal Fallback As Long) As Long
    Dim s As String
    s = CanonicalizeLevelText(levelText)

    Select Case s
        Case "自立"
            NormalizeAssistLevel = 2
        Case "見守り", "見守り（監視下）", "一部介助"
            NormalizeAssistLevel = 1
        Case "全介助"
            NormalizeAssistLevel = 0
        Case Else
            NormalizeAssistLevel = Fallback
    End Select
End Function

Private Function CanonicalizeLevelText(ByVal levelText As String) As String
    Dim s As String
    s = NormalizeWordItem(levelText)

    Select Case s
        Case "見守り"
            CanonicalizeLevelText = "見守り（監視下）"
        Case Else
            CanonicalizeLevelText = s
    End Select
End Function

Private Function NormalizeWordItem(ByVal src As String) As String
    Dim s As String
    s = Trim$(src)

    s = Replace(s, "　", " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeWordItem = Trim$(s)
End Function

Private Function BILevel_3(ByVal scoreText As String, ByVal v0 As String, ByVal v5 As String, ByVal v10 As String, ByVal Fallback As String) As String
    Select Case scoreText
        Case "0": BILevel_3 = v0
        Case "5": BILevel_3 = v5
        Case "10": BILevel_3 = v10
        Case Else: BILevel_3 = Fallback
    End Select
End Function

Private Function BILevel_2(ByVal scoreText As String, ByVal v0 As String, ByVal v5 As String, ByVal Fallback As String) As String
    Select Case scoreText
        Case "0": BILevel_2 = v0
        Case "5": BILevel_2 = v5
        Case Else: BILevel_2 = Fallback
    End Select
End Function

Private Sub DumpRows(ByVal title As String, ByVal rows As Variant)
    On Error GoTo EH

    Dim i As Long, j As Long
    Dim line As String

    Debug.Print "[Table] " & title & " rows=" & CStr(UBound(rows) - LBound(rows) + 1)
    For i = LBound(rows) To UBound(rows)
        line = "  - "
        For j = LBound(rows(i)) To UBound(rows(i))
            If j > LBound(rows(i)) Then line = line & " | "
            line = line & CStr(rows(i)(j))
        Next j
        Debug.Print line
    Next i
    Exit Sub
EH:
    Debug.Print "[Table] " & title & " <dump error>"
End Sub


