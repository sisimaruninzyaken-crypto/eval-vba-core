Attribute VB_Name = "modEvalIOEntry"


Option Explicit

Public Const EVAL_SHEET_NAME As String = "EvalData"
Private Const EVAL_INDEX_SHEET_NAME As String = "EvalIndex"
Private Const CLIENT_MASTER_SHEET_NAME As String = "ClientMaster"
Private Const EVAL_HISTORY_SHEET_PREFIX As String = "EV_"
Private Const HDR_ROWNO As String = "RowNo"
Private Const HDR_USER_ID As String = "UserID"
Private Const HDR_NAME As String = "Name"
Private Const HDR_KANA As String = "Kana"
Private Const HDR_SHEET As String = "SheetName"
Private Const HDR_FIRST_EVAL As String = "FirstEvalDate"
Private Const HDR_LATEST_EVAL As String = "LatestEvalDate"
Private Const HDR_RECORD_COUNT As String = "RecordCount"
Private Const HDR_BIRTH_DATE As String = "BirthDate"
Private Const HDR_GENDER As String = "Gender"
Private Const HDR_CARE_LEVEL As String = "CareLevel"
Private Const HDR_CREATED_DATE As String = "CreatedDate"
Public mDailyLogManual As Boolean    ' 日々の記録の手動保存フラグ



' === 補助具/リスク フレーム名（固定用） ===
Private Const FRM_AIDS As String = "Frame33"
Private Const FRM_RISK As String = "Frame34"
Private Const IO_TRACE As Boolean = False
Private Const MAIN_SAVE_MIN_FILLED_FIELDS As Long = 10
Private Const MAIN_SAVE_FEW_INPUT_MESSAGE As String = "入力項目が少ない状態です。" & vbCrLf & _
    "既存データを上書きすると元に戻せない可能性があります。" & vbCrLf & _
    "本当に保存しますか？"
Private Const MAIN_SAVE_MIN_CHANGE_COUNT As Long = 3
Private Const MAIN_SAVE_FEW_CHANGE_MESSAGE As String = "変更項目がほとんどありません。" & vbCrLf & _
    "誤って保存しようとしていないか確認してください。" & vbCrLf & _
    "本当に保存しますか？"
Private Const HDR_HOMEENV_CHECKS As String = "Basic.HomeEnv.Checks"
Private Const HDR_HOMEENV_NOTE As String = "Basic.HomeEnv.Note"
Private Const HDR_RISK_CHECKS As String = "Basic.Risk.Checks"
Private Const HDR_AIDS_CHECKS As String = "Basic.Aids.Checks"
Private Const HISTORY_LOAD_DEBUG As Boolean = True


Public Sub LoadEvaluation_CurrentRow()
    MsgBox "この入口は廃止しました。読み込みは「名前→直近候補から選択」に統一しています。", vbInformation
End Sub

' ★ここを退避名にして必ず閉じる
Private Sub LoadEvaluation_fromLastRow_OBSOLETE()
End Sub




Private Sub IO_T(ParamArray a())
    If Not IO_TRACE Then Exit Sub
    Dim i As Long, s As String
    For i = LBound(a) To UBound(a): s = s & IIf(i > 0, " ", "") & CStr(a(i)): Next
    Debug.Print Format(Now, "hh:nn:ss"), s
End Sub

Private Sub IO_SafeRunSave(procName As String, ws As Worksheet, r As Long, owner As Object)
    On Error GoTo EH
    IO_T "[RUN] target", procName

    
    IO_T "[SAVE] call", procName
    Application.Run procName, ws, r, owner
    IO_T "[SAVE] ok", procName
    Exit Sub
EH:
    IO_T "[SAVE] NG", procName, "Err", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub IO_SafeRunLoad(procName As String, ws As Worksheet, r As Long, owner As Object)
    On Error GoTo EH
    IO_T "[LOAD] call", procName
    Application.Run procName, ws, r, owner
    IO_T "[LOAD] ok", procName
    Exit Sub
EH:
    IO_T "[LOAD] NG", procName, "Err", Err.Number, Err.Description
    Err.Clear
End Sub


Private Sub t(ParamArray a())
    If Not IO_TRACE Then Exit Sub
    Dim i As Long, s As String
    For i = LBound(a) To UBound(a)
        s = s & IIf(i > 0, " ", "") & CStr(a(i))
    Next
    Debug.Print Format(Now, "hh:nn:ss"), s
End Sub



' ★Compat：旧入口。内部的には SaveEvaluation_Append_From に委譲する。
' 　どこかのボタンや古いマクロがまだ SaveEvaluation_Append を指していても、
' 　最終的な保存ルートは SaveEvaluation_Append_From に一本化される。
Public Sub SaveEvaluation_Append()
    EnsureFormLoaded                ' frmEval がロードされていなければロード
    SaveEvaluation_Append_From frmEval
End Sub


' ★[OBSOLETE] 直接呼ばない。読み込みは LoadEvaluation_ByName_From に一本化。
Private Sub LoadEvaluation_LastRow_OBSOLETE(owner As Object)

    MsgBox "この入口は廃止しました。読み込みは『名前→直近候補から選択』に統一しています。", vbInformation
End Sub


Private Sub SaveEvaluation_CurrentRow_OBSOLETE()
    MsgBox "この入口は廃止しました。保存は『追加保存（Append）』に統一しています。", vbInformation
End Sub
Private Sub LoadEvaluation_CurrentRow_OBSOLETE()
    ' OBSOLETE: this procedure must not be used.
    Debug.Assert False
    Exit Sub
End Sub

'======================== 実体：全部まとめて呼ぶ ========================

' ===== すべて保存 =====
Public Sub SaveAllSectionsToSheet(ws As Worksheet, r As Long, owner As Object)


   ' 保存ハブ：EvalData 1 行分にまとめて書き込む
' 保存順のイメージ：
'   1) 基本情報（Basic）
'   2) 麻痺 / ROM / 姿勢
'   3) MMT / 感覚 / トーン・反射
'   4) 疼痛（Pain IO）
'   5) テスト・評価（10m / TUG / 握力 / 5回立ち / セミタンデム）
'   6) 補助具 / リスク（チェック群）
'   7) ADL（IO_ADL）

   
   

    ' 基本情報（このモジュール内の実装）
    Call SaveBasicInfoToSheet_FromMe(ws, r, owner)



    ' 麻痺 / ROM（既にOK）
    IO_SafeRunSave "SaveParalysisToSheet", ws, r, owner
    IO_SafeRunSave "SaveROMToSheet", ws, r, owner
    IO_SafeRunSave "SavePostureToSheet", ws, r, owner
    


    ' 必要になったら順次ON
    IO_SafeRunSave "SaveMMTToSheet", ws, r, owner
    IO_SafeRunSave "SaveSensoryToSheet", ws, r, owner
     'Call Mirror_SensoryIO(ws, r)    'Legacy互換：現行仕様では未使用のため停止
    IO_SafeRunSave "modToneReflexIO.SaveToneReflexToSheet", ws, r, owner
  

    Call SavePainToSheet(ws, r, owner)
     Call Save_TestEvalToSheet(ws, r, owner)
     Call Save_WalkIndepToSheet(ws, r, owner)  '★歩行自立度 IO_WalkIndep 保存
     Call Save_WalkAbnToSheet(ws, r, owner)    '★異常歩行 IO_WalkAbn 保存
     Call Save_WalkRLAToSheet(ws, r, owner)    '★RLA IO_WalkRLA 保存



Call Save_ADL_AtRow(ws, r)




End Sub

' ===== すべて読込 =====
'====================================================================
' [HUB] 評価読み込みハブ
'  - 呼び出し元：LoadEvaluation_ByName_From（正規入口）など
'  - 役割：
'       1) 名前から「最新行」に r を差し替える（FindLatestRowByName）
'       2) BasicInfo / ROM / 姿勢 / MMT / 感覚・トーン / 疼痛 /
'          テスト評価 / 歩行 / 認知・精神 など各セクションの
'          Load*FromSheet をまとめて呼び出す
'  - 注意：
'       * 他モジュールからここを直接呼ぶのは極力避ける
'         （読み込み仕様の一元管理のため）
'       * 各セクションの UI レイアウト調整はここでは行わない
'====================================================================
Public Sub LoadAllSectionsFromSheet(ws As Worksheet, r As Long, owner As Object)

    Dim nm As String
    Dim rLatest As Long

    ' ★同じ名前なら、その人の「最新行」に読み込み行を差し替える
         nm = Trim$(owner.txtName.text)

    ' ★フォーム側が空なら、シートの氏名セルから拾う
    If Len(nm) = 0 Then
        Dim cName As Long
        cName = FindHeaderCol(ws, "Basic.Name")
        If cName = 0 Then cName = FindHeaderCol(ws, "氏名")
        If cName = 0 Then cName = FindHeaderCol(ws, "利用者名")
        If cName = 0 Then cName = FindHeaderCol(ws, "名前")


        If cName > 0 Then
            nm = Trim$(CStr(ws.Cells(r, cName).value))
        End If
    End If
    
    

    ' ★入口で r が指定されている場合は尊重する（ここで上書きしない）
If r < 2 And Len(nm) > 0 Then
    rLatest = FindLatestRowByName(ws, nm)
    If rLatest > 0 Then r = rLatest
End If




   ' 麻痺 / ROM / 姿勢の読込は LoadBasicInfoFromSheet_FromMe 内で
    ' chkLoadParalysis / chkLoadROM / chkLoadPosture に応じて実施
    
    Call LoadBasicInfoFromSheet_FromMe(ws, r, owner)
    IO_SafeRunLoad "Load_ADL_FromRow", ws, r, owner
   


    
    'Call LoadParalysisFromSheet(ws, r, owner)
    'Call LoadROMFromSheet(ws, r, owner)
    Call LoadSensoryFromSheet(ws, r, owner)
    'Call LoadPostureFromSheet(ws, r, owner)
    
   
    Call Load_TestEvalFromSheet(ws, r, owner)
    Call Load_WalkIndepFromSheet(ws, r, owner)
    Call Load_WalkAbnFromSheet(ws, r, owner)
    Call Load_WalkRLAFromSheet(ws, r, owner)   '★RLA読み込み

    'Call MMT.LoadMMTFromSheet(ws, r, owner)
    Call modToneReflexIO.LoadToneReflexFromSheet(ws, r, owner)


   

    IO_SafeRunLoad "LoadPainFromSheet", ws, r, owner
    
    ' 補助具
Dim cA As Long
cA = FindHeaderCol(ws, "補助具")
If cA > 0 Then
    DeserializeChecks owner, "Frame33", CStr(ws.Cells(r, cA).value), True   ' 補助具
End If

' リスク
Dim cR As Long
cR = FindHeaderCol(ws, "リスク")
If cR > 0 Then
    DeserializeChecks owner, "Frame34", CStr(ws.Cells(r, cR).value), False  ' リスク
End If
    
        Call Load_CognitionMental_FromRow(ws, r, owner)
        'Load_DailyLog_Latest_FromForm owner
        
End Sub


'====================================================================
' [ENTRY] 評価読み込みの正規入口
'  - UI 側（frmEval や他フォーム）は原則ここだけを呼び出す
'  - 名前（txtName）から EvalData 上の最新行を特定し、
'    LoadAllSectionsFromSheet に委譲する
'  - LoadAllSectionsFromSheet / 各セクションの Load*FromSheet は
'    他モジュールから直接呼ばないこと（読み込み仕様の分裂防止）
'====================================================================
Public Sub LoadEvaluation_ByName_From(owner As Object)



    EnsureFormLoaded
    Dim wsTarget As Worksheet
    Dim resolveMessage As String
    If ResolveUserHistorySheet(owner, False, wsTarget, resolveMessage) Then
        Dim validRow As Long
        Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
        Dim nameVal As String: nameVal = Trim$(owner.txtName.text)
        Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
        HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                               "resolvedSheet=" & HistoryLoadDebug_SheetName(wsTarget), _
                               "nameVal=" & HistoryLoadDebug_Quote(nameVal), _
                               "idVal=" & HistoryLoadDebug_Quote(idVal), _
                               "kanaVal=" & HistoryLoadDebug_Quote(kanaVal)

        If Len(idVal) > 0 Then
            validRow = FindLatestValidEvalRowByIdentity(wsTarget, nameVal, idVal, kanaVal)
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "identityLookupCalled=True", _
                                   "identityRow=" & CStr(validRow)
        End If
        If validRow = 0 Then
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "fallbackFindLatestRowByName=True"
            validRow = FindLatestRowByName(wsTarget, nameVal)
        Else
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "fallbackFindLatestRowByName=False"
        End If
        HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                               "finalValidRow=" & CStr(validRow)
        
        If validRow = 0 Then
             HistoryLoadDebug_ScanWorkbookForName nameVal, wsTarget
            MsgBox "対象の評価履歴が見つかりません。", vbInformation
            Exit Sub
        End If
        LoadAllSectionsFromSheet wsTarget, validRow, owner
        Exit Sub

    End If

    If Len(resolveMessage) > 0 Then
        MsgBox resolveMessage, vbExclamation
    End If
    ' ★ここまで

End Sub


' 下から遡って氏名一致の最新行を返す（見出しは「氏名」「利用者名」「名前」を順に探す）
Public Function FindLatestRowByName(ws As Worksheet, nameText As String) As Long

    Dim c As Long
    c = FindHeaderCol(ws, "氏名")
    If c = 0 Then c = FindHeaderCol(ws, "利用者名")
    If c = 0 Then c = FindHeaderCol(ws, "名前")
    If c = 0 Then
        HistoryLoadDebug_Print "[FindLatestRowByName]", _
                               "sheet=" & HistoryLoadDebug_SheetName(ws), _
                               "targetName=" & HistoryLoadDebug_Quote(nameText), _
                               "nameHeaderMissing=True"
        Exit Function
    End If

    HistoryLoadDebug_Print "[FindLatestRowByName]", _
                           "sheet=" & HistoryLoadDebug_SheetName(ws), _
                           "targetName=" & HistoryLoadDebug_Quote(nameText), _
                           "nameHeader=" & HistoryLoadDebug_Quote(CStr(ws.Cells(1, c).value)), _
                           "nameCol=" & CStr(c)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, c).End(xlUp).row
    Dim r As Long
    Dim rowName As String
    Dim normalizedTarget As String
    Dim normalizedRow As String

    normalizedTarget = NormalizeName(nameText)
    HistoryLoadDebug_Print "[FindLatestRowByName]", "lastRow=" & CStr(lastRow)
    For r = lastRow To 2 Step -1      ' 1行目は見出し想定
        rowName = CStr(ws.Cells(r, c).value)
        normalizedRow = NormalizeName(rowName)
        HistoryLoadDebug_Print "[FindLatestRowByName][SCAN]", _
                               "row=" & CStr(r), _
                               "raw=" & HistoryLoadDebug_Quote(rowName), _
                               "normalized=" & HistoryLoadDebug_Quote(normalizedRow), _
                               "matched=" & CStr(normalizedRow = normalizedTarget)
        If normalizedRow = normalizedTarget Then
            HistoryLoadDebug_Print "[FindLatestRowByName]", _
                                   "matchedRow=" & CStr(r)
            FindLatestRowByName = r
            Exit Function
        End If
    Next r

    HistoryLoadDebug_Print "[FindLatestRowByName]", "matchedRow=0"
End Function



Public Function CountRowsByName(ws As Worksheet, nameText As String) As Long
    Dim c As Long
    c = FindHeaderCol(ws, "氏名")
    If c = 0 Then c = FindHeaderCol(ws, "利用者名")
    If c = 0 Then c = FindHeaderCol(ws, "名前")
    If c = 0 Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, c).End(xlUp).row

    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, c).value), nameText, vbTextCompare) = 0 Then
            CountRowsByName = CountRowsByName + 1
        End If
    Next r
End Function



Public Function FindLatestRowByNameAndID( _
        ws As Worksheet, _
        nameText As String, _
        idVal As String) As Long

    Dim cName As Long, cID As Long
    cName = FindHeaderCol(ws, "氏名")
    If cName = 0 Then cName = FindHeaderCol(ws, "利用者名")
    If cName = 0 Then cName = FindHeaderCol(ws, "名前")
    If cName = 0 Then Exit Function

    cID = FindColByHeaderExact(ws, "Basic.ID")
    If cID = 0 Then cID = FindColByHeaderExact(ws, "ID")
    If cID = 0 Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, cName).End(xlUp).row

    ' 下から探す＝最新優先
    For r = lastRow To 2 Step -1
        If StrComp(CStr(ws.Cells(r, cName).value), nameText, vbTextCompare) = 0 Then
            If StrComp(CStr(ws.Cells(r, cID).value), idVal, vbTextCompare) = 0 Then
                FindLatestRowByNameAndID = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function FindEvalIndexRowsByUserID(ByVal indexWs As Worksheet, ByVal userID As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    Dim r As Long

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(indexWs.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then c.Add r
    Next r

    Set FindEvalIndexRowsByUserID = c
End Function

Private Function BuildDuplicateUserIDMessage(ByVal indexWs As Worksheet, ByVal userID As String, ByVal rowsByID As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long

    For i = 1 To rowsByID.count
        rowNo = CLng(rowsByID(i))
        lines = lines & _
            "- Name: " & Trim$(CStr(indexWs.Cells(rowNo, 2).value)) & _
            " / Kana: " & Trim$(CStr(indexWs.Cells(rowNo, 3).value)) & _
            " / Sheet: " & Trim$(CStr(indexWs.Cells(rowNo, 4).value))
        If i < rowsByID.count Then lines = lines & vbCrLf
    Next i

    BuildDuplicateUserIDMessage = _
       "EvalIndex内で同一IDが複数存在しています。" & vbCrLf & _
       "ID: " & userID & vbCrLf & vbCrLf & lines
End Function

Private Function BuildUserIdentityMismatchMessage(ByVal userID As String, _
                                                  ByVal inputName As String, _
                                                  ByVal indexName As String, _
                                                  ByVal inputKana As String, _
                                                  ByVal indexKana As String) As String
    Dim lines As String

    lines = lines & "ID不一致エラー" & vbCrLf
    lines = lines & "ID: " & userID & vbCrLf
    lines = lines & "入力氏名: " & inputName & vbCrLf
    lines = lines & "登録氏名: " & indexName

    If Len(Trim$(inputKana)) > 0 Or Len(Trim$(indexKana)) > 0 Then
        lines = lines & vbCrLf & "入力カナ: " & inputKana & vbCrLf & "登録カナ: " & indexKana
    End If

    BuildUserIdentityMismatchMessage = lines
End Function

Private Function IsSameKanaIfAvailable(ByVal leftKana As String, ByVal rightKana As String) As Boolean
    leftKana = Trim$(leftKana)
    rightKana = Trim$(rightKana)

    If Len(leftKana) = 0 Or Len(rightKana) = 0 Then
        IsSameKanaIfAvailable = True
    Else
        IsSameKanaIfAvailable = (StrComp(leftKana, rightKana, vbTextCompare) = 0)
    End If
End Function

Private Function FindLatestValidEvalRowByIdentity(ByVal ws As Worksheet, _
                                                  ByVal nameText As String, _
                                                  ByVal idVal As String, _
                                                  Optional ByVal kanaText As String = "") As Long
    Dim cEval As Long: cEval = FindColByHeaderExact(ws, "Basic.EvalDate")
    Dim cID As Long: cID = FindColByHeaderExact(ws, "Basic.ID")
    Dim cName As Long
    Dim cKana As Long: cKana = FindColByHeaderExact(ws, "Basic.NameKana")
    Dim lastRow As Long
    Dim r As Long
    Dim d As Date
    Dim bestDate As Date
    Dim bestRow As Long
    Dim rowName As String
    Dim rowKana As String

    If cEval = 0 Then Exit Function

    If cID = 0 Then cID = FindColByHeaderExact(ws, "ID")
    If cID = 0 Then Exit Function

    cName = FindColByHeaderExact(ws, "Basic.Name")
      If cName = 0 Then cName = FindHeaderCol(ws, "氏名")
      If cName = 0 Then cName = FindHeaderCol(ws, "利用者名")
      If cName = 0 Then cName = FindHeaderCol(ws, "Name")
    If cName = 0 Then Exit Function

    lastRow = LastDataRow(ws)

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, cID).value)), Trim$(idVal), vbTextCompare) <> 0 Then GoTo NextRow

        rowName = Trim$(CStr(ws.Cells(r, cName).value))
        If NormalizeName(rowName) <> NormalizeName(nameText) Then GoTo NextRow

        If Len(Trim$(kanaText)) > 0 And cKana > 0 Then
            rowKana = Trim$(CStr(ws.Cells(r, cKana).value))
            If Len(rowKana) > 0 Then
                If StrComp(rowKana, Trim$(kanaText), vbTextCompare) <> 0 Then GoTo NextRow
            End If
        End If

        If Not TryParseEvalDate(ws.Cells(r, cEval).value, d) Then GoTo NextRow

        If bestRow = 0 Then
            bestRow = r
            bestDate = d
        ElseIf d > bestDate Then
            bestRow = r
            bestDate = d
        ElseIf d = bestDate Then
            If r > bestRow Then bestRow = r
        End If
NextRow:
    Next r

    FindLatestValidEvalRowByIdentity = bestRow
End Function


'======================== 補助：フォーム／シート／行 ========================

Private Sub EnsureFormLoaded()
    On Error Resume Next
    Dim t$: t = frmEval.caption            ' 参照できればロード済み
    If Err.Number <> 0 Then Load frmEval
    On Error GoTo 0
    If frmEval.Visible = False Then frmEval.Show vbModeless   ' モデルレスで操作可
End Sub

Private Function EnsureEvalSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureEvalSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureEvalSheet Is Nothing Then
        Set EnsureEvalSheet = ThisWorkbook.Worksheets.Add(after:=Sheets(Sheets.count))
        On Error Resume Next
        EnsureEvalSheet.name = sheetName   ' 既存名ならExcelが自動リネーム
        On Error GoTo 0
    End If
End Function



Private Function LastDataRow(ws As Worksheet) As Long
    On Error Resume Next
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    LastDataRow = IIf(f Is Nothing, 1, f.row)
End Function

Private Function NextAppendRow(ws As Worksheet) As Long
    Dim lr As Long: lr = LastDataRow(ws)
    NextAppendRow = IIf(lr < 2, 2, lr + 1)
End Function

'====================================================================
' [ENTRY] 評価保存の正規入口
'  - UI 側（frmEval や他フォーム）は原則ここだけを呼び出す
'  - 行の決定（Append 行）はこの中で NextAppendRow により一元管理
'  - SaveAllSectionsToSheet / SaveBasicInfoToSheet_FromMe 等の下位関数を
'    直接他モジュールから呼ばないこと（スキーマ変更時の漏れ防止）
'====================================================================



Public Sub SaveEvaluation_Append_From(owner As Object)
    Dim wsUser As Worksheet
    Dim resolveMessage As String


    If ResolveUserHistorySheet(owner, True, wsUser, resolveMessage) Then
        EnsureHistorySheetInitialized wsUser
        EnsureClientMasterEntry owner
        
        Dim appendRow As Long
        appendRow = NextAppendRow(wsUser)
        
        wsUser.Cells(appendRow, EnsureHeader(wsUser, HDR_ROWNO)).value = appendRow - 1
        SaveAllSectionsToSheet wsUser, appendRow, owner
        Save_CognitionMental_AtRow wsUser, appendRow, owner
        MirrorBasicRow wsUser, appendRow
        
        Dim idxRow As Long
        idxRow = FindEvalIndexRowBySheetName(EnsureEvalIndexSheet(), wsUser.name)
        If idxRow > 0 Then
            UpdateEvalIndexMetadata owner, idxRow, wsUser.name
            UpdateEvalIndexStats idxRow, wsUser
        End If
        Exit Sub
    End If



    If Len(resolveMessage) > 0 Then
        MsgBox resolveMessage, vbExclamation
    Else
        MsgBox "保存先シートが見つからないため、保存を中断します。", vbExclamation
    End If
    
End Sub

Private Function ClientMasterHeaders() As Variant
    ClientMasterHeaders = Array(HDR_USER_ID, HDR_NAME, HDR_KANA, HDR_BIRTH_DATE, HDR_GENDER, HDR_CARE_LEVEL, HDR_CREATED_DATE)
End Function

Private Function EnsureClientMasterSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CLIENT_MASTER_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=Sheets(Sheets.count))
        ws.name = CLIENT_MASTER_SHEET_NAME
    End If

    Dim headers As Variant: headers = ClientMasterHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = CStr(headers(i))
    Next i

    Set EnsureClientMasterSheet = ws
End Function

Private Function FindClientMasterRowsByName(ByVal ws As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(NormalizeName(CStr(ws.Cells(r, 2).value)), NormalizeName(nameText), vbTextCompare) = 0 Then c.Add r
    Next r
    Set FindClientMasterRowsByName = c
End Function

Private Function FindClientMasterRowByUserID(ByVal ws As Worksheet, ByVal userID As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then
            FindClientMasterRowByUserID = r
            Exit Function
        End If
    Next r
End Function

Private Function FindClientMasterRow(ByVal ws As Worksheet, ByVal userID As String, ByVal nameText As String, ByRef shouldSkip As Boolean) As Long
    Dim rowsByName As Collection

    If Len(Trim$(userID)) > 0 Then
        FindClientMasterRow = FindClientMasterRowByUserID(ws, userID)
        Exit Function
    End If

    If Len(Trim$(nameText)) = 0 Then Exit Function

    Set rowsByName = FindClientMasterRowsByName(ws, nameText)
    If rowsByName.count = 1 Then
        FindClientMasterRow = CLng(rowsByName(1))
    ElseIf rowsByName.count > 1 Then
        shouldSkip = True
    End If
End Function

Private Function TryGetBirthDateForClientMaster(ByVal owner As Object, ByRef outDateText As String) As Boolean
    On Error GoTo EH

    Dim rawBirth As String
    rawBirth = Trim$(GetCtlTextGeneric(owner, "txtBirth"))
    If Len(rawBirth) = 0 Then Exit Function

    Dim dtBirth As Date
    If CallByName(owner, "TryGetBirthDateForStorage", VbMethod, rawBirth, dtBirth) Then
        outDateText = Format$(dtBirth, "yyyy/mm/dd")
        TryGetBirthDateForClientMaster = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Private Sub EnsureClientMasterEntry(ByVal owner As Object)
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = EnsureClientMasterSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim nameVal As String: nameVal = Trim$(GetCtlTextGeneric(owner, "txtName"))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    Dim genderVal As String: genderVal = Trim$(GetCtlTextGeneric(owner, "cboSex"))
    Dim careVal As String: careVal = Trim$(GetCtlTextGeneric(owner, "cboCare"))

    Dim skipRegistration As Boolean
    Dim hitRow As Long
    hitRow = FindClientMasterRow(ws, idVal, nameVal, skipRegistration)
    If hitRow > 0 Then Exit Sub
    If skipRegistration Then Exit Sub
    If Len(nameVal) = 0 Then Exit Sub

    Dim birthText As String
    Call TryGetBirthDateForClientMaster(owner, birthText)

    Dim newRow As Long
    newRow = NextAppendRow(ws)

    ws.Cells(newRow, 1).value = idVal
    ws.Cells(newRow, 2).value = nameVal
    ws.Cells(newRow, 3).value = kanaVal
    ws.Cells(newRow, 4).value = birthText
    ws.Cells(newRow, 5).value = genderVal
    ws.Cells(newRow, 6).value = careVal
    ws.Cells(newRow, 7).value = Format$(Date, "yyyy/mm/dd")
    Exit Sub
EH:
    Err.Clear
End Sub


Private Function GetSparseMainSaveWarningMessage(ws As Worksheet, ByVal patientName As String, owner As Object) As String
    Dim existingRow As Long
    existingRow = ResolveExistingEvalRow(ws, patientName, owner)
    If existingRow <= 0 Then Exit Function

    Dim totalCount As Long, blankCount As Long
    CountMainFormTextInputs owner, totalCount, blankCount
    

    Dim filledCount As Long
    filledCount = CountMainFormFilledFields(owner)

    Dim changeCount As Long
    changeCount = CountMainFormTextboxChanges(ws, existingRow, owner)

    If filledCount < MAIN_SAVE_MIN_FILLED_FIELDS Then
        GetSparseMainSaveWarningMessage = MAIN_SAVE_FEW_INPUT_MESSAGE
        Exit Function
    End If

    If changeCount < MAIN_SAVE_MIN_CHANGE_COUNT Then
        GetSparseMainSaveWarningMessage = MAIN_SAVE_FEW_CHANGE_MESSAGE
    End If
End Function

Private Function ResolveExistingEvalRow(ws As Worksheet, ByVal patientName As String, owner As Object) As Long
    ResolveExistingEvalRow = FindLatestRowByName(ws, patientName)

    Dim idVal As String
    idVal = Trim$(GetID_FromBasicInfo(owner))
    If Len(idVal) = 0 Then Exit Function

    Dim rowByID As Long
    rowByID = FindLatestRowByNameAndID(ws, patientName, idVal)
    If rowByID > 0 Then ResolveExistingEvalRow = rowByID
End Function

Private Function CountMainFormTextboxChanges(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Long
    Dim map As Variant
    map = MainSaveTextboxHeaderMap()

    Dim i As Long
    For i = LBound(map) To UBound(map)
        Dim headerName As String
        Dim ctlName As String
        Dim c As Long
        Dim curVal As String
        Dim oldVal As String

        headerName = CStr(map(i)(0))
        ctlName = CStr(map(i)(1))

        c = FindColByHeaderExact(ws, headerName)
        If c = 0 Then GoTo NextItem

        curVal = NormalizeCompareValue(GetCtlTextGeneric(owner, ctlName))
        oldVal = NormalizeCompareValue(CStr(ws.Cells(existingRow, c).value))

        If StrComp(curVal, oldVal, vbBinaryCompare) <> 0 Then
            CountMainFormTextboxChanges = CountMainFormTextboxChanges + 1
        End If
NextItem:
    Next i
End Function

Private Function CountMainFormFilledFields(owner As Object) As Long
    Dim map As Variant
    map = MainSaveTextboxHeaderMap()

    Dim i As Long
    For i = LBound(map) To UBound(map)
        If Len(NormalizeCompareValue(GetCtlTextGeneric(owner, CStr(map(i)(1))))) > 0 Then
            CountMainFormFilledFields = CountMainFormFilledFields + 1
        End If
    Next i
End Function


Private Function MainSaveTextboxHeaderMap() As Variant
    MainSaveTextboxHeaderMap = Array( _
        Array("評価日", "txtEDate"), _
        Array("年齢", "txtAge"), _
        Array("生年月日", "txtBirth"), _
        Array("Basic.Name", "txtName"), _
        Array("評価者", "txtEvaluator"), _
        Array("評価者職種", "txtEvaluatorJob"), _
        Array("発症日", "txtOnset"), _
        Array("患者Needs", "txtNeedsPt"), _
        Array("家族Needs", "txtNeedsFam"), _
        Array("生活状況", "txtLiving"), _
        Array("住宅備考", "txtBIHomeEnvNote"), _
        Array("主診断", "txtDx"), _
        Array("直近入院日", "txtAdmDate"), _
        Array("直近退院日", "txtDisDate"), _
        Array("治療経過", "txtTxCourse"), _
        Array("合併疾患", "txtComplications"), _
        Array("IO_Cog_DementiaNote", "txtDementiaNote"), _
        Array("IO_Mental_Note", "txtMentalNote") _
    )
End Function

Private Function NormalizeCompareValue(ByVal v As String) As String
    NormalizeCompareValue = Trim$(Replace(CStr(v), vbCrLf, vbLf))
End Function

Private Function ResolveDailyLogRoot(ByVal owner As Object) As Object
    If owner Is Nothing Then Exit Function
    Set ResolveDailyLogRoot = SafeGetControl(owner, "fraDailyLog")
End Function

Private Function ResolveDailyLogControl(ByVal owner As Object, ByVal controlName As String) As Object
    Dim root As Object
    If owner Is Nothing Then Exit Function
    If LenB(Trim$(controlName)) = 0 Then Exit Function

    Set root = ResolveDailyLogRoot(owner)
    If root Is Nothing Then Exit Function

    Set ResolveDailyLogControl = SafeGetControl(root, controlName)
End Function

Private Sub CountMainFormTextInputs(owner As Object, ByRef totalCount As Long, ByRef blankCount As Long)
    Dim dailyRoot As Object
    On Error Resume Next
    Set dailyRoot = ResolveDailyLogRoot(owner)
    On Error GoTo 0

    CountTextInputsRecursive owner, dailyRoot, totalCount, blankCount
End Sub

Private Sub CountTextInputsRecursive(ByVal container As Object, ByVal excludedRoot As Object, ByRef totalCount As Long, ByRef blankCount As Long)
    Dim ctrl As Object
    For Each ctrl In container.controls
        If Not excludedRoot Is Nothing Then
            If IsDescendantControl(ctrl, excludedRoot) Then GoTo NextControl
        End If

        Select Case TypeName(ctrl)
            Case "TextBox", "ComboBox"
                totalCount = totalCount + 1
                On Error Resume Next
                If Len(Trim$(CStr(ctrl.value))) = 0 Then blankCount = blankCount + 1
                On Error GoTo 0
        End Select

        On Error Resume Next
        Dim childCount As Long
        childCount = ctrl.controls.count
        If Err.Number = 0 And childCount > 0 Then
            On Error GoTo 0
            CountTextInputsRecursive ctrl, excludedRoot, totalCount, blankCount
        Else
            Err.Clear
            On Error GoTo 0
        End If
NextControl:
    Next ctrl
End Sub

Private Function IsDescendantControl(ByVal ctrl As Object, ByVal root As Object) As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = ctrl
    Do While Not p Is Nothing
        If p Is root Then
            IsDescendantControl = True
            Exit Function
        End If
        Set p = p.parent
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop
    On Error GoTo 0
End Function




Private Sub LoadEvaluation_LastRow_From_OBSOLETE(owner As Object)
    MsgBox "この入口は廃止しました。読み込みは『名前→直近候補から選択』に統一しています。", vbInformation
End Sub




' ====== 基本情報の保存/読込（このモジュール内） ======

' 見出しの列を取得（無ければ新規作成）
Private Function EnsureHeaderCol(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Then
        EnsureHeaderCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + IIf(ws.Cells(1, 1).value <> "", 1, 0)
        If EnsureHeaderCol = 0 Then EnsureHeaderCol = 1
        ws.Cells(1, EnsureHeaderCol).value = header
    Else
        EnsureHeaderCol = f.Column
    End If
End Function

' 見出しの列を探す（無ければ 0）
Private Function FindHeaderCol(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Then FindHeaderCol = 0 Else FindHeaderCol = f.Column
End Function

Private Function FindHeaderColAny(ws As Worksheet, headers As Variant) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        FindHeaderColAny = FindHeaderCol(ws, CStr(headers(i)))
        If FindHeaderColAny > 0 Then Exit Function
    Next i
End Function


Private Sub SetCtlValueSafe(owner As Object, ctlName As String, ByVal v As Variant)
    Dim o As Object
    Set o = FindCtlDeep(owner, ctlName)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.value = v
    On Error GoTo 0
End Sub

Private Function HomeEnvControlNames() As Variant
    HomeEnvControlNames = Array( _
        "chkBIHomeEnv_Entrance", _
        "chkBIHomeEnv_Genkan", _
        "chkBIHomeEnv_IndoorStep", _
        "chkBIHomeEnv_Stairs", _
        "chkBIHomeEnv_Handrail", _
        "chkBIHomeEnv_Slope", _
        "chkBIHomeEnv_NarrowPath" _
    )
End Function

Private Function SerializeNamedChecks(owner As Object, checkNames As Variant) As String
    Dim i As Long, o As Object, s As String
    For i = LBound(checkNames) To UBound(checkNames)
        Set o = FindCtlDeep(owner, CStr(checkNames(i)))
        If Not o Is Nothing Then
            On Error Resume Next
            If CBool(o.value) Then
                If Len(s) > 0 Then s = s & ","
                s = s & CStr(checkNames(i))
            End If
            On Error GoTo 0
        End If
    Next i
    SerializeNamedChecks = s
End Function

Private Sub DeserializeNamedChecks(owner As Object, checkNames As Variant, ByVal csv As String)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim p As Variant, i As Long, o As Object
    If Len(Trim$(csv)) > 0 Then
        For Each p In Split(csv, ",")
            dict(Trim$(CStr(p))) = True
        Next
    End If

    For i = LBound(checkNames) To UBound(checkNames)
        Set o = FindCtlDeep(owner, CStr(checkNames(i)))
        If Not o Is Nothing Then
            On Error Resume Next
            o.value = dict.exists(CStr(checkNames(i)))
            On Error GoTo 0
        End If
    Next i
End Sub


' 汎用：テキスト値を取得（TextBox/ComboBox/Labelなどに対応）
Private Function GetCtlTextGeneric(owner As Object, ctlName As String) As String
    Dim c As Object
    Set c = FindCtlDeep(owner, ctlName)
    If c Is Nothing Then Exit Function
    
    On Error Resume Next
    GetCtlTextGeneric = CStr(c.value)
End Function

Private Function GetHdrKanaText(owner As Object) As String
    Dim c As Object

    On Error Resume Next
    Set c = owner.controls("frHeader").controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.controls("txtHdrKana")
        On Error GoTo 0
    End If

    If c Is Nothing Then
        GetHdrKanaText = ""
    Else
        On Error Resume Next
        GetHdrKanaText = Trim$(CStr(c.value))
        On Error GoTo 0
    End If
End Function

Private Sub SetHdrKanaText(owner As Object, ByVal v As Variant)
    Dim c As Object

    On Error Resume Next
    Set c = owner.controls("frHeader").controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.controls("txtHdrKana")
        On Error GoTo 0
    End If

    If c Is Nothing Then Exit Sub

    On Error Resume Next
    c.value = CStr(v)
    On Error GoTo 0
End Sub

' 汎用：コンボを安全にセット（リストにある時だけ選択）
Private Sub SetComboSafe_Basic(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cb As MSForms.ComboBox
    Dim s As String, i As Long, hit As Long
    s = CStr(v)
    Set cb = FindCtlDeep(owner, ctlName)
    If cb Is Nothing Then Exit Sub
    hit = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i)) = s Then hit = i: Exit For
    Next
    If hit >= 0 Then cb.ListIndex = hit Else cb.ListIndex = -1
End Sub

Private Sub SyncAgeBeforeBasicSave(ByVal owner As Object)
    On Error GoTo EH

    ' frmEval ?J??N\bh??pAXR[vOQ?
    CallByName owner, "SyncAgeFromBirth", VbMethod
    Exit Sub
EH:
    Debug.Print "[Basic] SyncAgeBeforeBasicSave skipped:", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub WriteBirthTextCell(ByVal target As Range, ByVal birthText As String)
    On Error Resume Next
    target.NumberFormat = "@"
    On Error GoTo 0
    target.Value2 = CStr(birthText)
End Sub

Private Function ReadBirthTextCell(ByVal target As Range) As String
    Dim s As String
    s = CStr(target.text)
    If Len(s) = 0 Then s = CStr(target.Value2)
    If Len(s) > 0 And Left$(s, 1) = "'" Then s = Mid$(s, 2)
    ReadBirthTextCell = s
End Function

'====================================================================
' BasicInfo IO セクション（評価日・氏名・年齢・Needs 等）
'  - EvalData 上の Basic.* 系ヘッダとの対応を一元管理する窓口
'  - 新しい Basic 項目を追加する場合は、原則ここにマッピングを足す
'  - 列の別名統合やスキーマ統一は EnsureHeaderCol_BasicInfo 側で行う
'  - 他のモジュールからは、Basic.* の物理列を直接触らず、
'    必要なら GetID_FromBasicInfo / GetBasicInfoFrame などのヘルパを経由する
'====================================================================




' --- 保存 ---
Public Sub SaveBasicInfoToSheet_FromMe(ws As Worksheet, r As Long, owner As Object)
    
    Debug.Print "[Basic] Enter_SaveBasicInfo | ws=" & ws.name & " | r=" & r

    SyncAgeBeforeBasicSave owner
    
    
    '--- 単一値のマッピング（最後の要素に _ を付けない） ---
    Dim map As Variant
map = Array( _
    Array("評価日", "txtEDate"), _
    Array("年齢", "txtAge"), _
    Array("生年月日", "txtBirth"), _
    Array("性別", "cboSex"), _
    Array("Basic.Name", "txtName"), _
    Array("評価者", "txtEvaluator"), _
    Array("評価者職種", "txtEvaluatorJob"), _
    Array("発症日", "txtOnset"), _
    Array("患者Needs", "txtNeedsPt"), _
    Array("家族Needs", "txtNeedsFam"), _
    Array("生活状況", "txtLiving"), _
    Array("住宅備考", "txtBIHomeEnvNote"), _
    Array("主診断", "txtDx"), _
    Array("要介護度", "cboCare"), _
    Array("障害高齢者の日常生活自立度", "cboElder"), _
    Array("認知症高齢者の日常生活自立度", "cboDementia"), _
    Array("直近入院日", "txtAdmDate"), _
    Array("直近退院日", "txtDisDate"), _
    Array("治療経過", "txtTxCourse"), _
    Array("合併疾患", "txtComplications") _
)

    Call EnsureHeaderCol(ws, "N")

    '--- 既存のループ：単一値を書き込み ---
    Dim i As Long, head As String, ctl As String, c As Long, v As String
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0)):  ctl = CStr(map(i)(1))
        v = GetCtlTextGeneric(owner, ctl)
        If Len(v) > 0 Then
            c = FindColByHeaderExact(ws, head): If c = 0 Then c = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1: ws.Cells(1, c).value = head
            If ctl = "txtBirth" Then
                WriteBirthTextCell ws.Cells(r, c), v
            Else
                ws.Cells(r, c).value = v
            End If
            Debug.Print "[BASIC][SAVE]", head, "->", v
        End If
    Next i
    
    c = EnsureHeader(ws, "住宅状況")
    ws.Cells(r, c).value = SerializeNamedChecks(owner, HomeEnvControlNames())


    c = EnsureHeader(ws, "Basic.NameKana")
    ws.Cells(r, c).value = GetHdrKanaText(owner)
    Debug.Print "[BASIC][SAVE] Basic.NameKana ->", CStr(ws.Cells(r, c).value)
    
    Dim idVal As String: idVal = GetID_FromBasicInfo(owner)
    If Len(idVal) > 0 Then ws.Cells(r, EnsureHeader(ws, "Basic.ID")).value = idVal
    ws.Cells(r, EnsureHeader(ws, "Basic.EvalDate")).value = GetCtlTextGeneric(owner, "txtEDate")
    

    '--- ここから追記：チェック群のCSV保存（補助具／リスク）※ループの“後ろ” ---
    Dim s As String
    c = EnsureHeader(ws, "補助具")
s = SerializeChecks(owner, "Frame33", True)
Debug.Print "[BASIC][SAVE] 補助具 ->", s, " @col=", c
ws.Cells(r, c).value = s
c = EnsureHeader(ws, HDR_AIDS_CHECKS)
ws.Cells(r, c).value = s

   c = EnsureHeader(ws, "リスク")
s = SerializeChecks(owner, "Frame34", False)
Debug.Print "[BASIC][SAVE] リスク ->", s, " @col=", c
ws.Cells(r, c).value = s

c = EnsureHeader(ws, HDR_RISK_CHECKS)
ws.Cells(r, c).value = s

c = EnsureHeader(ws, HDR_HOMEENV_CHECKS)
ws.Cells(r, c).value = SerializeNamedChecks(owner, HomeEnvControlNames())

c = EnsureHeader(ws, HDR_HOMEENV_NOTE)
ws.Cells(r, c).value = GetCtlTextGeneric(owner, "txtBIHomeEnvNote")

    
    
    
    
End Sub




' --- 読込 ---
Public Sub LoadBasicInfoFromSheet_FromMe(ws As Worksheet, ByVal r As Long, owner As Object)

    On Error GoTo EH
    Debug.Print "[TRACE] Enter LoadBasicInfoFromSheet_FromMe r=" & r

    '--- 単一値のマッピング ---
    Dim map As Variant
map = Array( _
    Array("評価日", "txtEDate"), _
    Array("年齢", "txtAge"), _
    Array("生年月日", "txtBirth"), _
    Array("性別", "cboSex"), _
    Array("Basic.Name", "txtName"), _
    Array("評価者", "txtEvaluator"), _
    Array("評価者職種", "txtEvaluatorJob"), _
    Array("発症日", "txtOnset"), _
    Array("患者Needs", "txtNeedsPt"), _
    Array("家族Needs", "txtNeedsFam"), _
    Array("住宅備考", "txtBIHomeEnvNote"), _
    Array("生活状況", "txtLiving"), _
    Array("主診断", "txtDx"), _
    Array("要介護度", "cboCare"), _
    Array("障害高齢者の日常生活自立度", "cboElder"), _
    Array("認知症高齢者の日常生活自立度", "cboDementia"), _
    Array("直近入院日", "txtAdmDate"), _
    Array("直近退院日", "txtDisDate"), _
    Array("治療経過", "txtTxCourse"), _
    Array("合併疾患", "txtComplications") _
)



    '--- 単一値をフォームへ読込 ---
    Dim i As Long, head As String, ctl As String, c As Long, v As Variant
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0))
        ctl = CStr(map(i)(1))

        c = FindHeaderCol(ws, head)
        If c > 0 Then
            If ctl = "txtBirth" Then
                v = ReadBirthTextCell(ws.Cells(r, c))
            Else
                v = ws.Cells(r, c).value
            End If
            If Left$(ctl, 3) = "cbo" Then
                SetComboSafely owner, ctl, CStr(v)
            Else
                Dim o As Object: Set o = FindCtlDeep(owner, ctl)
                If Not o Is Nothing Then o.value = v
            End If
            
        End If
        Next i

    c = FindHeaderCol(ws, "住宅状況")
    If c > 0 Then DeserializeNamedChecks owner, HomeEnvControlNames(), CStr(ws.Cells(r, c).value)

    c = FindHeaderCol(ws, "Basic.NameKana")
    If c > 0 Then SetHdrKanaText owner, ws.Cells(r, c).value

    c = FindHeaderCol(ws, "Basic.NameKana")
    If c > 0 Then SetHdrKanaText owner, ws.Cells(r, c).value

    '--- チェック群の復元（補助具／リスク） ---
    Dim csv As String

    ' 補助具
c = FindHeaderCol(ws, "補助具")
If c > 0 Then
    csv = CStr(ws.Cells(r, c).value)
    DeserializeChecks owner, "Frame33", csv, True
End If

' リスク
c = FindHeaderCol(ws, "リスク")
If c > 0 Then
    csv = CStr(ws.Cells(r, c).value)
    DeserializeChecks owner, "Frame34", csv, False
End If

c = FindHeaderColAny(ws, Array(HDR_AIDS_CHECKS, "?"))
If c > 0 Then DeserializeChecks owner, "Frame33", CStr(ws.Cells(r, c).value), True

c = FindHeaderColAny(ws, Array(HDR_RISK_CHECKS, "XN"))
If c > 0 Then DeserializeChecks owner, "Frame34", CStr(ws.Cells(r, c).value), False

c = FindHeaderCol(ws, HDR_HOMEENV_CHECKS)
If c > 0 Then DeserializeNamedChecks owner, HomeEnvControlNames(), CStr(ws.Cells(r, c).value)

c = FindHeaderCol(ws, HDR_HOMEENV_NOTE)
If c > 0 Then SetCtlValueSafe owner, "txtBIHomeEnvNote", ws.Cells(r, c).value




If GetBool(owner, "chkLoadParalysis", True) Then Call IO_SafeRunLoad("LoadParalysisFromSheet", ws, r, owner)
If GetBool(owner, "chkLoadROM", True) Then Call IO_SafeRunLoad("LoadROMFromSheet", ws, r, owner)
Debug.Print "[TRACE] About to run POSTURE"
If GetBool(owner, "chkLoadPosture", True) Then Call IO_SafeRunLoad("LoadPostureFromSheet", ws, r, owner)
Debug.Print "[TRACE] Done POSTURE"

Debug.Print "[TRACE] About to run MMT"
If TypeName(owner) = "frmEval" Then
    owner.QueueMMTLoadAfterUI ws, r
Else
    Call MMT.LoadMMTFromSheet(ws, r, owner)
End If
Debug.Print "[TRACE] Done MMT"
ExitHere:
    Exit Sub

EH:
    Debug.Print "[ERR][LoadBasicInfo] Err=" & Err.Number & " Desc=" & Err.Description
    Resume ExitHere


End Sub


' EvalDataシート取得
Public Function GetEvalDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise 5, , "EvalData シートがありません。"
    Set GetEvalDataSheet = ws
End Function

' 見出しから列番号（完全一致）
Public Function FindColByHeaderExact(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindColByHeaderExact = c
            Exit Function
        End If
    Next c
End Function

' ID行を検索（無ければ末尾に作成してIDを入れる）
Public Function GetOrCreateRowByID(ByVal ws As Worksheet, ByVal idVal As String) As Long
    Dim idCol As Long: idCol = FindColByHeaderExact(ws, "Basic.ID")
    If idCol = 0 Then
        ' 旧来の命名ならここで作る
        idCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If
    If Len(idVal) = 0 Then Err.Raise 5, , "IDが空です。"

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, idCol).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, idCol).value) = idVal Then
            GetOrCreateRowByID = r
            Exit Function
        End If
    Next r
    ' 無ければ新規行
    r = lastRow + 1
    ws.Cells(r, idCol).value = idVal
    GetOrCreateRowByID = r
End Function





' ラベル「ID」の右にある TextBox から値を取得（コントロール名に依存しない）
Public Function GetID_FromBasicInfo(ByVal owner As Object) As String
    On Error Resume Next
    GetID_FromBasicInfo = Trim$(CStr(owner.controls("frHeader").controls("txtHdrPID").value))
    On Error GoTo 0
End Function


'================ Basic情報の共通ヘルパ ==================

Public Function GetBasicInfoFrame(ByVal owner As Object) As Object
    Dim f As MSForms.Frame
    Set f = FindFrameByCaptionDeep_(owner, "基本情報")
    If Not f Is Nothing Then
        Set GetBasicInfoFrame = f
    Else
        Set GetBasicInfoFrame = owner   ' フォールバック：直接オーナーを渡せるように
    End If
End Function

Public Function GetTextByLabelInFrame(ByVal frm As Object, ByVal labelCaption As String) As String
    ' null / 非Frame でも安全に抜ける
    If frm Is Nothing Then Exit Function
    On Error Resume Next
    Dim HasControls As Boolean
    HasControls = Not (frm.controls Is Nothing)
    On Error GoTo 0
    If Not HasControls Then Exit Function

    ' --- 以下は今のロジックそのまま ---
    Dim lb As Object, ctl As Object
    For Each ctl In frm.controls
        If TypeName(ctl) = "Label" Then
            If InStr(1, CStr(ctl.caption), labelCaption, vbTextCompare) > 0 Then
                Set lb = ctl: Exit For
            End If
        End If
    Next ctl
    If lb Is Nothing Then Exit Function

    Dim best As Object, bestScore As Double
    bestScore = 1E+20
    For Each ctl In frm.controls
        If TypeName(ctl) = "TextBox" Then
            Dim dy As Double: dy = Abs((ctl.Top + ctl.Height / 2) - (lb.Top + lb.Height / 2))
            If dy <= lb.Height Then
                Dim dx As Double: dx = ctl.Left - lb.Left
                If dx > -5 Then
                    Dim sc As Double: sc = dy * 10 + Abs(dx)
                    If sc < bestScore Then Set best = ctl: bestScore = sc
                End If
            End If
        End If
    Next ctl
    If Not best Is Nothing Then GetTextByLabelInFrame = CStr(best.value)
End Function


' Frame を Caption 部分一致で深さ優先探索（UserForm / Frame / MultiPage 対応）
Public Function FindFrameByCaptionDeep_(ByVal owner As Object, ByVal captionLike As String) As MSForms.Frame
    Set FindFrameByCaptionDeep_ = FindFrameByCaptionDeep_Walk(owner, captionLike)
End Function

Private Function FindFrameByCaptionDeep_Walk(ByVal container As Object, ByVal captionLike As String) As MSForms.Frame
    On Error Resume Next

    If TypeName(container) = "MultiPage" Then
        Dim pg As Object
        For Each pg In container.Pages
            Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(pg, captionLike)
            If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function
        Next pg
    End If

    Dim tmp As Object: Set tmp = container.controls
    If Err.Number <> 0 Then Err.Clear: Exit Function

    Dim ctl As Object
    For Each ctl In container.controls
        Select Case TypeName(ctl)
            Case "Frame"
                If InStr(1, CStr(ctl.caption), captionLike, vbTextCompare) > 0 Then
                    Set FindFrameByCaptionDeep_Walk = ctl: Exit Function
                End If
                Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function

            Case "MultiPage"
                Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function

            Case Else
                Err.Clear
                Set tmp = ctl.controls
                If Err.Number = 0 Then
                    Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                    If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function
                Else
                    Err.Clear
                End If
        End Select
    Next ctl
End Function
'================ ここまで貼る ==================









' ==== BasicInfo の列名を Basic.* に統一し、不足は作る（安全マージ付き） ====
Public Sub EnsureHeaderCol_BasicInfo(ByVal ws As Worksheet)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    ' --- 単項目（主にテキスト/コンボ） ---
    d("BasicInfo_ID") = "Basic.ID":                  d("ID") = "Basic.ID": d("Pid") = "Basic.ID"
    d("BasicInfo_氏名") = "Basic.Name":              d("氏名") = "Basic.Name": d("Name") = "Basic.Name"
    d("BasicInfo_評価日") = "Basic.EvalDate":        d("評価日") = "Basic.EvalDate": d("EvalDate") = "Basic.EvalDate"
    d("BasicInfo_評価者") = "Basic.Evaluator":       d("評価者") = "Basic.Evaluator"
    d("BasicInfo_年齢") = "Basic.Age":               d("年齢") = "Basic.Age": d("Age") = "Basic.Age"
    d("BasicInfo_評価者職種") = "Basic.EvaluatorJob": d("評価者職種") = "Basic.EvaluatorJob": d("EvaluatorJob") = "Basic.EvaluatorJob"
    d("BasicInfo_性別") = "Basic.Sex":               d("性別") = "Basic.Sex": d("Sex") = "Basic.Sex"
    d("BasicInfo_主診断") = "Basic.PrimaryDx":       d("主診断") = "Basic.PrimaryDx": d("主病名") = "Basic.PrimaryDx"
    d("BasicInfo_発症日") = "Basic.OnsetDate":       d("発症日") = "Basic.OnsetDate"
    d("BasicInfo_要介護度") = "Basic.CareLevel":     d("要介護度") = "Basic.CareLevel"
    d("BasicInfo_認知症自立度") = "Basic.DementiaADL"
    d("BasicInfo_認知症高齢者の日常生活自立度") = "Basic.DementiaADL"
    d("認知症高齢者の日常生活自立度") = "Basic.DementiaADL"
    d("BasicInfo_生活状況") = "Basic.LifeStatus":    d("生活状況") = "Basic.LifeStatus"
    d("BasicInfo_患者Needs") = "Basic.Needs.Patient": d("患者Needs") = "Basic.Needs.Patient"
    d("BasicInfo_家族Needs") = "Basic.Needs.Family":  d("家族Needs") = "Basic.Needs.Family"

    ' --- 補助具（チェック）→ Basic.Aids.* へ ---
    AddAlias d, "BasicInfo_補助具_杖", "Basic.Aids.杖"
    AddAlias d, "BasicInfo_補助具_歩行器", "Basic.Aids.歩行器"
    AddAlias d, "BasicInfo_補助具_短下肢装具", "Basic.Aids.短下肢装具"
    AddAlias d, "BasicInfo_補助具_手すり", "Basic.Aids.手すり"
    AddAlias d, "BasicInfo_補助具_シルバーカー", "Basic.Aids.シルバーカー"
    AddAlias d, "BasicInfo_補助具_車いす", "Basic.Aids.車いす": AddAlias d, "BasicInfo_補助具_車椅子", "Basic.Aids.車いす"
    AddAlias d, "BasicInfo_補助具_介助ベルト", "Basic.Aids.介助ベルト"
    AddAlias d, "BasicInfo_補助具_スロープ", "Basic.Aids.スロープ"

    ' --- リスク（チェック）→ Basic.Risk.* へ ---
    AddAlias d, "BasicInfo_リスク_転倒", "Basic.Risk.転倒"
    AddAlias d, "BasicInfo_リスク_窒息", "Basic.Risk.窒息"
    AddAlias d, "BasicInfo_リスク_低栄養", "Basic.Risk.低栄養"
    AddAlias d, "BasicInfo_リスク_せん妄", "Basic.Risk.せん妄"
    AddAlias d, "BasicInfo_リスク_誤嚥", "Basic.Risk.誤嚥"
    AddAlias d, "BasicInfo_リスク_褥瘡", "Basic.Risk.褥瘡"
    AddAlias d, "BasicInfo_リスク_ADL低下", "Basic.Risk.ADL低下"
    AddAlias d, "Basic.Aids.Checks", "Basic.Aids.Checks"
    AddAlias d, "Basic.Risk.Checks", "Basic.Risk.Checks"
    AddAlias d, "BasicInfo_BI.HomeEnv.Note", "Basic.HomeEnv.Note"
    
    

    ' 1) 既存ヘッダをマージ改名
    ApplyAliasesMerge_Basic ws, d

    ' 2) 最低限必要な列がなければ追加（Save/Loadの対象を漏れなく）
    Dim need As Variant, mustHave As Variant
    mustHave = Array( _
        "Basic.ID", "Basic.Name", "Basic.EvalDate", "Basic.Evaluator", _
        "BI.EvaluatorJob", _
        "Basic.Age", "Basic.Sex", "Basic.PrimaryDx", "Basic.OnsetDate", _
        "Basic.CareLevel", "Basic.DementiaADL", "Basic.LifeStatus", _
        "Basic.Needs.Patient", "Basic.Needs.Family", _
        "Basic.Medical.AdmitDate", "Basic.Medical.DischargeDate", _
        "Basic.Medical.CourseNote", "Basic.Medical.ComplicationNote", _
        "Basic.HomeEnv.Checks", "Basic.HomeEnv.Note", _
        "Basic.Aids.Checks", "Basic.Risk.Checks" _
    )
    For Each need In mustHave
        EnsureHeaderExists ws, CStr(need)
    Next need
End Sub

' === ヘルパー ===
Private Sub AddAlias(ByVal d As Object, ByVal src As String, ByVal dst As String)
    d(src) = dst
End Sub

' エイリアス改名（衝突時はマージして旧列を削除）
Private Sub ApplyAliasesMerge_Basic(ByVal ws As Worksheet, ByVal d As Object)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long
    For j = lastCol To 1 Step -1
        Dim h As String: h = Trim$(CStr(ws.Cells(1, j).value))
        If Len(h) = 0 Then GoTo NextJ
        If d.exists(h) Then
            Dim dst As String: dst = CStr(d(h))
            Dim dstCol As Long: dstCol = modSchema.FindColByHeaderExact(ws, dst)
            If dstCol > 0 And dstCol <> j Then
                ' マージ（空欄だけ埋める）
                Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, j).End(xlUp).row
                Dim r As Long
                For r = 2 To lastRow
                    If Len(ws.Cells(r, dstCol).value) = 0 And Len(ws.Cells(r, j).value) > 0 Then
                        ws.Cells(r, dstCol).value = ws.Cells(r, j).value
                    End If
                Next r
                ws.Columns(j).Delete
            Else
                ws.Cells(1, j).value = dst
            End If
        End If
NextJ:
    Next j
End Sub

Private Sub EnsureHeaderExists(ByVal ws As Worksheet, ByVal hdr As String)
    If modSchema.FindColByHeaderExact(ws, hdr) = 0 Then
        Dim lc As Long: lc = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        ws.Cells(1, lc + 1).value = hdr
    End If
End Sub










' EvalDataのID行を見つける（無ければ作る）
' 既存スキーマのどちらにも対応：Basic.ID / BasicInfo_ID
Public Function GetOrCreateRowByID_Basic(ByVal ws As Worksheet, ByVal idVal As String) As Long
    If Len(idVal) = 0 Then Err.Raise 5, , "IDが空です。"

    Dim idCol As Long
    idCol = FindColByHeaderExact(ws, "Basic.ID")
    If idCol = 0 Then idCol = FindColByHeaderExact(ws, "BasicInfo_ID")
    If idCol = 0 Then
        ' 無ければ Basic.ID を作る（既存に合わせてOK・後でスキーマ統一可）
        idCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, idCol).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, idCol).value) = idVal Then GetOrCreateRowByID_Basic = r: Exit Function
    Next r

    r = lastRow + 1
    ws.Cells(r, idCol).value = idVal
    GetOrCreateRowByID_Basic = r
End Function











'--- コンボボックスに安全に値を反映（一覧に無い値なら未選択にする） ---
Private Sub SetComboSafely(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cb As Object  ' MSForms.ComboBox を late binding で扱う
    Dim i As Long, hit As Long
    Dim s As String

    On Error Resume Next
    Set cb = FindCtlDeep(owner, ctlName)
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub

    s = CStr(v)
    hit = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i)) = s Then
            hit = i
            Exit For
        End If
    Next

    If hit >= 0 Then
        cb.ListIndex = hit               ' 一致が見つかったら選択
    Else
        cb.ListIndex = -1                ' 見つからなければ未選択に（DropDownListでも安全）
        ' ※DropDownListの場合、cb.Text には入れません
    End If
End Sub











Private Function FindControlDeep(ByVal parent As Object, ByVal targetName As String) As Object
    Dim c As Object, hit As Object

    ' 1) 自分自身が一致なら即返す
    On Error Resume Next
    If Not parent Is Nothing Then
        If parent.name = targetName Then Set FindControlDeep = parent: Exit Function
    End If
    On Error GoTo 0

    ' 2) MultiPage は Pages を走査
    If TypeName(parent) = "MultiPage" Then
        Dim pg As Object
        For Each pg In parent.Pages
            Set hit = FindControlDeep(pg, targetName)
            If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function
        Next pg
        Exit Function
    End If

    ' 3) 直下に同名があれば取得（存在しない型でも例外にしない）
    On Error Resume Next
    Set hit = parent.controls(targetName)
    On Error GoTo 0
    If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function

    ' 4) 子コントロールを再帰走査（Controls を持たない型はスキップ）
    On Error Resume Next
    For Each c In parent.controls
        Err.Clear
        Set hit = FindControlDeep(c, targetName)
        If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function
    Next c
    On Error GoTo 0
End Function


' 代表キャプションから親フレームを推定
Private Function FindGroupByAnyCaption(frm As Object, captions As Variant) As Object
    Dim cont As Object, c As Object, cap As Variant
    For Each cont In frm.controls
        On Error Resume Next
        ' コンテナ（Frame/Pageなど）だけ調べる
        If Not cont.controls Is Nothing Then
            For Each c In cont.controls
                If TypeName(c) = "CheckBox" Then
                    For Each cap In captions
                        If Trim$(c.caption) = CStr(cap) Then
                            Set FindGroupByAnyCaption = cont
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    Next
End Function

' 名前→無ければ代表キャプションで補助具/リスクのフレームを取得
Private Function ResolveGroup(frm As Object, targetName As String, isAids As Boolean) As Object
    ' 1) 名前で探す（自前のFindControlDeepを使う）
    Set ResolveGroup = frm.controls(targetName)
    If Not ResolveGroup Is Nothing Then Exit Function

    ' 2) キャプションから推定
    Dim seeds As Variant
    If isAids Then
        seeds = Array("杖", "歩行器", "シルバーカー", "車いす", "介助ベルト", "スロープ", "経下肢装具", "手すり")
    Else
        seeds = Array("転倒", "誤嚥", "褥瘡", "失禁", "低栄養", "せん妄", "徘徊", "ADL低下")
    End If
    Set ResolveGroup = FindGroupByAnyCaption(frm, seeds)
End Function

' CSV化（Captionをキー）：targetNameが無くても代表キャプションで検出
Public Function SerializeChecks(frm As Object, targetName As String, Optional isAids As Boolean = True) As String
    Dim grp As Object: Set grp = ResolveGroup(frm, targetName, isAids)
    If grp Is Nothing Then Exit Function

    Dim s As String, c As Object
    For Each c In grp.controls
        If TypeName(c) = "CheckBox" Then
            If c.value = True Then
                If LenB(s) > 0 Then s = s & ","
                s = s & Trim$(c.caption)
            End If
        End If
    Next
    SerializeChecks = s
End Function

' CSV → チェック復元
Public Sub DeserializeChecks(frm As Object, targetName As String, ByVal csv As String, Optional isAids As Boolean = True)
    Dim grp As Object: Set grp = ResolveGroup(frm, targetName, isAids)
    If grp Is Nothing Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim p As Variant
    If Len(Trim$(csv)) > 0 Then
        For Each p In Split(csv, ",")
            dict(Trim$(CStr(p))) = True
        Next
    End If

    Dim c As Object
    For Each c In grp.controls
        If TypeName(c) = "CheckBox" Then
            c.value = dict.exists(Trim$(c.caption))
        End If
    Next
End Sub

' IDの最大値+1
Public Function NextID(ws As Worksheet, ByVal cID As Long) As Long
    Dim last As Long: last = ws.Cells(ws.rows.count, cID).End(xlUp).row
    If last < 2 Then NextID = 1: Exit Function
    On Error Resume Next
    NextID = WorksheetFunction.Max(ws.Range(ws.Cells(2, cID), ws.Cells(last, cID))) + 1
    If Err.Number <> 0 Then NextID = 1: Err.Clear
    On Error GoTo 0
End Function


Private Function GetBool(owner As Object, ctlName As String, Optional defaultValue As Boolean = True) As Boolean
    On Error Resume Next
    GetBool = CBool(owner.controls(ctlName).value)
    If Err.Number <> 0 Then GetBool = defaultValue
    On Error GoTo 0
End Function




'=== Compat: SENSE_IO を IO_Sensory にミラー（行 r のみ） ===
Private Sub Mirror_SensoryIO(ws As Worksheet, ByVal r As Long)
    Dim cSrc As Variant, cDst As Long
    cSrc = Application.Match("SENSE_IO", ws.rows(1), 0)
    If IsError(cSrc) Then Exit Sub

    ' 宛先ヘッダ IO_Sensory を確保
    Dim m As Variant, lastCol As Long
    m = Application.Match("IO_Sensory", ws.rows(1), 0)
    If IsError(m) Then
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = "IO_Sensory"
        cDst = lastCol + 1
    Else
        cDst = CLng(m)
    End If

    ws.Cells(r, cDst).Value2 = CStr(ws.Cells(r, CLng(cSrc)).value)
End Sub


'====================================================================
' Debug / Probe セクション（EvalData のスナップショット・ROMヘッダ等）
'  - 本番処理（保存・読込）からは直接呼ばない
'  - 必要なときだけ、Immediate や専用テストマクロから手動で呼び出す
'  - 将来的には modEvalIODebug など別モジュールへ切り出す候補
'====================================================================



Public Sub Debug_IO_Sensory_ADL_Snapshot(ByVal ws As Worksheet, ByVal r As Long)
#If APP_DEBUG Then
    Dim s As String

    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[SENSE][IO]", _
                "row=" & r, _
                "| len=" & Len(s), _
                "| head=" & Left$(s, 200)

    s = ReadStr_Compat("IO_ADL", r, ws)
    Debug.Print "[ADL][IO]", _
                "row=" & r, _
                "| len=" & Len(s), _
                "| head=" & Left$(s, 200)
#End If
End Sub



Public Sub Debug_Sensory_ADL_Raw(ByVal ws As Worksheet, ByVal r As Long)
    Dim cSense As Variant, cADL As Variant, cIOSense As Variant
    Dim lastCol As Long

    cSense = Application.Match("SENSE_IO", ws.rows(1), 0)
    cADL = Application.Match("IO_ADL", ws.rows(1), 0)
    cIOSense = Application.Match("IO_Sensory", ws.rows(1), 0)

    Debug.Print "=== [RAW SENSE/ADL] row=" & r & " ==="

    If Not IsError(cSense) Then
        Debug.Print "SENSE_IO(col" & cSense & ") =", ws.Cells(r, cSense).text
    Else
        Debug.Print "SENSE_IO: <no header>"
    End If

    If Not IsError(cADL) Then
        Debug.Print "IO_ADL(col" & cADL & ") =", ws.Cells(r, cADL).text
    Else
        Debug.Print "IO_ADL: <no header>"
    End If

    If Not IsError(cIOSense) Then
        Debug.Print "IO_Sensory(col" & cIOSense & ") =", ws.Cells(r, cIOSense).text
    Else
        Debug.Print "IO_Sensory: <no header>"
    End If

    ' 近傍確認（構造見る用）
    Debug.Print "SENSE近傍(146-155)=", Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 146), ws.Cells(r, 155)).value)), " | ")
    Debug.Print "ADL近傍  (156-165)=", Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 156), ws.Cells(r, 165)).value)), " | ")

    Debug.Print "=== [/RAW SENSE/ADL] ==="
End Sub







Public Sub Debug_Cols_146_165_Headers(ByVal ws As Worksheet)
    Dim c As Long
    Debug.Print "=== [HEADERS 146-165] ==="
    For c = 146 To 165
        Debug.Print c, ws.Cells(1, c).value
    Next c
    Debug.Print "=== [/HEADERS] ==="
End Sub



Public Sub Debug_Find_IO_Sense_ADL_Sample(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim sSense As String
    Dim sADL As String

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        sSense = Trim$(ws.Cells(r, 152).value) 'SENSE_IO
        sADL = Trim$(ws.Cells(r, 159).value)   'IO_ADL

        If Len(sSense) > 0 Or Len(sADL) > 0 Then
            Debug.Print "=== [Found IO Sample] row=" & r & " ==="
            Debug.Print "SENSE_IO:", Left$(sSense, 200)
            Debug.Print "IO_ADL:", Left$(sADL, 200)
            Exit For
        End If
    Next r

    If r > lastRow Then
        Debug.Print "=== [No SENSE_IO / IO_ADL found] ==="
    End If
End Sub



Public Sub Debug_ListROMHeaders()
    Dim ws As Worksheet
    Dim c As Long
    Dim firstCol As Long, lastCol As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")

    ' ROM系が並んでいる想定レンジだけを見る（必要なら後で微調整）
    firstCol = 150
    lastCol = 260

    Debug.Print "=== [ROM_* HEADERS 150-260] ==="
    For c = firstCol To lastCol
        If LCase$(Left$(CStr(ws.Cells(1, c).value), 4)) = "rom_" Then
            Debug.Print c, ws.Cells(1, c).value
        End If
    Next c
    Debug.Print "=== [/ROM_* HEADERS] ==="
End Sub



Public Sub Debug_ROMRow_Values(ByVal r As Long)
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim h As String, v As String
    Dim hit As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Debug.Print "=== [ROM VALUES row=" & r & "] ==="
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If LCase$(Left$(h, 4)) = "rom_" Then
            v = CStr(ws.Cells(r, c).value)
            If Len(v) > 0 Then
                Debug.Print c, h, v
                hit = hit + 1
                If hit >= 40 Then Exit For   ' ログ暴発防止
            End If
        End If
    Next c

    If hit = 0 Then
        Debug.Print "(no ROM_* values found)"
    End If
    Debug.Print "=== [/ROM VALUES] ==="
End Sub



Public Sub Debug_Find_IO_ROM_Header()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim h As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Debug.Print "=== [FIND IO_ROM HEADER] ==="
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If InStr(1, h, "IO_ROM", vbTextCompare) > 0 Then
            Debug.Print c, "[" & h & "]"
        End If
    Next c
    Debug.Print "=== [/FIND IO_ROM HEADER] ==="
End Sub



Public Sub Debug_Find_ROM_SampleRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim h As String, v As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    Debug.Print "=== [ROM SAMPLE ROW SEARCH] ==="
    For r = 2 To lastRow
        For c = 158 To 260
            h = CStr(ws.Cells(1, c).value)
            If LCase$(Left$(h, 4)) = "rom_" Then
                v = CStr(ws.Cells(r, c).value)
                If Len(v) > 0 Then
                    Debug.Print "row=" & r & ", col=" & c & ", header=" & h
                    Debug.Print "=== [/ROM SAMPLE ROW SEARCH] ==="
                    Exit Sub
                End If
            End If
        Next c
    Next r

    Debug.Print "(no ROM_* values found in 158-260)"
    Debug.Print "=== [/ROM SAMPLE ROW SEARCH] ==="
End Sub



Public Sub Cleanup_ExtraROMColumns()
    Dim ws As Worksheet
    Dim lastCol As Long, c As Long
    Dim h As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' 本来使うROMブロックより右側だけをゴミ候補とする（とりあえず300列以降）
    For c = lastCol To 261 Step -1
        h = CStr(ws.Cells(1, c).value)
        If LCase$(Left$(h, 4)) = "rom_" Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub




Public Function Build_TestEval_IO(owner As Object) As String
    Dim s As String
    Dim v10 As String
    Dim vTUG As String
    Dim v5x As String
    Dim vSemi As String
    Dim vGripR As String
    Dim vGripL As String

    With owner
        v10 = Trim$(.txtTenMWalk.value)
        vTUG = Trim$(.txtTUG.value)
        v5x = Trim$(.txtFiveSts.value)   ' ※コントロール名が違う場合はここだけ調整
        vSemi = Trim$(.txtSemi.value)
        vGripR = Trim$(.txtGripR.value)
        vGripL = Trim$(.txtGripL.value)
    End With

    s = "Test_10MWalk_sec=" & v10
    s = s & "|Test_TUG_sec=" & vTUG
    s = s & "|Test_Grip_R_kg=" & vGripR
    s = s & "|Test_Grip_L_kg=" & vGripL
    s = s & "|Test_5xSitStand_sec=" & v5x
    s = s & "|Test_SemiTandem_sec=" & vSemi

    Build_TestEval_IO = s
End Function





Public Sub Save_TestEvalToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_TestEval 用の列を確保
    c = EnsureHeader(ws, "IO_TestEval")

    ' フォーム上の値から IO 文字列を生成（今は空のままでもOK）
    s = Build_TestEval_IO(owner)

        ' 指定行に上書き保存
    ws.Cells(r, c).Value2 = CStr(s)
    SaveTestEvalMemoColumns ws, r, owner
    ws.Cells(r, 181).value = val(owner.txtTUG.value)


End Sub




Public Sub Load_TestEvalFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
      Dim s As String
    s = ReadStr_Compat("IO_TestEval", r, ws)

    If Len(s) > 0 Then s = Replace(s, "=", ": ")

    
    owner.txtTenMWalk.value = IO_GetVal(s, "Test_10MWalk_sec")
    owner.txtTUG.value = IO_GetVal(s, "Test_TUG_sec")
    owner.txtFiveSts.value = IO_GetVal(s, "Test_5xSitStand_sec")
    owner.txtGripR.value = IO_GetVal(s, "Test_Grip_R_kg")
    owner.txtGripL.value = IO_GetVal(s, "Test_Grip_L_kg")
    owner.txtSemi.value = IO_GetVal(s, "Test_SemiTandem_sec")
    LoadTestEvalMemoColumns ws, r, owner

    ' TODO: ここから下は後で実装（今は触らない）
    ' IO_TestEval を分解して
    ' owner（frmEval）の txtTenMWalk / txtTUG / txtFiveSts /
    ' txtGripR / txtGripL / txtSemi に流し込む
    
    
    ws.Cells(r, 181).value = val(owner.txtTUG.value)
   
End Sub

Private Sub SaveTestEvalMemoColumns(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_10mWalk", "txtMemo_10mWalk"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_TUG", "txtMemo_TUG"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_STS5", "txtMemo_STS5"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_SemiTandem", "txtMemo_SemiTandem"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripR", "txtMemo_GripR"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripL", "txtMemo_GripL"
End Sub

Private Sub LoadTestEvalMemoColumns(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_10mWalk", "txtMemo_10mWalk"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_TUG", "txtMemo_TUG"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_STS5", "txtMemo_STS5"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_SemiTandem", "txtMemo_SemiTandem"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripR", "txtMemo_GripR"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripL", "txtMemo_GripL"
End Sub

Private Sub SaveTestEvalMemoColumn(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, _
                                   ByVal header As String, ByVal ctlName As String)
    Dim c As Long
    c = EnsureHeader(ws, header)
    ws.Cells(r, c).Value2 = GetCtlTextGeneric(owner, ctlName)
End Sub

Private Sub LoadTestEvalMemoColumn(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, _
                                   ByVal header As String, ByVal ctlName As String)
    Dim c As Long
    c = FindColByHeaderExact(ws, header)
    If c = 0 Then Exit Sub

    SetCtlValueSafe owner, ctlName, CStr(ws.Cells(r, c).Value2)
End Sub

Public Sub Load_WalkIndepFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim v As String
    Dim cmb As Object
    Dim c As Object
    Dim parts() As String
    Dim i As Long
    Dim nm As String
    Dim vLevel As String       '★ 追加：自立度
    Dim cLvl As Object         '★ 追加：自立度コンボ用


       ' IO_WalkIndep の文字列を取得
    s = ReadStr_Compat("IO_WalkIndep", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' TestEval と同じパターンに合わせて "Key=Val" → "Key: Val" に変形
    s = Replace(s, "=", ": ")

    ' --- 自立度（Walk_IndepLevel） ---
    vLevel = IO_GetVal(s, "Walk_IndepLevel")
    If Len(vLevel) > 0 Then
        ' Tag="WalkIndepLevel" のコンボを探して値を戻す
        Set cLvl = Nothing
        For Each c In owner.controls
            If TypeName(c) = "ComboBox" Then
                If c.tag = "WalkIndepLevel" Then
                    Set cLvl = c
                    Exit For
                End If
            End If
        Next c
        If Not cLvl Is Nothing Then
            cLvl.value = vLevel
        End If
    End If

    ' --- 距離 ---
    v = IO_GetVal(s, "Walk_Distance")
    Set cmb = FindControlRecursive(owner, "cmbWalkDistance")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 屋外 ---
    v = IO_GetVal(s, "Walk_Outdoor")
    Set cmb = FindControlRecursive(owner, "cmbWalkOutdoor")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 速度 ---
    v = IO_GetVal(s, "Walk_Speed")
    Set cmb = FindControlRecursive(owner, "cmbWalkSpeed")
    If cmb Is Nothing Then Set cmb = FindControlByTagRecursive(owner, "cmbGaitSpeedDetail")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 安定性チェック（chkWalkStab_*）を一度全部OFF ---
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If StrComp(Left$(nm, 12), "chkWalkStab_", vbTextCompare) = 0 Then
                c.value = False
            End If
        End If
    Next c

    ' --- 安定性の保存文字列を展開して、該当チェックをON ---
    v = IO_GetVal(s, "Walk_Stab")   ' 例： "chkWalkStab_Furatsuki/chkWalkStab_FallRisk"
    If Len(v) > 0 Then
        parts = Split(v, "/")
        For i = LBound(parts) To UBound(parts)
            nm = Trim$(parts(i))
            If Len(nm) > 0 Then
                Set c = FindControlRecursive(owner, nm)
                If Not c Is Nothing Then
                    If TypeName(c) = "CheckBox" Then c.value = True
                End If
            End If
        Next i
    End If

    v = IO_GetVal(s, "Walk_Assistive")
    DeserializeCheckedCaptionsByTag owner, "AssistiveGroup", v
End Sub

Public Sub Load_WalkRLAFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim phases As Variant
    Dim phase As Variant
    Dim probs As String
    Dim level As String
    Dim parts() As String
    Dim i As Long
    Dim c As Object
    Dim nm As String

    ' IO_WalkRLA の文字列を取得
    s = ReadStr_Compat("IO_WalkRLA", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' まず、RLA 関連のチェック・レベルを全部リセット
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "RLA_", vbTextCompare) = 1 Then
                c.value = False
            End If
        ElseIf TypeName(c) = "OptionButton" Then
            If InStr(1, c.groupName, "IC", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "LR", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "MSt", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "TSt", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "PSw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "ISw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "MSw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "TSw", vbTextCompare) = 1 Then
                c.value = False
            End If
        End If
    Next c

    ' TestEval と同じパターンに合わせて "Key=Val" → "Key: Val"
    s = Replace(s, "=", ": ")

    ' 立脚期＋遊脚期のキー
    phases = Array("IC", "LR", "MSt", "TSt", "PSw", "ISw", "MSw", "TSw")

    For Each phase In phases
        ' Problems と Level を取り出し
        probs = IO_GetVal(s, "RLA_" & CStr(phase) & "_Problems")
        level = IO_GetVal(s, "RLA_" & CStr(phase) & "_Level")

        ' --- 問題（CheckBox：Caption一致でON） ---
        If Len(probs) > 0 Then
            parts = Split(probs, "/")
            For i = LBound(parts) To UBound(parts)
                Dim cap As String
                cap = Trim$(parts(i))
                If Len(cap) > 0 Then
                    For Each c In owner.controls
                        If TypeName(c) = "CheckBox" Then
                            nm = CStr(c.name)
                            If InStr(1, nm, "RLA_" & CStr(phase) & "_", vbTextCompare) = 1 Then
                                If CStr(c.caption) = cap Then
                                    c.value = True
                                End If
                            End If
                        End If
                    Next c
                End If
            Next i
        End If

        ' --- レベル（OptionButton：GroupName=phase & Caption一致でON） ---
        If Len(level) > 0 Then
            For Each c In owner.controls
                If TypeName(c) = "OptionButton" Then
                    If StrComp(c.groupName, CStr(phase), vbTextCompare) = 0 Then
                        If CStr(c.caption) = level Then
                            c.value = True
                        End If
                    End If
                End If
            Next c
        End If
    Next phase
End Sub



Public Sub Load_WalkAbnFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim v As String
    Dim parts() As String
    Dim i As Long
    Dim nm As String
    Dim c As Object

    ' IO_WalkAbn の文字列取得
    s = ReadStr_Compat("IO_WalkAbn", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' 一旦、fraWalkAbn_* の全チェックをOFFにする
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "fraWalkAbn_", vbTextCompare) = 1 Then
                c.value = False
            End If
        End If
    Next c

    ' s の中身（例： "fraWalkAbn_A_chk0|fraWalkAbn_C_chk3"）を展開
    parts = Split(s, "|")
    For i = LBound(parts) To UBound(parts)
        nm = Trim$(parts(i))
        If Len(nm) > 0 Then
            Set c = FindControlRecursive(owner, nm)
            If Not c Is Nothing Then
                If TypeName(c) = "CheckBox" Then
                    c.value = True
                End If
            End If
        End If
    Next i
End Sub





Public Sub Save_WalkIndepToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkIndep 用の列を確保
    c = EnsureHeader(ws, "IO_WalkIndep")

    ' フォーム上の値から IO 文字列を生成
    s = Build_WalkIndep_IO(owner)

    ' 指定行に上書き保存
    ws.Cells(r, c).Value2 = CStr(s)

End Sub



Private Function FindControlRecursive(parent As Object, name As String) As Object
    Dim ctl As Object
    For Each ctl In parent.controls
        If StrComp(ctl.name, name, vbTextCompare) = 0 Then
            Set FindControlRecursive = ctl
            Exit Function
        End If
        ' Frame や MultiPage の場合は再帰検索
        On Error Resume Next
        If ctl.controls.count > 0 Then
            Dim subCtl As Object
            Set subCtl = FindControlRecursive(ctl, name)
            If Not subCtl Is Nothing Then
                Set FindControlRecursive = subCtl
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next ctl
End Function

Private Function FindControlByTagRecursive(parent As Object, tagName As String) As Object
    Dim ctl As Object
    For Each ctl In parent.controls
        If StrComp(CStr(ctl.tag), tagName, vbTextCompare) = 0 Then
            Set FindControlByTagRecursive = ctl
            Exit Function
        End If
        On Error Resume Next
        If ctl.controls.count > 0 Then
            Dim subCtl As Object
            Set subCtl = FindControlByTagRecursive(ctl, tagName)
            If Not subCtl Is Nothing Then
                Set FindControlByTagRecursive = subCtl
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next ctl
End Function

Private Function SerializeCheckedCaptionsByTag(parent As Object, groupTag As String) As String
    Dim ctl As Object
    Dim s As String
    Dim childCsv As String

    For Each ctl In parent.controls
        If TypeName(ctl) = "CheckBox" Then
            If StrComp(CStr(ctl.tag), groupTag, vbTextCompare) = 0 Then
                If ctl.value = True Then
                    If LenB(s) > 0 Then s = s & ","
                    s = s & Trim$(ctl.caption)
                End If
            End If
        End If

        On Error Resume Next
        If ctl.controls.count > 0 Then
            childCsv = SerializeCheckedCaptionsByTag(ctl, groupTag)
            If LenB(childCsv) > 0 Then
                If LenB(s) > 0 Then s = s & ","
                s = s & childCsv
            End If
        End If
        On Error GoTo 0
    Next ctl

    SerializeCheckedCaptionsByTag = s
End Function

Private Sub DeserializeCheckedCaptionsByTag(parent As Object, groupTag As String, ByVal csv As String)
    Dim dict As Object
    Dim parts As Variant
    Dim p As Variant
    Dim ctl As Object

    Set dict = CreateObject("Scripting.Dictionary")
    If Len(Trim$(csv)) > 0 Then
        parts = Split(csv, ",")
        For Each p In parts
            dict(Trim$(CStr(p))) = True
        Next p
    End If

    For Each ctl In parent.controls
        If TypeName(ctl) = "CheckBox" Then
            If StrComp(CStr(ctl.tag), groupTag, vbTextCompare) = 0 Then
                ctl.value = dict.exists(Trim$(ctl.caption))
            End If
        End If

        On Error Resume Next
        If ctl.controls.count > 0 Then
            DeserializeCheckedCaptionsByTag ctl, groupTag, csv
        End If
        On Error GoTo 0
    Next ctl
End Sub


Public Function Build_WalkIndep_IO(owner As Object) As String
    Dim vDist As String
    Dim vOut As String
    Dim vSpeed As String
    Dim vAssistive As String
    Dim s As String
    Dim hits As Collection
    Dim c As Object
    Dim nm As String
    Dim stab As String
    Dim i As Long
    Dim vLevel As String   '★ 自立度



   Dim cLvl As Object
Set cLvl = FindControlRecursive(owner, "cmbWalkIndep")
If cLvl Is Nothing Then
    ' タグで検索する（今回の正式ルート）
    For Each c In owner.controls
        If TypeName(c) = "ComboBox" Then
            If c.tag = "WalkIndepLevel" Then
                Set cLvl = c
                Exit For
            End If
        End If
    Next c
End If
If Not cLvl Is Nothing Then vLevel = Trim$(cLvl.value)



    ' 距離・屋外・速度
    Set c = FindControlRecursive(owner, "cmbWalkDistance")
    If Not c Is Nothing Then vDist = Trim$(c.value)

    Set c = FindControlRecursive(owner, "cmbWalkOutdoor")
    If Not c Is Nothing Then vOut = Trim$(c.value)

    Set c = FindControlRecursive(owner, "cmbWalkSpeed")
    If c Is Nothing Then Set c = FindControlByTagRecursive(owner, "cmbGaitSpeedDetail")
    If Not c Is Nothing Then vSpeed = Trim$(c.value)
    

    ' 安定性チェック（chkWalkStab_～ を全部拾う）
    Set hits = New Collection
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If StrComp(Left$(nm, 12), "chkWalkStab_", vbTextCompare) = 0 Then
                If c.value = True Then
                    ' 名前そのものか、末尾だけにするかはあとで調整可
                    hits.Add nm
                End If
            End If
        End If
    Next c

    ' 安定性のチェック名を「/」区切りで1本の文字列にまとめる
    For i = 1 To hits.count
        If i > 1 Then stab = stab & "/"
        stab = stab & hits(i)
    Next i

        ' IO 文字列組み立て
    s = "Walk_IndepLevel=" & vLevel
    s = s & "|Walk_Distance=" & vDist
    s = s & "|Walk_Outdoor=" & vOut
    s = s & "|Walk_Stab=" & stab
    s = s & "|Walk_Speed=" & vSpeed

    vAssistive = SerializeCheckedCaptionsByTag(owner, "AssistiveGroup")
    s = s & "|Walk_Assistive=" & vAssistive

    Build_WalkIndep_IO = s
End Function



Public Function Build_WalkAbn_IO(owner As Object) As String
    Dim c As Object
    Dim hits As Collection
    Dim s As String
    Dim nm As String
    
    Set hits = New Collection
    
    For Each c In owner.controls
        ' fraWalkAbn_?_chk? という名前の CheckBox だけ拾う
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "fraWalkAbn_", vbTextCompare) = 1 Then
                If c.value = True Then
                    hits.Add nm
                End If
            End If
        End If
    Next c
    
    ' 1つもチェックが無ければ空文字を返す
    If hits.count = 0 Then
        Build_WalkAbn_IO = ""
        Exit Function
    End If
    
    ' fraWalkAbn_A_chk0|fraWalkAbn_A_chk3|… という形で連結
    Dim i As Long
    For i = 1 To hits.count
        If i > 1 Then s = s & "|"
        s = s & hits(i)
    Next i
    
    Build_WalkAbn_IO = s
End Function


Public Sub Save_WalkAbnToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkAbn 用の列を確保
    c = EnsureHeader(ws, "IO_WalkAbn")

    ' フォームのチェック状態から IO 文字列を生成
    s = Build_WalkAbn_IO(owner)

    ' 指定行に上書き保存
    ws.Cells(r, c).Value2 = CStr(s)

End Sub




Public Function Build_WalkRLA_IO(owner As Object) As String
    Dim phases As Variant
    Dim phase As Variant
    Dim c As Object
    Dim probs As Collection
    Dim probsStr As String
    Dim level As String
    Dim s As String
    Dim first As Boolean
    Dim i As Long
    Dim nm As String

    ' 立脚期＋遊脚期のキー（Build_RLA_ChecksPart と同じ）
    phases = Array("IC", "LR", "MSt", "TSt", "PSw", "ISw", "MSw", "TSw")
    first = True

    For Each phase In phases
        Set probs = New Collection
        probsStr = ""
        level = ""

        ' --- チェック（RLA_<phase>_～）を拾う ---
        For Each c In owner.controls
            If TypeName(c) = "CheckBox" Then
                nm = CStr(c.name)
                If InStr(1, nm, "RLA_" & CStr(phase) & "_", vbTextCompare) = 1 Then
                    If c.value = True Then
                        probs.Add c.caption   ' 例）可動域不足 / 筋力低下 など
                    End If
                End If
            End If
        Next c

        ' 問題リストを "/" 区切りで 1 本にする
        If probs.count > 0 Then
            For i = 1 To probs.count
                If i > 1 Then probsStr = probsStr & "/"
                probsStr = probsStr & probs(i)
            Next i
        End If

        ' --- レベル（OptionButton, GroupName=phase）を拾う ---
        For Each c In owner.controls
            If TypeName(c) = "OptionButton" Then
                If StrComp(c.groupName, CStr(phase), vbTextCompare) = 0 Then
                    If c.value = True Then
                        level = CStr(c.caption)   ' 軽度 / 中等度 / 高度
                        Exit For
                    End If
                End If
            End If
        Next c

        ' --- IO セグメント組み立て ---
        Dim seg As String
        seg = "RLA_" & CStr(phase) & "_Problems=" & probsStr & _
              "|RLA_" & CStr(phase) & "_Level=" & level

        If first Then
            s = seg
            first = False
        Else
            s = s & "|" & seg
        End If
    Next phase

    Build_WalkRLA_IO = s
End Function



Public Sub Save_WalkRLAToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkRLA 用の列を確保（列4にヘッダ IO_WalkRLA がある前提）
    c = EnsureHeader(ws, "IO_WalkRLA")

    ' フォーム上のRLAチェック・レベルからIO文字列を生成
    s = Build_WalkRLA_IO(owner)

    ' 指定行に上書き保存
    ws.Cells(r, c).Value2 = CStr(s)

End Sub

Private Function GetCogTabsSafe(ByVal owner As Object) As Object
    Dim mp As Object

    On Error Resume Next
    Set mp = owner.GetCogTabs
    On Error GoTo 0
    
    Set GetCogTabsSafe = mp
    If Not mp Is Nothing Then Set GetCogTabsSafe = mp
End Function



Public Sub Save_CognitionMental_AtRow(ws As Worksheet, r As Long, owner As Object)
    Dim frm As Object
    Dim col As Long
    Dim v As Variant
    Dim f As MSForms.Frame
    Dim c As MSForms.Control
    Dim bpsd As String
    Dim mpCog As Object
    Dim pgCog As Object
    Dim pgMental As Object
    
    Set frm = owner   ' frmEval を受け取る想定
    Set mpCog = GetCogTabsSafe(frm)
    If mpCog Is Nothing Then Exit Sub
    Set pgCog = mpCog.Pages("pgCognition")
    Set pgMental = mpCog.Pages("pgMental")
        
    
    '=== 認知：中核6項目 =====================================
    
    ' 記憶
    col = HeaderCol_Compat("IO_Cog_Memory", ws)
    If col > 0 Then
        v = pgCog.controls("cmbCogMemory").value


        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 注意
    col = HeaderCol_Compat("IO_Cog_Attention", ws)
    If col > 0 Then
        v = pgCog.controls("cmbCogAttention").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 見当識
    col = HeaderCol_Compat("IO_Cog_Orientation", ws)
    If col > 0 Then
            v = pgCog.controls("cmbCogOrientation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 判断
    col = HeaderCol_Compat("IO_Cog_Judgment", ws)
    If col > 0 Then
            v = pgCog.controls("cmbCogJudgement").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 遂行機能
    col = HeaderCol_Compat("IO_Cog_Executive", ws)
    If col > 0 Then
             v = pgCog.controls("cmbCogExecutive").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 言語
    col = HeaderCol_Compat("IO_Cog_Language", ws)
    If col > 0 Then
             v = pgCog.controls("cmbCogLanguage").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    '=== 認知：認知症の種類＋備考 ==============================
    
    col = HeaderCol_Compat("IO_Cog_DementiaType", ws)
    If col > 0 Then
             v = pgCog.controls("cmbDementiaType").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    col = HeaderCol_Compat("IO_Cog_DementiaNote", ws)
    If col > 0 Then
            v = pgCog.controls("txtDementiaNote").text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
        '=== 認知：BPSD（チェックが入っている項目を | 区切りで保存） ===
    
    bpsd = ""
     With pgCog
        For Each c In .controls
            If TypeName(c) = "CheckBox" Then
                If c.value = True Then
                    If Len(bpsd) > 0 Then bpsd = bpsd & "|"
                    bpsd = bpsd & CStr(c.caption)
                End If
            End If
        Next c
    End With
    
    col = HeaderCol_Compat("IO_Cog_BPSD", ws)
    If col > 0 Then
        ws.Cells(r, col).value = bpsd
    End If

    
    '=== 精神面タブ ============================================
    
    ' 気分
    col = HeaderCol_Compat("IO_Mental_Mood", ws)
    If col > 0 Then
             v = pgMental.controls("cmbMood").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 意欲
    col = HeaderCol_Compat("IO_Mental_Motivation", ws)
    If col > 0 Then
            v = pgMental.controls("cmbMotivation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 不安
    col = HeaderCol_Compat("IO_Mental_Anxiety", ws)
    If col > 0 Then
            v = pgMental.controls("cmbAnxiety").value
            
        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 対人関係
    col = HeaderCol_Compat("IO_Mental_Relation", ws)
    If col > 0 Then
            v = pgMental.controls("cmbRelation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 睡眠
    col = HeaderCol_Compat("IO_Mental_Sleep", ws)
    If col > 0 Then
            v = pgMental.controls("cmbSleep").value
            
        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 精神面・備考
    col = HeaderCol_Compat("IO_Mental_Note", ws)
    If col > 0 Then
            v = pgMental.controls("txtMentalNote").text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
End Sub

   



Public Sub Load_CognitionMental_FromRow(ws As Worksheet, ByVal r As Long, owner As Object)

    Dim f As MSForms.Frame
    Dim mp As Object
    Dim pgCog As Object
    Dim pgMental As Object

    Dim v As Variant
    Dim s As String
    Dim arr() As String
    Dim i As Long, j As Long
    Dim chk As MSForms.CheckBox

    '=== UI ルート取得（絶対名を前提）===
    Set mp = GetCogTabsSafe(owner)
    If mp Is Nothing Then Exit Sub
    Set pgCog = mp.Pages("pgCognition")
    Set pgMental = mp.Pages("pgMental")

    '=== 認知側 combobox 群 ===
    LoadComboValueByHeader ws, r, "IO_Cog_Memory", pgCog, "cmbCogMemory"
    LoadComboValueByHeader ws, r, "IO_Cog_Attention", pgCog, "cmbCogAttention"
    LoadComboValueByHeader ws, r, "IO_Cog_Orientation", pgCog, "cmbCogOrientation"
    LoadComboValueByHeader ws, r, "IO_Cog_Judgment", pgCog, "cmbCogJudgement"
    LoadComboValueByHeader ws, r, "IO_Cog_Executive", pgCog, "cmbCogExecutive"
    LoadComboValueByHeader ws, r, "IO_Cog_Language", pgCog, "cmbCogLanguage"
    LoadComboValueByHeader ws, r, "IO_Cog_DementiaType", pgCog, "cmbDementiaType"

    v = ReadValueByCompatHeader(ws, r, "IO_Cog_DementiaNote")
    pgCog.controls("txtDementiaNote").text = CStr(v)

    '=== BPSD（chkBPSD0?10）===
    ' 1) 全部一度クリア
    For i = 0 To 10
        Set chk = pgCog.controls("chkBPSD" & CStr(i))
        chk.value = False
    Next i

    ' 2) セル文字列を | で分解し、Caption と一致するチェックボックスをON
    s = CStr(ReadValueByCompatHeader(ws, r, "IO_Cog_BPSD"))
    If Len(s) > 0 Then
        arr = Split(s, "|")
        For i = LBound(arr) To UBound(arr)
            For j = 0 To 10
                Set chk = pgCog.controls("chkBPSD" & CStr(j))
                If chk.caption = arr(i) Then
                    chk.value = True
                    Exit For
                End If
            Next j
        Next i
    End If

    '=== 精神面 combobox / note ===
    LoadComboValueByHeader ws, r, "IO_Mental_Mood", pgMental, "cmbMood"
    LoadComboValueByHeader ws, r, "IO_Mental_Motivation", pgMental, "cmbMotivation"
    LoadComboValueByHeader ws, r, "IO_Mental_Anxiety", pgMental, "cmbAnxiety"
    LoadComboValueByHeader ws, r, "IO_Mental_Relation", pgMental, "cmbRelation"
    LoadComboValueByHeader ws, r, "IO_Mental_Sleep", pgMental, "cmbSleep"

    v = ReadValueByCompatHeader(ws, r, "IO_Mental_Note")
    pgMental.controls("txtMentalNote").text = CStr(v)
End Sub


Private Sub LoadComboValueByHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal headerName As String, _
                                   ByVal parent As Object, ByVal comboName As String)
    Dim cmb As MSForms.ComboBox
    Dim v As String

    Set cmb = parent.controls(comboName)
    v = CStr(ReadValueByCompatHeader(ws, r, headerName))

    If Len(v) = 0 Then
        cmb.ListIndex = -1
    ElseIf ComboBoxHasValue(cmb, v) Then
        cmb.value = v
    Else
        cmb.ListIndex = -1
        Debug.Print "[Load_CognitionMental] skip invalid combo value"; _
                    " header=" & headerName & " combo=" & comboName & " value=" & v
    End If
End Sub

Private Function ReadValueByCompatHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal headerName As String) As Variant
    Dim headers As Variant
    Dim i As Long
    Dim col As Long
    Dim v As Variant

    headers = CompatHeaderNames(headerName)
    For i = LBound(headers) To UBound(headers)
        col = HeaderCol(CStr(headers(i)), ws)
        If col > 0 Then
            v = ws.Cells(r, col).value
            If Not IsNull(v) Then
                If Len(CStr(v)) > 0 Then
                    ReadValueByCompatHeader = v
                    Exit Function
                End If
            End If
        End If
    Next i

    ReadValueByCompatHeader = vbNullString
End Function

Private Function ComboBoxHasValue(ByVal cmb As MSForms.ComboBox, ByVal target As String) As Boolean
    Dim i As Long

    For i = 0 To cmb.ListCount - 1
        If StrComp(CStr(cmb.List(i)), target, vbBinaryCompare) = 0 Then
            ComboBoxHasValue = True
            Exit Function
        End If
    Next i
End Function



Private Function ComposeDailyLogBody(ByVal training As String, ByVal reaction As String, ByVal abnormal As String, ByVal plan As String) As String
    ComposeDailyLogBody = "【実施内容】" & vbCrLf & training & vbCrLf & vbCrLf & _
                          "【利用者の反応】" & vbCrLf & reaction & vbCrLf & vbCrLf & _
                          "【異常所見】" & vbCrLf & abnormal & vbCrLf & vbCrLf & _
                          "【今後の方針】" & vbCrLf & plan
End Function

Private Sub FillDailyLogFieldsFromBody(ByVal body As String, ByRef training As String, ByRef reaction As String, ByRef abnormal As String, ByRef plan As String)
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long

    training = ""
    reaction = ""
    abnormal = ""
    plan = ""

    p1 = InStr(body, "【実施内容】")
    p2 = InStr(body, "【利用者の反応】")
    p3 = InStr(body, "【異常所見】")
    p4 = InStr(body, "【今後の方針】")

    If p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 Then
        training = Trim$(Mid$(body, p1 + Len("【実施内容】"), p2 - (p1 + Len("【実施内容】"))))
        reaction = Trim$(Mid$(body, p2 + Len("【利用者の反応】"), p3 - (p2 + Len("【利用者の反応】"))))
        abnormal = Trim$(Mid$(body, p3 + Len("【異常所見】"), p4 - (p3 + Len("【異常所見】"))))
        plan = Trim$(Mid$(body, p4 + Len("【今後の方針】")))
    Else
        training = body
    End If
End Sub


Public Function EnsureDailyLogFolderPath() As String

    Dim dataFolder As String
    Dim logsFolder As String

    dataFolder = ThisWorkbook.path & "\data"
    If Dir(dataFolder, vbDirectory) = "" Then MkDir dataFolder

    logsFolder = dataFolder & "\logs"
    If Dir(logsFolder, vbDirectory) = "" Then MkDir logsFolder

    EnsureDailyLogFolderPath = logsFolder & "\"

End Function


Public Function GetDailyLogFilePath(ByVal d As Date) As String

    Dim basePath As String
    basePath = EnsureDailyLogFolderPath()

    GetDailyLogFilePath = basePath & "DailyLog_" & Format$(d, "yyyy") & ".xlsx"

End Function

Private Function EnsureDailyLogSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("DailyLog")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count))
        ws.name = "DailyLog"
    End If

      If Trim$(CStr(ws.Cells(1, 1).value)) = "" Then ws.Cells(1, 1).value = "LogID"
      If Trim$(CStr(ws.Cells(1, 2).value)) = "" Then ws.Cells(1, 2).value = "利用者ID"
      If Trim$(CStr(ws.Cells(1, 3).value)) = "" Then ws.Cells(1, 3).value = "利用者名"
      If Trim$(CStr(ws.Cells(1, 4).value)) = "" Then ws.Cells(1, 4).value = "利用日"
      If Trim$(CStr(ws.Cells(1, 5).value)) = "" Then ws.Cells(1, 5).value = "記録本文"
      If Trim$(CStr(ws.Cells(1, 6).value)) = "" Then ws.Cells(1, 6).value = "記録者"
      If Trim$(CStr(ws.Cells(1, 7).value)) = "" Then ws.Cells(1, 7).value = "更新日時"

    Set EnsureDailyLogSheet = ws
End Function

Private Function OpenDailyLogWorkbook(ByVal d As Date, ByVal createIfMissing As Boolean, ByRef openedHere As Boolean) As Workbook
    Dim filePath As String
    Dim wb As Workbook

    Call EnsureDailyLogFolderPath
    filePath = GetDailyLogFilePath(d)
    
    

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenDailyLogWorkbook = wb
            openedHere = False
            Exit Function
        End If
    Next wb

    If Dir(filePath) <> "" Then
        Set OpenDailyLogWorkbook = Application.Workbooks.Open(filePath)
        openedHere = True
        Exit Function
    End If
     

    If Not createIfMissing Then Exit Function
    
    
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Call EnsureDailyLogSheet(wb)
    wb.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    Set OpenDailyLogWorkbook = wb
    openedHere = True
End Function


Public Function GetDailyLogSheetByDate(ByVal d As Date, ByVal createIfMissing As Boolean, ByRef wb As Workbook, ByRef openedHere As Boolean) As Worksheet
    Set wb = OpenDailyLogWorkbook(d, createIfMissing, openedHere)
    If wb Is Nothing Then Exit Function
    Set GetDailyLogSheetByDate = EnsureDailyLogSheet(wb)
End Function

Private Function GenerateDailyLogID(ByVal ws As Worksheet, ByVal logDate As Date) As String
    Dim y As String
    Dim lastRow As Long
    Dim r As Long
    Dim maxSeq As Long
    Dim token As String
    Dim seqPart As String

    y = Format$(logDate, "yyyy")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
    token = Trim$(CStr(ws.Cells(r, 1).value))
    If Len(token) >= 12 And Left$(token, 4) = y And Mid$(token, 5, 1) = "-" Then
        seqPart = Mid$(token, 6)
        If IsNumeric(seqPart) Then
            If CLng(seqPart) > maxSeq Then maxSeq = CLng(seqPart)
        End If
    End If
Next r

GenerateDailyLogID = y & "-" & Format$(maxSeq + 1, "000000")
End Function

Public Sub Save_DailyLog_FromForm(owner As Object)
    ' 手動保存時は SaveDailyLog_Append を使用
    mDailyLogManual = True
    SaveDailyLog_Append owner
    mDailyLogManual = False
End Sub
    

Private Sub SortDailyLogFilesByYearDesc(ByRef years() As Long, ByRef files() As String, ByVal itemCount As Long)
    Dim i As Long
    Dim j As Long
    Dim tmpYear As Long
    Dim tmpFile As String

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If years(j) > years(i) Then
                tmpYear = years(i)
                years(i) = years(j)
                years(j) = tmpYear

                tmpFile = files(i)
                files(i) = files(j)
                files(j) = tmpFile
            End If
        Next j
    Next i
End Sub
    




Public Sub Load_DailyLog_Latest_FromForm(owner As Object)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wbOpenedHere As Boolean
    Dim candidateWb As Workbook
    Dim candidateOpenedHere As Boolean
    Dim f As Object
    Dim txtName As Object
    Dim txtDate As Object
    Dim txtStaff As Object
    Dim txtTraining As Object
    Dim txtReaction As Object
    Dim txtAbnormal As Object
    Dim txtPlan As Object
    Dim hdr As Object
    Dim txtHdrPID As Object
    Dim lastRow As Long
    Dim r As Long
    Dim targetName As String
    Dim targetPid As String
    Dim hit As Boolean
    Dim body As String
    Dim basePath As String
    Dim fileName As String
    Dim token As String
    Dim itemCount As Long
    Dim years() As Long
    Dim files() As String
    Dim idx As Long
    Dim y As Long
    Dim awb As Workbook
    Dim filePath As String


    '--- フォーム上のコントロール取得 ---
    Set txtName = SafeGetControl(owner, "txtName")
    Set f = ResolveDailyLogRoot(owner)
    If txtName Is Nothing Or f Is Nothing Then Exit Sub

    Set txtDate = ResolveDailyLogControl(owner, "txtDailyDate")
    Set txtStaff = ResolveDailyLogControl(owner, "txtDailyStaff")
    Set txtTraining = ResolveDailyLogControl(owner, "txtDailyTraining")
    Set txtReaction = ResolveDailyLogControl(owner, "txtDailyReaction")
    Set txtAbnormal = ResolveDailyLogControl(owner, "txtDailyAbnormal")
    Set txtPlan = ResolveDailyLogControl(owner, "txtDailyPlan")
    Set hdr = SafeGetControl(owner, "frHeader")
    Set txtHdrPID = SafeGetControl(hdr, "txtHdrPID")
    If txtDate Is Nothing Or txtStaff Is Nothing Or txtTraining Is Nothing Or txtReaction Is Nothing Or txtAbnormal Is Nothing Or txtPlan Is Nothing Or txtHdrPID Is Nothing Then Exit Sub
    



    '--- 該当利用者の「最新（いちばん下）」の行を探す ---
    targetName = Trim$(CStr(txtName.value))
    targetPid = Trim$(CStr(txtHdrPID.value))
    If targetPid = "" And targetName = "" Then GoTo FinallyExit

    basePath = EnsureDailyLogFolderPath()
    fileName = Dir(basePath & "DailyLog_*.xlsx")

    Do While fileName <> ""
        token = Mid$(fileName, Len("DailyLog_") + 1, 4)
        If Len(token) = 4 And IsNumeric(token) Then
            y = CLng(token)
            itemCount = itemCount + 1
            ReDim Preserve years(1 To itemCount)
            ReDim Preserve files(1 To itemCount)
            years(itemCount) = y
            files(itemCount) = fileName
        End If
        fileName = Dir()
    Loop

    If itemCount = 0 Then GoTo FinallyExit

    SortDailyLogFilesByYearDesc years, files, itemCount


    hit = False
    For idx = 1 To itemCount
        filePath = basePath & files(idx)
        Set candidateWb = Nothing
        candidateOpenedHere = False

        For Each awb In Application.Workbooks
            If StrComp(awb.FullName, filePath, vbTextCompare) = 0 Then
                Set candidateWb = awb
                Exit For
            End If
        Next awb

        If candidateWb Is Nothing Then
            Set candidateWb = Application.Workbooks.Open(filePath)
            candidateOpenedHere = True
        End If
        

        Set ws = EnsureDailyLogSheet(candidateWb)
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.rows.count, 3).End(xlUp).row
            For r = lastRow To 2 Step -1
                 If targetPid <> "" Then
                    If Trim$(CStr(ws.Cells(r, 2).value)) = targetPid Then
                        hit = True
                        Set wb = candidateWb
                        wbOpenedHere = candidateOpenedHere
                        Exit For
                    End If
                ElseIf Trim$(CStr(ws.Cells(r, 3).value)) = targetName Then
                    hit = True
                    Set wb = candidateWb
                    wbOpenedHere = candidateOpenedHere
                    Exit For
                End If
            Next r
        End If

        If hit Then Exit For

        If candidateOpenedHere Then
            candidateWb.Close SaveChanges:=False
        End If
    Next idx


    If Not hit Then GoTo FinallyExit
    
    

    '--- 見つかった行をフォームへ反映 ---
    body = CStr(ws.Cells(r, 5).value)

    txtDate.value = ws.Cells(r, 4).value
    txtStaff.value = ws.Cells(r, 6).value
    FillDailyLogFieldsFromBody body, txtTraining.value, txtReaction.value, txtAbnormal.value, txtPlan.value

FinallyExit:
    If wbOpenedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub




Public Sub SaveDailyLog_Append(owner As Object)

    
    ' 専用ボタンからの呼び出し以外では何もしない
    If Not mDailyLogManual Then Exit Sub


    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wbOpenedHere As Boolean
    Dim r As Long
    Dim f As Object
    Dim dt As Variant
    Dim nm As String
    Dim pid As String
    Dim staff As String
    Dim note As String
    Dim training As String
    Dim reaction As String
    Dim abnormal As String
    Dim plan As String
    Dim logDate As Date
    Dim lastRow As Long
    Dim hitRow As Long


    Set f = ResolveDailyLogRoot(owner)
    If f Is Nothing Then Exit Sub

    Dim txtDailyDate As Object
    Dim txtDailyStaff As Object
    Dim txtDailyTraining As Object
    Dim txtDailyReaction As Object
    Dim txtDailyAbnormal As Object
    Dim txtDailyPlan As Object
    Dim hdr As Object
    Dim txtHdrName As Object
    Dim txtHdrPID As Object
    
    

    Set txtDailyDate = ResolveDailyLogControl(owner, "txtDailyDate")
    Set txtDailyStaff = ResolveDailyLogControl(owner, "txtDailyStaff")
    Set txtDailyTraining = ResolveDailyLogControl(owner, "txtDailyTraining")
    Set txtDailyReaction = ResolveDailyLogControl(owner, "txtDailyReaction")
    Set txtDailyAbnormal = ResolveDailyLogControl(owner, "txtDailyAbnormal")
    Set txtDailyPlan = ResolveDailyLogControl(owner, "txtDailyPlan")
    Set hdr = SafeGetControl(owner, "frHeader")
    Set txtHdrName = SafeGetControl(hdr, "txtHdrName")
    If txtDailyDate Is Nothing Or txtDailyStaff Is Nothing Or txtDailyTraining Is Nothing Or txtDailyReaction Is Nothing Or txtDailyAbnormal Is Nothing Or txtDailyPlan Is Nothing Or txtHdrName Is Nothing Then Exit Sub
    
    Set txtHdrPID = SafeGetControl(hdr, "txtHdrPID")
    If txtDailyDate Is Nothing Or txtDailyStaff Is Nothing Or txtDailyTraining Is Nothing Or txtDailyReaction Is Nothing Or txtDailyAbnormal Is Nothing Or txtDailyPlan Is Nothing Or txtHdrName Is Nothing Or txtHdrPID Is Nothing Then Exit Sub

    dt = txtDailyDate.value
    nm = Trim$(txtHdrName.value)
    pid = Trim$(txtHdrPID.value)
    staff = Trim$(txtDailyStaff.value)
    training = CStr(txtDailyTraining.value)
    reaction = CStr(txtDailyReaction.value)
    abnormal = CStr(txtDailyAbnormal.value)
    plan = CStr(txtDailyPlan.value)
    note = ComposeDailyLogBody(training, reaction, abnormal, plan)


    '--- 入力チェック ---
    If nm = "" Then
     MsgBox "利用者名を入力してください。", vbExclamation
     Exit Sub
    End If

    If Not IsDate(dt) Then
        MsgBox "記録日の欄に正しい日付を入力してください。", vbExclamation
        Exit Sub
    End If

    If Trim$(training & reaction & abnormal & plan) = "" Then
        If MsgBox("記録内容が空ですが保存しますか？", vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
    
 End If

    logDate = CDate(dt)
    Set ws = GetDailyLogSheetByDate(logDate, True, wb, wbOpenedHere)
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    hitRow = 0

    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 4).value) Then
            If CLng(CDate(ws.Cells(r, 4).value)) = CLng(logDate) Then
                If pid <> "" Then
                    If Trim$(CStr(ws.Cells(r, 2).value)) = pid Then
                        hitRow = r
                        Exit For
                    End If
                ElseIf Trim$(CStr(ws.Cells(r, 3).value)) = nm Then
                    hitRow = r
                    Exit For
                End If
            End If
        End If
    Next r

    If hitRow = 0 Then
        hitRow = IIf(lastRow < 2, 2, lastRow + 1)
        ws.Cells(hitRow, 1).value = GenerateDailyLogID(ws, logDate)
    ElseIf Trim$(CStr(ws.Cells(hitRow, 1).value)) = "" Then
        ws.Cells(hitRow, 1).value = GenerateDailyLogID(ws, logDate)
    
    End If
    
    

    '--- 追記行を決める（1行目に見出しがある前提）---
    ws.Cells(hitRow, 2).value = pid
    ws.Cells(hitRow, 3).value = nm
    ws.Cells(hitRow, 4).value = logDate
    ws.Cells(hitRow, 4).NumberFormatLocal = "yyyy/mm/dd"
    ws.Cells(hitRow, 5).value = note
    ws.Cells(hitRow, 6).value = staff
    ws.Cells(hitRow, 7).value = Now
    ws.Cells(hitRow, 7).NumberFormatLocal = "yyyy/mm/dd hh:mm"
    

End Sub




'=== Basic.* と日本語ヘッダ列をミラーする汎用ヘルパー =====================

Private Sub MirrorBasicPair( _
        ByVal ws As Worksheet, ByVal rowNum As Long, _
        ByVal colBasic As Long, ByVal colJp As Long)

    Dim vBasic As Variant, vJp As Variant

    If colBasic <= 0 Or colJp <= 0 Then Exit Sub

    vBasic = ws.Cells(rowNum, colBasic).value
    vJp = ws.Cells(rowNum, colJp).value

    ' どちらか片方だけ入っている場合、もう片方へコピー
    If Len(vBasic) = 0 And Len(vJp) > 0 Then
        ws.Cells(rowNum, colBasic).value = vJp
    ElseIf Len(vJp) = 0 And Len(vBasic) > 0 Then
        ws.Cells(rowNum, colJp).value = vBasic
    End If
End Sub

'=== 基本情報の新旧列をミラーする =====================================
'  ・Basic.* と 日本語ヘッダ の両方を「空いている方へ」コピーする
'  ・どちらか片方にしか値がなければ、その値をもう片方へ写すだけ
'  ・両方に値がある場合は何もしない（衝突回避）
Public Sub MirrorBasicRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    ' ID
    MirrorBasicPair ws, rowNum, "Basic.ID", "ID"
    ' 氏名
    MirrorBasicPair ws, rowNum, "Basic.Name", "氏名"
    ' 評価日
    MirrorBasicPair ws, rowNum, "Basic.EvalDate", "評価日"
    ' 年齢
    MirrorBasicPair ws, rowNum, "Basic.Age", "年齢"
    ' 性別
    MirrorBasicPair ws, rowNum, "Basic.Sex", "性別"
    ' 評価者
    MirrorBasicPair ws, rowNum, "Basic.Evaluator", "評価者"
    ' 評価者職種
    MirrorBasicPair ws, rowNum, "Basic.EvaluatorJob", "評価者職種"
    ' 発症日
    MirrorBasicPair ws, rowNum, "Basic.OnsetDate", "発症日"
    ' 患者Needs
    MirrorBasicPair ws, rowNum, "Basic.Needs.Patient", "患者Needs"
    ' 家族Needs
    MirrorBasicPair ws, rowNum, "Basic.Needs.Family", "家族Needs"
    ' 生活状況
    MirrorBasicPair ws, rowNum, "Basic.LifeStatus", "生活状況"
    ' 主診断
    MirrorBasicPair ws, rowNum, "Basic.PrimaryDx", "主診断"
    ' 要介護度
    MirrorBasicPair ws, rowNum, "Basic.CareLevel", "要介護度"

    Exit Sub

ErrHandler:
    
End Sub
'=== 基本情報の新旧列をミラーする =====================================
'  ・Basic.* と 日本語ヘッダ の両方を「空いている方へ」コピーする
'  ・どちらか片方にしか値がなければ、その値をもう片方へ写すだけ
'  ・両方に値がある場合は何もしない（衝突回避）
Public Sub MirrorBasicRow_Eval(ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    Dim pairs As Variant
    Dim i As Long
    Dim headerNew As String, headerOld As String
    Dim cNew As Long, cOld As Long
    Dim vNew As Variant, vOld As Variant
    Dim sNew As String, sOld As String

    ' 対象ペア一覧（左が Basic.*、右が日本語ヘッダ）
    pairs = Array( _
        Array("Basic.ID", "ID"), _
        Array("Basic.Name", "氏名"), _
        Array("Basic.EvalDate", "評価日"), _
        Array("Basic.Age", "年齢"), _
        Array("Basic.Sex", "性別"), _
        Array("Basic.Evaluator", "評価者"), _
        Array("評価者職種", "txtEvaluatorJob"), _
        Array("Basic.OnsetDate", "発症日"), _
        Array("Basic.Needs.Patient", "患者Needs"), _
        Array("Basic.Needs.Family", "家族Needs"), _
        Array("Basic.LifeStatus", "生活状況"), _
        Array("Basic.PrimaryDx", "主診断"), _
        Array("Basic.CareLevel", "要介護度") _
    )

    For i = LBound(pairs) To UBound(pairs)
        headerNew = pairs(i)(0)
        headerOld = pairs(i)(1)

        ' 見出し列を取得（どちらか無ければスキップ）
        cNew = FindColByHeaderExact(ws, headerNew)
        cOld = FindColByHeaderExact(ws, headerOld)
        If cNew = 0 Or cOld = 0 Then GoTo NextPair

        vNew = ws.Cells(rowNum, cNew).value
        vOld = ws.Cells(rowNum, cOld).value

        sNew = Trim$(CStr(vNew))
        sOld = Trim$(CStr(vOld))

        ' どちらかだけ埋まっている場合、空いている方へコピー
        If sNew = "" And sOld <> "" Then
            ws.Cells(rowNum, cNew).value = vOld
        ElseIf sOld = "" And sNew <> "" Then
            ws.Cells(rowNum, cOld).value = vNew
        End If

NextPair:
    Next i

    Exit Sub

ErrHandler:
    Debug.Print "[WARN] MirrorBasicRow_Eval error at row=" & rowNum & _
                " : " & Err.Number & " " & Err.Description
End Sub






Sub Probe_ID_Candidates_Deep()
    Walk frmEval
End Sub

Private Sub Walk(ByVal parent As Object)
    Dim c As Object
    For Each c In parent.controls
        DumpIfIDLike c
        If HasControls(c) Then Walk c
    Next
End Sub

Private Function HasControls(ByVal o As Object) As Boolean
    On Error Resume Next
    Dim n As Long
    n = o.controls.count
    HasControls = (Err.Number = 0 And n > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub DumpIfIDLike(ByVal c As Object)
    Dim nm As String, tg As String, cap As String, val As String
    nm = LCase$(c.name)
    On Error Resume Next
    tg = LCase$(Trim$(c.tag))
    cap = LCase$(CStr(c.caption))
    val = CStr(c.value)
    On Error GoTo 0
    
    If InStr(nm, "id") > 0 Or InStr(tg, "id") > 0 Or InStr(cap, "id") > 0 Then
        Debug.Print TypeName(c) & "  name=" & c.name & "  tag=" & tg & "  caption=" & cap & "  value=" & val
    End If
End Sub





Private Function EvalIndexHeaders() As Variant
    EvalIndexHeaders = Array(HDR_USER_ID, HDR_NAME, HDR_KANA, HDR_SHEET, HDR_FIRST_EVAL, HDR_LATEST_EVAL, HDR_RECORD_COUNT)
End Function

Private Function EnsureEvalIndexSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EVAL_INDEX_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=Sheets(Sheets.count))
        ws.name = EVAL_INDEX_SHEET_NAME
    End If

    Dim headers As Variant: headers = EvalIndexHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = CStr(headers(i))
    Next i

    Set EnsureEvalIndexSheet = ws
End Function
Private Function BasicInfoLegacyHeaders() As Variant
    BasicInfoLegacyHeaders = Array( _
        "氏名", "フリガナ", "性別", "生年月日", "年齢", "住所", _
        "電話番号", "本人Needs", "家族Needs", "主病名", "要介護度", _
        "発症日", "既往歴", "高齢者の日常生活自立度", "認知症高齢者の日常生活自立度", _
        "評価日", "初回評価日", "経過", "備考", "要支援", "補助具", "生活状況")
End Function

Private Function CommonHistoryHeaders() As Variant

    Dim headers As Collection: Set headers = New Collection
    Dim v As Variant

    For Each v In Array( _
        HDR_ROWNO, "Basic.ID", "Basic.Name", "Basic.NameKana", _
        "Basic.EvalDate", "Basic.Evaluator", "Basic.EvaluatorJob", "Basic.Age", _
        "Basic.BirthDate", "Basic.Sex", "Basic.PrimaryDx", "Basic.OnsetDate", _
        "Basic.CareLevel", "Basic.DementiaADL", "Basic.LifeStatus", _
        "Basic.Needs.Patient", "Basic.Needs.Family", _
        "Basic.Medical.AdmitDate", "Basic.Medical.DischargeDate", _
        "Basic.Medical.CourseNote", "Basic.Medical.ComplicationNote", _
        HDR_HOMEENV_CHECKS, HDR_HOMEENV_NOTE, HDR_AIDS_CHECKS, HDR_RISK_CHECKS, _
         "IO_Cog_Memory", "IO_Cog_Attention", "IO_Cog_Orientation", _
        "IO_Cog_Judgment", "IO_Cog_Executive", "IO_Cog_Language", _
        "IO_Cog_DementiaType", "IO_Cog_DementiaNote", "IO_Cog_BPSD", _
        "IO_Mental_Mood", "IO_Mental_Motivation", "IO_Mental_Anxiety", _
        "IO_Mental_Relation", "IO_Mental_Sleep", "IO_Mental_Note")
        headers.Add CStr(v)
    Next v

    For Each v In BasicInfoLegacyHeaders()
        headers.Add CStr(v)
    Next v

    Dim arr() As String
    ReDim arr(0 To headers.count - 1) As String
    Dim i As Long
    For i = 1 To headers.count
        arr(i - 1) = CStr(headers(i))
    Next i
    CommonHistoryHeaders = arr
    
    End Function

Private Sub EnsureHistorySheetInitialized(ByVal ws As Worksheet)
    Dim headers As Variant: headers = CommonHistoryHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If FindColByHeaderExact(ws, CStr(headers(i))) = 0 Then
            ws.Cells(1, ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1).value = CStr(headers(i))
        End If
    Next i
End Sub

Private Function NextHistorySheetName(ByVal indexWs As Worksheet) As String
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 4).End(xlUp).row
    Dim maxNo As Long, r As Long, nm As String, n As Long
    For r = 2 To lastRow
        nm = CStr(indexWs.Cells(r, 4).value)
        If Left$(nm, Len(EVAL_HISTORY_SHEET_PREFIX)) = EVAL_HISTORY_SHEET_PREFIX Then
            On Error Resume Next
            n = CLng(Mid$(nm, Len(EVAL_HISTORY_SHEET_PREFIX) + 1))
            On Error GoTo 0
            If n > maxNo Then maxNo = n
        End If
    Next r
    NextHistorySheetName = EVAL_HISTORY_SHEET_PREFIX & Format$(maxNo + 1, "0000")
End Function

Private Function FindEvalIndexRowsByName(ByVal indexWs As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 2).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(NormalizeName(CStr(indexWs.Cells(r, 2).value)), NormalizeName(nameText), vbTextCompare) = 0 Then c.Add r
    Next r
    Set FindEvalIndexRowsByName = c
End Function



Private Function FindEvalIndexRowByUserID(ByVal indexWs As Worksheet, ByVal userID As String) As Long
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(indexWs.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then
            FindEvalIndexRowByUserID = r
            Exit Function
        End If
    Next r
End Function


Private Function FindEvalIndexRowBySheetName(ByVal indexWs As Worksheet, ByVal sheetName As String) As Long
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 4).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(indexWs.Cells(r, 4).value), sheetName, vbTextCompare) = 0 Then FindEvalIndexRowBySheetName = r: Exit Function
    Next r
End Function

Private Function FindEvalIndexRowsByNameWithoutUserID(ByVal indexWs As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim rowsByName As Collection
    Dim i As Long
    Dim rowNo As Long

    Set rowsByName = FindEvalIndexRowsByName(indexWs, nameText)
    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        If Len(Trim$(CStr(indexWs.Cells(rowNo, 1).value))) = 0 Then c.Add rowNo
    Next i

    Set FindEvalIndexRowsByNameWithoutUserID = c
End Function

Private Function BuildLegacyTransferCandidatesMessage(ByVal indexWs As Worksheet, ByVal rowsByName As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long
    Dim kanaVal As String
    Dim latestVal As String
    Dim recCount As String
    Dim sheetName As String

    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        kanaVal = Trim$(CStr(indexWs.Cells(rowNo, 3).value))
        sheetName = Trim$(CStr(indexWs.Cells(rowNo, 4).value))
        latestVal = Trim$(CStr(indexWs.Cells(rowNo, 6).value))
        recCount = Trim$(CStr(indexWs.Cells(rowNo, 7).value))

        lines = lines & CStr(i) & ") Sheet:" & sheetName
        If Len(kanaVal) > 0 Then lines = lines & " / Kana:" & kanaVal
        If Len(latestVal) > 0 Then lines = lines & " / Latest:" & latestVal
        If Len(recCount) > 0 Then lines = lines & " / Count:" & recCount
        If i < rowsByName.count Then lines = lines & vbCrLf
    Next i

    If Len(lines) = 0 Then Exit Function

    BuildLegacyTransferCandidatesMessage = _
        "同姓同名の利用者が複数存在します。" & vbCrLf & _
        "対象者を特定するため、IDを入力してください。" & vbCrLf & _
        "（または候補から選択してください）" & vbCrLf & vbCrLf & _
        lines
End Function

Private Function PickLegacyTransferIndexRow(ByVal indexWs As Worksheet, _
                                            ByVal rowsByName As Collection, _
                                            ByVal userID As String, _
                                            ByVal personName As String, _
                                            ByVal forSave As Boolean) As Long
    Dim prompt As String
    Dim picked As Variant
    Dim i As Long
    Dim n As Long
    Dim hasUnassigned As Boolean
    
    If rowsByName Is Nothing Then Exit Function
    If rowsByName.count = 0 Then Exit Function

    For i = 1 To rowsByName.count
        If Len(Trim$(CStr(indexWs.Cells(CLng(rowsByName(i)), 1).value))) = 0 Then
            hasUnassigned = True
            Exit For
        End If
    Next i

    prompt = BuildLegacyTransferCandidatesMessage(indexWs, rowsByName) & vbCrLf & vbCrLf
    If hasUnassigned Then
        prompt = prompt & "ID未設定の旧記録が見つかりました。" & vbCrLf
    Else
        prompt = prompt & "同姓同名の利用者が複数存在します。" & vbCrLf
    End If

    prompt = prompt & "対象者: " & personName & " / ID: " & userID & vbCrLf & _
         "引き継ぐ記録の番号を入力してください。"

    picked = Application.InputBox(prompt, "旧記録の引き継ぎ", Type:=1)
    If VarType(picked) = vbBoolean Then Exit Function
    If IsError(picked) Then Exit Function
    If Not IsNumeric(picked) Then Exit Function
    If Len(CStr(picked)) = 0 Then Exit Function

    n = CLng(picked)
    If n < 1 Or n > rowsByName.count Then Exit Function

    PickLegacyTransferIndexRow = CLng(rowsByName(n))
End Function

Private Function WriteUserIDToLegacyHistory(ByVal wsTarget As Worksheet, _
                                            ByVal userID As String, _
                                            ByVal personName As String) As Long
    Dim cID As Long
    Dim cName As Long
    Dim lastRow As Long
    Dim r As Long

    cID = FindColByHeaderExact(wsTarget, "Basic.ID")
    If cID = 0 Then cID = EnsureHeader(wsTarget, "Basic.ID")

    cName = FindColByHeaderExact(wsTarget, "Basic.Name")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "氏名")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "利用者名")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "Name")

    lastRow = LastDataRow(wsTarget)
    For r = 2 To lastRow
        If Len(Trim$(CStr(wsTarget.Cells(r, cID).value))) > 0 Then GoTo NextRow
        If cName > 0 Then
            If NormalizeName(CStr(wsTarget.Cells(r, cName).value)) <> NormalizeName(personName) Then GoTo NextRow
        End If
        wsTarget.Cells(r, cID).value = userID
        WriteUserIDToLegacyHistory = WriteUserIDToLegacyHistory + 1
NextRow:
    Next r
End Function

Private Function AssignUserIDToHistoryEntry(ByVal indexWs As Worksheet, _
                                            ByVal indexRow As Long, _
                                            ByVal userID As String, _
                                            ByVal personName As String, _
                                            ByVal kanaVal As String, _
                                            ByRef wsTarget As Worksheet) As Boolean
    Dim sheetName As String

    If indexRow <= 0 Then Exit Function

    If Len(CStr(indexWs.Cells(indexRow, 4).value)) = 0 Then indexWs.Cells(indexRow, 4).value = NextHistorySheetName(indexWs)
    sheetName = CStr(indexWs.Cells(indexRow, 4).value)

    indexWs.Cells(indexRow, 1).value = userID
    indexWs.Cells(indexRow, 2).value = personName
    If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal

    Set wsTarget = EnsureEvalSheet(sheetName)
    EnsureHistorySheetInitialized wsTarget
    Call WriteUserIDToLegacyHistory(wsTarget, userID, personName)

    AssignUserIDToHistoryEntry = True
End Function



Private Function ResolveUserHistorySheet(owner As Object, ByVal forSave As Boolean, ByRef wsTarget As Worksheet, ByRef message As String) As Boolean
    Dim nm As String: nm = Trim$(owner.txtName.text)
    If Len(nm) = 0 Then message = "氏名が未入力です": Exit Function

    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    Dim rowsByName As Collection
    Dim rowsByID As Collection
    Dim rowsByNameWithoutID As Collection
    Dim indexRow As Long
    Dim newRow As Long
    Dim pickedRow As Long

    Set rowsByName = FindEvalIndexRowsByName(indexWs, nm)
    HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                           "forSave=" & CStr(forSave), _
                           "name=" & HistoryLoadDebug_Quote(nm), _
                           "id=" & HistoryLoadDebug_Quote(idVal), _
                           "kana=" & HistoryLoadDebug_Quote(kanaVal), _
                           "rowsByName=" & CStr(rowsByName.count)
    
    
    If Len(idVal) = 0 And rowsByName.count = 1 Then
        indexRow = CLng(rowsByName(1))
        HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                               "branch=noID_uniqueName", _
                               "indexRow=" & CStr(indexRow), _
                               "indexSheetCellBefore=" & HistoryLoadDebug_Quote(CStr(indexWs.Cells(indexRow, 4).value))

 
        If Len(CStr(indexWs.Cells(indexRow, 4).value)) = 0 Then indexWs.Cells(indexRow, 4).value = NextHistorySheetName(indexWs)
        Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(indexRow, 4).value))
        EnsureHistorySheetInitialized wsTarget
        HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                               "branch=noID_uniqueName", _
                               "resolvedSheet=" & HistoryLoadDebug_SheetName(wsTarget), _
                               "sheetLastDataRow=" & CStr(LastDataRow(wsTarget))


        If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
        ResolveUserHistorySheet = True
        Exit Function
    End If

    
    If Len(idVal) > 0 Then
        Set rowsByID = FindEvalIndexRowsByUserID(indexWs, idVal)
        If rowsByID.count > 1 Then
            message = BuildDuplicateUserIDMessage(indexWs, idVal, rowsByID)
            Exit Function
        End If

        If rowsByID.count = 1 Then
            indexRow = CLng(rowsByID(1))

            Dim indexName As String: indexName = Trim$(CStr(indexWs.Cells(indexRow, 2).value))
            Dim indexKana As String: indexKana = Trim$(CStr(indexWs.Cells(indexRow, 3).value))
            If NormalizeName(indexName) <> NormalizeName(nm) _
               Or Not IsSameKanaIfAvailable(kanaVal, indexKana) Then
                message = BuildUserIdentityMismatchMessage(idVal, nm, indexName, kanaVal, indexKana)
                Exit Function
            End If

            If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
            If Len(CStr(indexWs.Cells(indexRow, 4).value)) = 0 Then indexWs.Cells(indexRow, 4).value = NextHistorySheetName(indexWs)

            Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(indexRow, 4).value))

            EnsureHistorySheetInitialized wsTarget
            ResolveUserHistorySheet = True
            Exit Function
        End If
        
    End If

    
    If rowsByName.count = 0 Then
        
        If Not forSave Then
           message = "対象の評価履歴が見つかりません。"
            Exit Function
        End If
        
        newRow = NextAppendRow(indexWs)
        indexWs.Cells(newRow, 1).value = idVal
        indexWs.Cells(newRow, 2).value = nm
        indexWs.Cells(newRow, 3).value = kanaVal
        indexWs.Cells(newRow, 4).value = NextHistorySheetName(indexWs)
        Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(newRow, 4).value))
        
        ResolveUserHistorySheet = True
        Exit Function
    
    End If

    If Len(idVal) > 0 Then
        Set rowsByNameWithoutID = FindEvalIndexRowsByNameWithoutUserID(indexWs, nm)
        pickedRow = PickLegacyTransferIndexRow(indexWs, rowsByNameWithoutID, idVal, nm, forSave)
        If pickedRow > 0 Then
            If AssignUserIDToHistoryEntry(indexWs, pickedRow, idVal, nm, kanaVal, wsTarget) Then
                ResolveUserHistorySheet = True
                Exit Function
            End If
        End If

        If forSave Then
            newRow = NextAppendRow(indexWs)
            indexWs.Cells(newRow, 1).value = idVal
            indexWs.Cells(newRow, 2).value = nm
            indexWs.Cells(newRow, 3).value = kanaVal
            indexWs.Cells(newRow, 4).value = NextHistorySheetName(indexWs)
            Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(newRow, 4).value))
            EnsureHistorySheetInitialized wsTarget
            ResolveUserHistorySheet = True
            Exit Function
        End If

        message = "同姓同名の利用者が複数存在します。" & vbCrLf & _
          "該当する履歴を選択してください。"
        If Not rowsByNameWithoutID Is Nothing Then
            If rowsByNameWithoutID.count > 0 Then
                message = message & vbCrLf & vbCrLf & BuildLegacyTransferCandidatesMessage(indexWs, rowsByNameWithoutID)
            End If
        End If
        Exit Function
    End If

    message = "同姓同名の利用者が複数いるため、IDまたは履歴を選択してください。" & _
          BuildDuplicateNameCandidatesMessage(indexWs, rowsByName)
End Function

Private Function BuildDuplicateNameCandidatesMessage(ByVal indexWs As Worksheet, ByVal rowsByName As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long
    Dim idVal As String
    Dim kanaVal As String
    Dim latestVal As String

    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        idVal = Trim$(CStr(indexWs.Cells(rowNo, 1).value))
        kanaVal = Trim$(CStr(indexWs.Cells(rowNo, 3).value))
        latestVal = Trim$(CStr(indexWs.Cells(rowNo, 6).value))
        lines = lines & "- ID: " & idVal & " / かな: " & kanaVal & " / 最新: " & latestVal
        If i < rowsByName.count Then lines = lines & vbCrLf
    Next i

    If Len(lines) > 0 Then
                BuildDuplicateNameCandidatesMessage = vbCrLf & vbCrLf & ":" & vbCrLf & lines & _
            vbCrLf & vbCrLf & "新規の場合は次のIDを使用できます:" & vbCrLf & _
            BuildNextAvailableUserIDCandidate(indexWs)
    End If
End Function

Private Function BuildNextAvailableUserIDCandidate(ByVal indexWs As Worksheet) As String
    Dim lastRow As Long
    Dim r As Long
    Dim rawID As String
    Dim idNum As Long
    Dim maxID As Long
    Dim maxDigits As Long
    Dim hasNumericID As Boolean

    lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        rawID = Trim$(CStr(indexWs.Cells(r, 1).value))
        If TryParseNumericUserID(rawID, idNum) Then
            hasNumericID = True
            If idNum > maxID Then maxID = idNum
            If Len(rawID) > maxDigits Then maxDigits = Len(rawID)
        End If
    Next r

    If Not hasNumericID Then
        BuildNextAvailableUserIDCandidate = "001"
        Exit Function
    End If

    If maxDigits < 3 Then maxDigits = 3
    BuildNextAvailableUserIDCandidate = Format$(maxID + 1, String$(maxDigits, "0"))
End Function

Private Function TryParseNumericUserID(ByVal rawID As String, ByRef parsedID As Long) As Boolean
    Dim numericValue As Double

    If Len(rawID) = 0 Then Exit Function
    If Not IsNumeric(rawID) Then Exit Function

    On Error GoTo EH
    numericValue = CDbl(rawID)
    If numericValue < 0 Then Exit Function
    If numericValue <> Fix(numericValue) Then Exit Function
    If numericValue > CLng(&H7FFFFFFF) Then Exit Function

    parsedID = CLng(numericValue)
    TryParseNumericUserID = True
    Exit Function
EH:
    TryParseNumericUserID = False
End Function

Private Sub HistoryLoadDebug_Print(ParamArray args())
    If Not HISTORY_LOAD_DEBUG Then Exit Sub

    Dim i As Long
    Dim msg As String
    For i = LBound(args) To UBound(args)
        msg = msg & IIf(i > 0, " | ", "") & CStr(args(i))
    Next i
    Debug.Print Format$(Now, "hh:nn:ss"), msg
End Sub

Private Function HistoryLoadDebug_Quote(ByVal valueText As String) As String
    HistoryLoadDebug_Quote = Chr$(34) & Replace$(CStr(valueText), Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function HistoryLoadDebug_SheetName(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        HistoryLoadDebug_SheetName = "(Nothing)"
    Else
        HistoryLoadDebug_SheetName = ws.name
    End If
End Function

Private Sub HistoryLoadDebug_ScanWorkbookForName(ByVal targetName As String, ByVal resolvedSheet As Worksheet)
    Dim ws As Worksheet
    Dim c As Long
    Dim lastRow As Long
    Dim r As Long
    Dim rowName As String
    Dim normalizedTarget As String
    Dim matched As Boolean

    normalizedTarget = NormalizeName(targetName)
    HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                           "targetName=" & HistoryLoadDebug_Quote(targetName), _
                           "resolvedSheet=" & HistoryLoadDebug_SheetName(resolvedSheet)

    For Each ws In ThisWorkbook.Worksheets
        If (Left$(ws.name, Len(EVAL_HISTORY_SHEET_PREFIX)) = EVAL_HISTORY_SHEET_PREFIX) _
           Or StrComp(ws.name, EVAL_SHEET_NAME, vbTextCompare) = 0 Then
            c = FindColByHeaderExact(ws, "Basic.Name")
                 If c = 0 Then c = FindHeaderCol(ws, "氏名")
                 If c = 0 Then c = FindHeaderCol(ws, "利用者名")
                 If c = 0 Then c = FindHeaderCol(ws, "Name")

            If c = 0 Then
                HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                       "sheet=" & ws.name, _
                                       "nameHeaderMissing=True"
            Else
                lastRow = LastDataRow(ws)
                matched = False
                For r = 2 To lastRow
                    rowName = CStr(ws.Cells(r, c).value)
                    If NormalizeName(rowName) = normalizedTarget Then
                        HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                               "sheet=" & ws.name, _
                                               "row=" & CStr(r), _
                                               "raw=" & HistoryLoadDebug_Quote(rowName), _
                                               "isResolvedSheet=" & CStr(StrComp(ws.name, HistoryLoadDebug_SheetName(resolvedSheet), vbTextCompare) = 0)
                        matched = True
                    End If
                Next r
                If Not matched Then
                    HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                           "sheet=" & ws.name, _
                                           "matchedRow=0", _
                                           "nameCol=" & CStr(c), _
                                           "lastRow=" & CStr(lastRow)
                End If
            End If
        End If
    Next ws
End Sub


Private Sub UpdateEvalIndexMetadata(ByVal owner As Object, ByVal indexRow As Long, ByVal sheetName As String)
    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    If Len(Trim$(owner.txtName.text)) > 0 Then indexWs.Cells(indexRow, 2).value = Trim$(owner.txtName.text)
    If Len(idVal) > 0 And Len(CStr(indexWs.Cells(indexRow, 1).value)) = 0 Then indexWs.Cells(indexRow, 1).value = idVal
    If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
    indexWs.Cells(indexRow, 4).value = sheetName
End Sub

Private Function TryParseEvalDate(ByVal v As Variant, ByRef normalizedDate As Date) As Boolean
    On Error GoTo EH
    If IsDate(v) Then normalizedDate = DateValue(CDate(v)): TryParseEvalDate = True
    Exit Function
EH:
    TryParseEvalDate = False
End Function

Private Function GetLatestValidEvalRow(ByVal ws As Worksheet) As Long
    Dim cEval As Long: cEval = FindColByHeaderExact(ws, "Basic.EvalDate")
    If cEval = 0 Then Exit Function
    Dim lastRow As Long: lastRow = LastDataRow(ws)
    Dim r As Long, d As Date
    For r = lastRow To 2 Step -1
        If TryParseEvalDate(ws.Cells(r, cEval).value, d) Then GetLatestValidEvalRow = r: Exit Function
    Next r
End Function

Public Sub GetUserEvalDateStats(ByVal wsTarget As Worksheet, _
                                ByRef firstEvalDate As String, _
                                ByRef latestEvalDate As String, _
                                ByRef previousEvalDate As String, _
                                ByRef recordCount As Long)
    Dim cEval As Long: cEval = FindColByHeaderExact(wsTarget, "Basic.EvalDate")
    If cEval = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = LastDataRow(wsTarget)
    Dim r As Long, d As Date
    Dim firstD As Date, latestD As Date, prevD As Date
    Dim hasFirst As Boolean, hasLatest As Boolean, hasPrev As Boolean

    For r = 2 To lastRow
        If TryParseEvalDate(wsTarget.Cells(r, cEval).value, d) Then
            recordCount = recordCount + 1

            If (Not hasFirst) Or d < firstD Then
                firstD = d
                hasFirst = True
            End If

            If (Not hasLatest) Or d > latestD Then
                If hasLatest And d <> latestD Then
                    If (Not hasPrev) Or latestD > prevD Then
                        prevD = latestD
                        hasPrev = True
                    End If
                End If
                latestD = d
                hasLatest = True
            ElseIf d < latestD Then
                If (Not hasPrev) Or d > prevD Then
                    prevD = d
                    hasPrev = True
                End If
            End If
        End If
    Next r

    If hasFirst Then firstEvalDate = Format$(firstD, "yyyy/mm/dd")
    If hasLatest Then latestEvalDate = Format$(latestD, "yyyy/mm/dd")
    If hasPrev Then previousEvalDate = Format$(prevD, "yyyy/mm/dd")
End Sub

Public Function GetPreviousEvalDateText(ByVal wsTarget As Worksheet) As String
    Dim firstEvalDate As String, latestEvalDate As String, previousEvalDate As String
    Dim recordCount As Long
    GetUserEvalDateStats wsTarget, firstEvalDate, latestEvalDate, previousEvalDate, recordCount
    GetPreviousEvalDateText = previousEvalDate
End Function

Private Sub UpdateEvalIndexStats(ByVal indexRow As Long, ByVal wsTarget As Worksheet)
    Dim firstEvalDate As String, latestEvalDate As String, previousEvalDate As String
    Dim recordCount As Long

    GetUserEvalDateStats wsTarget, firstEvalDate, latestEvalDate, previousEvalDate, recordCount

    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    
    indexWs.Cells(indexRow, 5).value = firstEvalDate
    indexWs.Cells(indexRow, 6).value = latestEvalDate
    indexWs.Cells(indexRow, 7).value = recordCount
End Sub
