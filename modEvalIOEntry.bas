Attribute VB_Name = "modEvalIOEntry"
'=== modEvalIOEntry : 評価フォーム IO ハブ ============================
' 役割：
'   - frmEval から EvalData シートへの保存／読込のハブ
'   - 各セクション IO モジュール（ROM / 姿勢 / MMT / 感覚・筋緊張・疼痛 /
'     ADL / 認知・精神 / テスト評価 / 日々の記録）をここから呼び出す
'   - EvalData の行決定（新規行 / 既存行）と、ID / BasicInfo の管理
'
' このモジュールが「知ってよい」こと：
'   - EvalData のヘッダ名・列番号（HeaderCol_Compat / Module2 / modHeaderMap 経由）
'   - frmEval を owner As Object として扱うこと（ただしレイアウト詳細は知らない）
'
' このモジュールが「やってはいけない」こと：
'   - フォームのレイアウト変更（Left/Top/Width/Height の書き換え）
'   - タブ構造の生成・破壊（MultiPage.Pages.Add など）
'   - コントロールの新規作成や削除
'   - EvalData 以外のシート IO（他シートの書き換え）
'
' 今後のリファクタ方針：
'   - Save/Load の入口は原則ここに集約する
'   - FromSheet 系は modEvalIOEntry および専用 IO モジュールからのみ呼び出す
'   - UI レイアウト系の処理は専用 Layout モジュールへ徐々に退避していく
'====================================================================




Option Explicit

Public Const EVAL_SHEET_NAME As String = "EvalData"
Public mDailyLogManual As Boolean    ' 日々の記録の手動保存フラグ



' === 補助具/リスク フレーム名（固定用） ===
Private Const FRM_AIDS As String = "Frame33"
Private Const FRM_RISK As String = "Frame34"
Private Const IO_TRACE As Boolean = False



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
         nm = Trim$(owner.txtName.Text)

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
If r <= 0 And Len(nm) > 0 Then
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
    Dim ws As Worksheet: Set ws = EnsureEvalSheet(EVAL_SHEET_NAME)
    Dim nm As String: nm = Trim$(owner.txtName.Text)
    Dim r As Long
    Dim r2 As Long

    If Len(nm) = 0 Then
        MsgBox "氏名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    
    r = FindLatestRowByName(ws, nm)   ' ★この1行を追加
    
    Debug.Print "r=" & r

    

    ' --- 同姓同名回避：IDが入っていて、同名が複数ある時だけ ID を使う ---
Dim idVal As String
idVal = Trim$(GetID_FromBasicInfo(owner))

If Len(idVal) > 0 Then
    r2 = FindLatestRowByNameAndID(ws, nm, idVal)
    If r2 > 0 Then r = r2
End If





    ' ★ここに追加（誤読込ガード）
    Dim cName As Long
    cName = FindColByHeaderExact(ws, "氏名")
    If cName = 0 Then cName = FindColByHeaderExact(ws, "利用者名")
    If cName = 0 Then cName = FindColByHeaderExact(ws, "名前")
    If cName = 0 Then
        MsgBox "氏名列が見つかりません。", vbExclamation
        Exit Sub
    End If
    If StrComp(NormalizeName(CStr(ws.Cells(r, cName).value)), NormalizeName(nm), vbTextCompare) <> 0 Then
        MsgBox "選択行の氏名が入力名と一致しません。読み込みを中止します。", vbExclamation
        Exit Sub
    End If
    ' ★ここまで

    t "[ENTRY] Load by NAME", ws.name, "row", r
    LoadAllSectionsFromSheet ws, r, owner
End Sub


' 下から遡って氏名一致の最新行を返す（見出しは「氏名」「利用者名」「名前」を順に探す）
Public Function FindLatestRowByName(ws As Worksheet, nameText As String) As Long

    Dim c As Long
    c = FindHeaderCol(ws, "氏名")
    If c = 0 Then c = FindHeaderCol(ws, "利用者名")
    If c = 0 Then c = FindHeaderCol(ws, "名前")
    If c = 0 Then Exit Function

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, c).End(xlUp).row
    Dim r As Long
    For r = lastRow To 2 Step -1      ' 1行目は見出し想定
        If NormalizeName(CStr(ws.Cells(r, c).value)) = NormalizeName(nameText) Then
            FindLatestRowByName = r
            Exit Function
        End If
    Next r
End Function



Public Function CountRowsByName(ws As Worksheet, nameText As String) As Long
    Dim c As Long
    c = FindHeaderCol(ws, "氏名")
    If c = 0 Then c = FindHeaderCol(ws, "利用者名")
    If c = 0 Then c = FindHeaderCol(ws, "名前")
    If c = 0 Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, c).End(xlUp).row

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
    lastRow = ws.Cells(ws.rows.Count, cName).End(xlUp).row

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
        Set EnsureEvalSheet = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
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
    Dim ws As Worksheet: Set ws = EnsureEvalSheet(EVAL_SHEET_NAME)
    Dim r As Long: r = NextAppendRow(ws)
    'r = WorksheetFunction.Max(ws.Cells(ws.rows.Count, 156).End(xlUp).row, ws.Cells(ws.rows.Count, 157).End(xlUp).row) + 1


    t "[ENTRY] Save to", ws.name, "row", r
    ' ★変更点のみ保存（chkDiffOnly=ONなら前回値を事前コピー）
Dim nm As String: nm = Trim$(owner.txtName.Text)
If Len(nm) = 0 Then MsgBox "氏名を入力してから保存してください。", vbExclamation: Exit Sub

Dim diffOnly As Boolean
On Error Resume Next
diffOnly = CBool(owner.Controls("chkDiffOnly").value)  ' 無ければ False のまま
On Error GoTo 0

If False Then ' diffOnly And Len(nm) > 0 Then
    Dim rOld As Long: rOld = FindLatestRowByName(ws, nm)
    If rOld > 0 Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).value = _
            ws.Range(ws.Cells(rOld, 1), ws.Cells(rOld, lastCol)).value
    End If
End If

   ' ★氏名セルを必ず現在入力で上書き（前回コピーの取り違い防止）
Dim cName As Long
cName = FindColByHeaderExact(ws, "氏名"): If cName = 0 Then cName = FindColByHeaderExact(ws, "利用者名"): If cName = 0 Then cName = FindColByHeaderExact(ws, "名前")
If cName > 0 Then ws.Cells(r, cName).value = nm

 ws.Cells(r, 1).value = r
  
    SaveAllSectionsToSheet ws, r, owner
    t "[ENTRY] Save done"
    
    
        Save_CognitionMental_AtRow ws, r, frmEval
        'Save_DailyLog_FromForm owner
        
        Call MirrorBasicRow(ws, r)

    
End Sub

Private Sub LoadEvaluation_LastRow_From_OBSOLETE(owner As Object)
    MsgBox "この入口は廃止しました。読み込みは『名前→直近候補から選択』に統一しています。", vbInformation
End Sub




' ====== 基本情報の保存/読込（このモジュール内） ======

' 見出しの列を取得（無ければ新規作成）
Private Function EnsureHeaderCol(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Then
        EnsureHeaderCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + IIf(ws.Cells(1, 1).value <> "", 1, 0)
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
    Set c = owner.Controls("frHeader").Controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.Controls("txtHdrKana")
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
    Set c = owner.Controls("frHeader").Controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.Controls("txtHdrKana")
        On Error GoTo 0
    End If

    If c Is Nothing Then Exit Sub

    On Error Resume Next
    c.value = CStr(v)
    On Error GoTo 0
End Sub

' 汎用：コンボを安全にセット（リストにある時だけ選択）
Private Sub SetComboSafe_Basic(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cB As MSForms.ComboBox
    Dim s As String, i As Long, hit As Long
    s = CStr(v)
    Set cB = FindCtlDeep(owner, ctlName)
    If cB Is Nothing Then Exit Sub
    hit = -1
    For i = 0 To cB.ListCount - 1
        If CStr(cB.List(i)) = s Then hit = i: Exit For
    Next
    If hit >= 0 Then cB.ListIndex = hit Else cB.ListIndex = -1
End Sub

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

    
    '--- 単一値のマッピング（最後の要素に _ を付けない） ---
    Dim map As Variant
map = Array( _
    Array("評価日", "txtEDate"), _
    Array("年齢", "txtAge"), _
    Array("性別", "cboSex"), _
    Array("Basic.Name", "txtName"), _
    Array("評価者", "txtEvaluator"), _
    Array("発症日", "txtOnset"), _
    Array("患者Needs", "txtNeedsPt"), _
    Array("家族Needs", "txtNeedsFam"), _
    Array("生活状況", "txtLiving"), _
    Array("主診断", "txtDx"), _
    Array("要介護度", "cboCare"), _
    Array("障害高齢者の日常生活自立度", "cboElder"), _
    Array("認知症高齢者の日常生活自立度", "cboDementia") _
    )


    '--- 既存のループ：単一値を書き込み ---
    Dim i As Long, head As String, ctl As String, c As Long, v As String
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0)):  ctl = CStr(map(i)(1))
        v = GetCtlTextGeneric(owner, ctl)
        If Len(v) > 0 Then
            c = FindColByHeaderExact(ws, head): If c = 0 Then c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1: ws.Cells(1, c).value = head
            ws.Cells(r, c).value = v
            Debug.Print "[BASIC][SAVE]", head, "->", v
        End If
    Next i

    c = EnsureHeader(ws, "Basic.NameKana")
    ws.Cells(r, c).value = GetHdrKanaText(owner)
    Debug.Print "[BASIC][SAVE] Basic.NameKana ->", CStr(ws.Cells(r, c).value)
    
    Dim idVal As String: idVal = GetID_FromBasicInfo(owner)
    If Len(idVal) > 0 Then ws.Cells(r, EnsureHeader(ws, "Basic.ID")).value = idVal


    '--- ここから追記：チェック群のCSV保存（補助具／リスク）※ループの“後ろ” ---
    Dim s As String
    c = EnsureHeader(ws, "補助具")
s = SerializeChecks(owner, "Frame33", True)
Debug.Print "[BASIC][SAVE] 補助具 ->", s, " @col=", c
ws.Cells(r, c).value = s

   c = EnsureHeader(ws, "リスク")
s = SerializeChecks(owner, "Frame34", False)
Debug.Print "[BASIC][SAVE] リスク ->", s, " @col=", c
ws.Cells(r, c).value = s


    
    
    
    
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
        Array("性別", "cboSex"), _
        Array("氏名", "txtName"), _
        Array("評価者", "txtEvaluator"), _
        Array("発症日", "txtOnset"), _
        Array("患者Needs", "txtNeedsPt"), _
        Array("家族Needs", "txtNeedsFam"), _
        Array("生活状況", "txtLiving"), _
        Array("主診断", "txtDx"), _
        Array("要介護度", "cboCare"), _
        Array("障害高齢者の日常生活自立度", "cboElder"), _
        Array("認知症高齢者の日常生活自立度", "cboDementia") _
    )

    '--- 単一値をフォームへ読込 ---
    Dim i As Long, head As String, ctl As String, c As Long, v As Variant
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0))
        ctl = CStr(map(i)(1))

        c = FindHeaderCol(ws, head)
        If c > 0 Then
            v = ws.Cells(r, c).value
            If Left$(ctl, 3) = "cbo" Then
                SetComboSafely owner, ctl, CStr(v)
            Else
                Dim o As Object: Set o = FindCtlDeep(owner, ctl)
                If Not o Is Nothing Then o.value = v
            End If
            
        End If
    Next i

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


If GetBool(owner, "chkLoadParalysis", True) Then Call IO_SafeRunLoad("LoadParalysisFromSheet", ws, r, owner)
If GetBool(owner, "chkLoadROM", True) Then Call IO_SafeRunLoad("LoadROMFromSheet", ws, r, owner)
Debug.Print "[TRACE] About to run POSTURE"
If GetBool(owner, "chkLoadPosture", True) Then Call IO_SafeRunLoad("LoadPostureFromSheet", ws, r, owner)
Debug.Print "[TRACE] Done POSTURE"

Debug.Print "[TRACE] About to run MMT"
Call MMT.LoadMMTFromSheet(ws, r, owner)
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
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
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
        idCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If
    If Len(idVal) = 0 Then Err.Raise 5, , "IDが空です。"

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, idCol).End(xlUp).row
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
    GetID_FromBasicInfo = Trim$(CStr(owner.Controls("frHeader").Controls("txtHdrPID").value))
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
    HasControls = Not (frm.Controls Is Nothing)
    On Error GoTo 0
    If Not HasControls Then Exit Function

    ' --- 以下は今のロジックそのまま ---
    Dim lb As Object, ctl As Object
    For Each ctl In frm.Controls
        If TypeName(ctl) = "Label" Then
            If InStr(1, CStr(ctl.caption), labelCaption, vbTextCompare) > 0 Then
                Set lb = ctl: Exit For
            End If
        End If
    Next ctl
    If lb Is Nothing Then Exit Function

    Dim best As Object, bestScore As Double
    bestScore = 1E+20
    For Each ctl In frm.Controls
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

    Dim tmp As Object: Set tmp = container.Controls
    If Err.Number <> 0 Then Err.Clear: Exit Function

    Dim ctl As Object
    For Each ctl In container.Controls
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
                Set tmp = ctl.Controls
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

    ' 1) 既存ヘッダをマージ改名
    ApplyAliasesMerge_Basic ws, d

    ' 2) 最低限必要な列がなければ追加（Save/Loadの対象を漏れなく）
    Dim need As Variant, mustHave As Variant
    mustHave = Array( _
        "Basic.ID", "Basic.Name", "Basic.EvalDate", "Basic.Evaluator", _
        "Basic.Age", "Basic.Sex", "Basic.PrimaryDx", "Basic.OnsetDate", _
        "Basic.CareLevel", "Basic.DementiaADL", "Basic.LifeStatus", _
        "Basic.Needs.Patient", "Basic.Needs.Family" _
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
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim j As Long
    For j = lastCol To 1 Step -1
        Dim h As String: h = Trim$(CStr(ws.Cells(1, j).value))
        If Len(h) = 0 Then GoTo NextJ
        If d.exists(h) Then
            Dim dst As String: dst = CStr(d(h))
            Dim dstCol As Long: dstCol = modSchema.FindColByHeaderExact(ws, dst)
            If dstCol > 0 And dstCol <> j Then
                ' マージ（空欄だけ埋める）
                Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, j).End(xlUp).row
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
        Dim lc As Long: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
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
        idCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, idCol).End(xlUp).row
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
    Dim cB As Object  ' MSForms.ComboBox を late binding で扱う
    Dim i As Long, hit As Long
    Dim s As String

    On Error Resume Next
    Set cB = FindCtlDeep(owner, ctlName)
    On Error GoTo 0
    If cB Is Nothing Then Exit Sub

    s = CStr(v)
    hit = -1
    For i = 0 To cB.ListCount - 1
        If CStr(cB.List(i)) = s Then
            hit = i
            Exit For
        End If
    Next

    If hit >= 0 Then
        cB.ListIndex = hit               ' 一致が見つかったら選択
    Else
        cB.ListIndex = -1                ' 見つからなければ未選択に（DropDownListでも安全）
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
    Set hit = parent.Controls(targetName)
    On Error GoTo 0
    If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function

    ' 4) 子コントロールを再帰走査（Controls を持たない型はスキップ）
    On Error Resume Next
    For Each c In parent.Controls
        Err.Clear
        Set hit = FindControlDeep(c, targetName)
        If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function
    Next c
    On Error GoTo 0
End Function


' 代表キャプションから親フレームを推定
Private Function FindGroupByAnyCaption(frm As Object, captions As Variant) As Object
    Dim cont As Object, c As Object, cap As Variant
    For Each cont In frm.Controls
        On Error Resume Next
        ' コンテナ（Frame/Pageなど）だけ調べる
        If Not cont.Controls Is Nothing Then
            For Each c In cont.Controls
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
    Set ResolveGroup = frm.Controls(targetName)
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
    For Each c In grp.Controls
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
    For Each c In grp.Controls
        If TypeName(c) = "CheckBox" Then
            c.value = dict.exists(Trim$(c.caption))
        End If
    Next
End Sub

' IDの最大値+1
Public Function NextID(ws As Worksheet, ByVal cID As Long) As Long
    Dim last As Long: last = ws.Cells(ws.rows.Count, cID).End(xlUp).row
    If last < 2 Then NextID = 1: Exit Function
    On Error Resume Next
    NextID = WorksheetFunction.Max(ws.Range(ws.Cells(2, cID), ws.Cells(last, cID))) + 1
    If Err.Number <> 0 Then NextID = 1: Err.Clear
    On Error GoTo 0
End Function


Private Function GetBool(owner As Object, ctlName As String, Optional defaultValue As Boolean = True) As Boolean
    On Error Resume Next
    GetBool = CBool(owner.Controls(ctlName).value)
    If Err.Number <> 0 Then GetBool = defaultValue
    On Error GoTo 0
End Function




'=== Compat: SENSE_IO を IO_Sensory にミラー（行 r のみ） ===
Private Sub Mirror_SensoryIO(ws As Worksheet, ByVal r As Long)
    Dim cSrc As Variant, cDst As Long
    cSrc = Application.Match("SENSE_IO", ws.rows(1), 0)
    If IsError(cSrc) Then Exit Sub

    ' 宛先ヘッダ IO_Sensory を確保
    Dim M As Variant, lastCol As Long
    M = Application.Match("IO_Sensory", ws.rows(1), 0)
    If IsError(M) Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = "IO_Sensory"
        cDst = lastCol + 1
    Else
        cDst = CLng(M)
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
        Debug.Print "SENSE_IO(col" & cSense & ") =", ws.Cells(r, cSense).Text
    Else
        Debug.Print "SENSE_IO: <no header>"
    End If

    If Not IsError(cADL) Then
        Debug.Print "IO_ADL(col" & cADL & ") =", ws.Cells(r, cADL).Text
    Else
        Debug.Print "IO_ADL: <no header>"
    End If

    If Not IsError(cIOSense) Then
        Debug.Print "IO_Sensory(col" & cIOSense & ") =", ws.Cells(r, cIOSense).Text
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

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

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
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

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
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

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
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

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
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

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

    ' TODO: ここから下は後で実装（今は触らない）
    ' IO_TestEval を分解して
    ' owner（frmEval）の txtTenMWalk / txtTUG / txtFiveSts /
    ' txtGripR / txtGripL / txtSemi に流し込む
    
    
    ws.Cells(r, 181).value = val(owner.txtTUG.value)

    
   
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
        For Each c In owner.Controls
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
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 安定性チェック（chkWalkStab_*）を一度全部OFF ---
    For Each c In owner.Controls
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
    For Each c In owner.Controls
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
                    For Each c In owner.Controls
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
            For Each c In owner.Controls
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
    For Each c In owner.Controls
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
    For Each ctl In parent.Controls
        If StrComp(ctl.name, name, vbTextCompare) = 0 Then
            Set FindControlRecursive = ctl
            Exit Function
        End If
        ' Frame や MultiPage の場合は再帰検索
        On Error Resume Next
        If ctl.Controls.Count > 0 Then
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


Public Function Build_WalkIndep_IO(owner As Object) As String
    Dim vDist As String
    Dim vOut As String
    Dim vSpeed As String
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
    For Each c In owner.Controls
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
    On Error Resume Next
    vDist = Trim$(owner.Controls("cmbWalkDistance").value)
    vOut = Trim$(owner.Controls("cmbWalkOutdoor").value)
    vSpeed = Trim$(owner.Controls("cmbWalkSpeed").value)
    On Error GoTo 0

    ' 安定性チェック（chkWalkStab_～ を全部拾う）
    Set hits = New Collection
    For Each c In owner.Controls
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
    For i = 1 To hits.Count
        If i > 1 Then stab = stab & "/"
        stab = stab & hits(i)
    Next i

        ' IO 文字列組み立て
    s = "Walk_IndepLevel=" & vLevel
    s = s & "|Walk_Distance=" & vDist
    s = s & "|Walk_Outdoor=" & vOut
    s = s & "|Walk_Stab=" & stab
    s = s & "|Walk_Speed=" & vSpeed

    

    Build_WalkIndep_IO = s
End Function



Public Function Build_WalkAbn_IO(owner As Object) As String
    Dim c As Object
    Dim hits As Collection
    Dim s As String
    Dim nm As String
    
    Set hits = New Collection
    
    For Each c In owner.Controls
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
    If hits.Count = 0 Then
        Build_WalkAbn_IO = ""
        Exit Function
    End If
    
    ' fraWalkAbn_A_chk0|fraWalkAbn_A_chk3|… という形で連結
    Dim i As Long
    For i = 1 To hits.Count
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
        For Each c In owner.Controls
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
        If probs.Count > 0 Then
            For i = 1 To probs.Count
                If i > 1 Then probsStr = probsStr & "/"
                probsStr = probsStr & probs(i)
            Next i
        End If

        ' --- レベル（OptionButton, GroupName=phase）を拾う ---
        For Each c In owner.Controls
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




Public Sub Save_CognitionMental_AtRow(ws As Worksheet, r As Long, owner As Object)
    Dim frm As Object
    Dim col As Long
    Dim v As Variant
    Dim f As MSForms.Frame
    Dim c As MSForms.Control
    Dim bpsd As String
    
    Set frm = owner   ' frmEval を受け取る想定
    
    '=== 認知：中核6項目 =====================================
    
    ' 記憶
    col = HeaderCol_Compat("IO_Cog_Memory", ws)
    If col > 0 Then
        v = frm.Controls("Frame31").Controls("mpCogMental") _
        .Pages("pgCognition").Controls("cmbCogMemory").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 注意
    col = HeaderCol_Compat("IO_Cog_Attention", ws)
    If col > 0 Then
        v = frm.Controls("Frame31").Controls("mpCogMental") _
        .Pages("pgCognition").Controls("cmbCogAttention").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 見当識
    col = HeaderCol_Compat("IO_Cog_Orientation", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("cmbCogOrientation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 判断
    col = HeaderCol_Compat("IO_Cog_Judgement", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("cmbCogJudgement").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 遂行機能
    col = HeaderCol_Compat("IO_Cog_Executive", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("cmbCogExecutive").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 言語
    col = HeaderCol_Compat("IO_Cog_Language", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("cmbCogLanguage").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    '=== 認知：認知症の種類＋備考 ==============================
    
    col = HeaderCol_Compat("IO_Cog_DementiaType", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("cmbDementiaType").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    col = HeaderCol_Compat("IO_Cog_DementiaNote", ws)
    If col > 0 Then
           v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgCognition").Controls("txtDementiaNote").Text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
        '=== 認知：BPSD（チェックが入っている項目を | 区切りで保存） ===
    
    bpsd = ""
    With frm.Controls("Frame31").Controls("mpCogMental").Pages("pgCognition")
        For Each c In .Controls
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
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("cmbMood").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 意欲
    col = HeaderCol_Compat("IO_Mental_Motivation", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("cmbMotivation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 不安
    col = HeaderCol_Compat("IO_Mental_Anxiety", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("cmbAnxiety").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 対人関係
    col = HeaderCol_Compat("IO_Mental_Relation", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("cmbRelation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 睡眠
    col = HeaderCol_Compat("IO_Mental_Sleep", ws)
    If col > 0 Then
            v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("cmbSleep").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 精神面・備考
    col = HeaderCol_Compat("IO_Mental_Note", ws)
    If col > 0 Then
           v = frm.Controls("Frame31").Controls("mpCogMental") _
            .Pages("pgMental").Controls("txtMentalNote").Text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
End Sub

   



Public Sub Load_CognitionMental_FromRow(ws As Worksheet, ByVal r As Long, owner As Object)
    Const COL_COG_MEMORY       As Long = 165
    Const COL_COG_ATTENTION    As Long = 166
    Const COL_COG_ORIENTATION  As Long = 167
    Const COL_COG_JUDGEMENT    As Long = 168
    Const COL_COG_EXECUTIVE    As Long = 169
    Const COL_COG_LANGUAGE     As Long = 170
    Const COL_COG_DEMENTIA     As Long = 171
    Const COL_COG_DEM_NOTE     As Long = 172
    Const COL_COG_BPSD         As Long = 173
    Const COL_MENTAL_MOOD      As Long = 174
    Const COL_MENTAL_MOTIV     As Long = 175
    Const COL_MENTAL_ANXIETY   As Long = 176
    Const COL_MENTAL_RELATION  As Long = 177
    Const COL_MENTAL_SLEEP     As Long = 178
    Const COL_MENTAL_NOTE      As Long = 179

    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pgCog As MSForms.Page
    Dim pgMental As MSForms.Page

    Dim v As Variant
    Dim s As String
    Dim arr() As String
    Dim i As Long, j As Long
    Dim chk As MSForms.CheckBox

    '=== UI ルート取得（絶対名を前提）===
    Set f = owner.Frame31
    Set mp = f.Controls("mpCogMental")
    Set pgCog = mp.Pages("pgCognition")
    Set pgMental = mp.Pages("pgMental")

    '=== 認知側 combobox 群 ===
    v = ws.Cells(r, COL_COG_MEMORY).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogMemory").value = v

    v = ws.Cells(r, COL_COG_ATTENTION).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogAttention").value = v

    v = ws.Cells(r, COL_COG_ORIENTATION).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogOrientation").value = v

    v = ws.Cells(r, COL_COG_JUDGEMENT).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogJudgement").value = v

    v = ws.Cells(r, COL_COG_EXECUTIVE).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogExecutive").value = v

    v = ws.Cells(r, COL_COG_LANGUAGE).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbCogLanguage").value = v

    v = ws.Cells(r, COL_COG_DEMENTIA).value
    If IsNull(v) Then v = ""
    pgCog.Controls("cmbDementiaType").value = v

    v = ws.Cells(r, COL_COG_DEM_NOTE).value
    If IsNull(v) Then v = ""
    pgCog.Controls("txtDementiaNote").Text = v

    '=== BPSD（chkBPSD0?10）===
    ' 1) 全部一度クリア
    For i = 0 To 10
        Set chk = pgCog.Controls("chkBPSD" & CStr(i))
        chk.value = False
    Next i

    ' 2) セル文字列を | で分解し、Caption と一致するチェックボックスをON
    s = ws.Cells(r, COL_COG_BPSD).value & ""
    If Len(s) > 0 Then
        arr = Split(s, "|")
        For i = LBound(arr) To UBound(arr)
            For j = 0 To 10
                Set chk = pgCog.Controls("chkBPSD" & CStr(j))
                If chk.caption = arr(i) Then
                    chk.value = True
                    Exit For
                End If
            Next j
        Next i
    End If

    '=== 精神面 combobox / note ===
    v = ws.Cells(r, COL_MENTAL_MOOD).value
    If IsNull(v) Then v = ""
    pgMental.Controls("cmbMood").value = v

    v = ws.Cells(r, COL_MENTAL_MOTIV).value
    If IsNull(v) Then v = ""
    pgMental.Controls("cmbMotivation").value = v

    v = ws.Cells(r, COL_MENTAL_ANXIETY).value
    If IsNull(v) Then v = ""
    pgMental.Controls("cmbAnxiety").value = v

    v = ws.Cells(r, COL_MENTAL_RELATION).value
    If IsNull(v) Then v = ""
    pgMental.Controls("cmbRelation").value = v

    v = ws.Cells(r, COL_MENTAL_SLEEP).value
    If IsNull(v) Then v = ""
    pgMental.Controls("cmbSleep").value = v

    v = ws.Cells(r, COL_MENTAL_NOTE).value
    If IsNull(v) Then v = ""
    pgMental.Controls("txtMentalNote").Text = v
End Sub


Public Sub Save_DailyLog_FromForm(owner As Object)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim f As Object
    Dim txtName As Object
    Dim txtDate As Object
    Dim txtStaff As Object
    Dim txtNote As Object
    Dim lastRow As Long
    Dim r As Long

    Set wb = ThisWorkbook

    '--- DailyLog シート取得 or 作成 ---
    For Each sh In wb.Worksheets
        If sh.name = "DailyLog" Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = "DailyLog"
        ws.Range("A1").value = "記録日"
        ws.Range("B1").value = "利用者名"
        ws.Range("C1").value = "記録者"
        ws.Range("D1").value = "記録内容"
    End If

    ' 既存シートでもヘッダが空ならヘッダを補正
    If ws.Cells(1, 1).value = "" And _
       ws.Cells(1, 2).value = "" And _
       ws.Cells(1, 3).value = "" And _
       ws.Cells(1, 4).value = "" Then

        ws.Range("A1").value = "記録日"
        ws.Range("B1").value = "利用者名"
        ws.Range("C1").value = "記録者"
        ws.Range("D1").value = "記録内容"
    End If

    '--- 書き込み行を決定（最終行の次） ---
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If lastRow < 1 Then lastRow = 1
    r = lastRow + 1

    '--- フォーム上のコントロール取得 ---
    Set txtName = owner.Controls("txtName")          ' 利用者名（frmEval 共通）
    Set f = owner.Controls("fraDailyLog")            ' モニタリング用フレーム

    Set txtDate = f.Controls("txtDailyDate")         ' 記録日
    Set txtStaff = f.Controls("txtDailyStaff")       ' 記録者
    Set txtNote = f.Controls("txtDailyNote")         ' 記録内容

    '--- DailyLog シートへ保存 ---
    ws.Cells(r, 1).value = CStr(txtDate.value)
    ws.Cells(r, 2).value = CStr(txtName.value)
    ws.Cells(r, 3).value = CStr(txtStaff.value)
    ws.Cells(r, 4).value = CStr(txtNote.value)
    ws.Cells(r, 1).NumberFormatLocal = "yyyy/mm/dd"   ' ←これを追加
    


End Sub




Public Sub Load_DailyLog_Latest_FromForm(owner As Object)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim f As Object
    Dim txtName As Object
    Dim txtDate As Object
    Dim txtStaff As Object
    Dim txtNote As Object
    Dim lastRow As Long
    Dim r As Long
    Dim targetName As String
    Dim hit As Boolean

    Set wb = ThisWorkbook

    '--- DailyLog シート取得（無ければ何もしない） ---
    For Each sh In wb.Worksheets
        If sh.name = "DailyLog" Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then

        Exit Sub
    End If

    '--- フォーム上のコントロール取得 ---
    Set txtName = owner.Controls("txtName")
    Set f = owner.Controls("fraDailyLog")
    Set txtDate = f.Controls("txtDailyDate")
    Set txtStaff = f.Controls("txtDailyStaff")
    Set txtNote = f.Controls("txtDailyNote")

    targetName = Trim$(CStr(txtName.value))
    If targetName = "" Then

        Exit Sub
    End If

    '--- 該当利用者の「最新（いちばん下）」の行を探す ---
    lastRow = ws.Cells(ws.rows.Count, 2).End(xlUp).row   ' B列＝利用者名
    If lastRow < 2 Then

        Exit Sub
    End If

    hit = False
    For r = lastRow To 2 Step -1
        If Trim$(CStr(ws.Cells(r, 2).value)) = targetName Then
            hit = True
            Exit For
        End If
    Next r

    If Not hit Then

        Exit Sub
    End If

    '--- 見つかった行をフォームへ反映 ---
    txtDate.value = ws.Cells(r, 1).value     ' 記録日
    txtStaff.value = ws.Cells(r, 3).value    ' 記録者
    txtNote.value = ws.Cells(r, 4).value     ' 記録内容


End Sub









Public Sub SaveDailyLog_Append(owner As Object)




    
    ' 専用ボタンからの呼び出し以外では何もしない
    If Not mDailyLogManual Then Exit Sub

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim f As Object
    Dim dt As Variant
    Dim nm As String
    Dim staff As String
    Dim note As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("DailyLog")  ' ★ 日々の記録シート名（変えるならここ）
    Set f = owner.Controls("fraDailyLog")

    '--- 入力値取得 ---
    dt = f.Controls("txtDailyDate").value
    nm = Trim$(owner.Controls("frHeader").Controls("txtHdrName").value)
    staff = Trim$(f.Controls("txtDailyStaff").value)
    note = f.Controls("txtDailyNote").value

    '--- 入力チェック ---
    If nm = "" Then
        MsgBox "氏名を入力してください。", vbExclamation
        Exit Sub
    End If

    If Not IsDate(dt) Then
        MsgBox "記録日の欄に正しい日付を入力してください。", vbExclamation
        Exit Sub
    End If

    If note = "" Then
        If MsgBox("記録内容が空ですが保存しますか？", vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
    End If

    '--- 追記行を決める（1行目に見出しがある前提）---
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).row + 1

    '--- 書き込み ---
    ws.Cells(r, 1).value = CDate(dt)   ' 記録日
    ws.Cells(r, 2).value = nm          ' 利用者名
    ws.Cells(r, 3).value = Trim$(owner.Controls("frHeader").Controls("txtHdrPID").value) ' ★ID
    ws.Cells(r, 4).value = staff       ' 記録者
    ws.Cells(r, 5).value = note        ' 記録内容

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
    For Each c In parent.Controls
        DumpIfIDLike c
        If HasControls(c) Then Walk c
    Next
End Sub

Private Function HasControls(ByVal o As Object) As Boolean
    On Error Resume Next
    Dim n As Long
    n = o.Controls.Count
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

