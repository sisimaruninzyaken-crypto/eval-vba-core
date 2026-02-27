VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEval 
   Caption         =   "評価フォーム"
   ClientHeight    =   8580.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17124
   OleObjectBlob   =   "frmEval.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

'=== frmEval ヘッダ：共通で使う変数 ===
Private mp As MSForms.MultiPage            ' ルート MultiPage
Private mpWalk As MSForms.MultiPage        ' 歩行評価タブ内のサブ MultiPage  ← これが今回の追加
Private hostBasic As MSForms.Frame, hostPost As MSForms.Frame, hostMove As MSForms.Frame
Private hostTests As MSForms.Frame, hostWalk As MSForms.Frame, hostCog As MSForms.Frame
Private WithEvents btnLoadPrevCtl As MSForms.CommandButton
Attribute btnLoadPrevCtl.VB_VarHelpID = -1
Private WithEvents btnSaveCtl     As MSForms.CommandButton
Attribute btnSaveCtl.VB_VarHelpID = -1
Private WithEvents btnCloseCtl    As MSForms.CommandButton
Attribute btnCloseCtl.VB_VarHelpID = -1
Private Const ADL_TOP_OFFSET As Long = -30
Private ImeHooks As Collection
Private MPHs As Collection
Private builtPosture As Boolean
Private Const CAP_POSTURE_PAGE As String = "身体機能＋起居動作"
Private Const POSTURE_TAG_PREFIX As String = "POSTURE|"
Private Const POSTURE_COLS As Long = 2
Private mStyleDone As Boolean
Private mMPHooks  As Collection
Private mTxtHooks As Collection
Private mHooked As Boolean
Private mLayoutDone As Boolean
Private mPainBuilt As Boolean
Private mPainLayoutDone As Boolean
' 既存の「シートへ保存」「前回の値を読み込む」にフックする
Private WithEvents btnHdrSave     As MSForms.CommandButton
Attribute btnHdrSave.VB_VarHelpID = -1
Private WithEvents btnHdrLoadPrev As MSForms.CommandButton
Attribute btnHdrLoadPrev.VB_VarHelpID = -1
Private fBasicRef As MSForms.Frame
Private nextTop As Single
' BIコンボのイベントフックを保持
Private BIHooks As Collection
' レイアウト用
Private FWIDTH As Single, COL_LX As Single, COL_RX As Single, lblW As Single, pad As Single, rowH As Single
' === ウィンドウ制約（スクロール無し用の最小サイズ）===
Private Const MIN_W As Long = 720
Private Const MIN_H As Long = 520
Private WithEvents mVAS As MSForms.ScrollBar
Attribute mVAS.VB_VarHelpID = -1
Private WithEvents mBtnPainSum As MSForms.CommandButton
Attribute mBtnPainSum.VB_VarHelpID = -1
Private mPainTidyBusy As Boolean
Private WithEvents mDailyExtract As MSForms.CommandButton
Attribute mDailyExtract.VB_VarHelpID = -1
Private WithEvents mDailySave As MSForms.CommandButton
Attribute mDailySave.VB_VarHelpID = -1
Private mPlacedGlobalSave As Boolean
Private WithEvents mGlobalSave As clsGlobalSaveButton
Attribute mGlobalSave.VB_VarHelpID = -1
Private WithEvents mGlobalClear As clsGlobalSaveButton
Attribute mGlobalClear.VB_VarHelpID = -1
Private WithEvents mDailyList As clsDailyLogList
Attribute mDailyList.VB_VarHelpID = -1
Private mBaseLayoutDone As Boolean
Private mLayoutBuilt As Boolean
Private Const PAD_SIDE As Single = 8
Private Const GAP_V As Single = 8
Private Const HEADER_H As Single = 62
Private mHdr1 As clsHeaderBtnEvents
Private mHdr2 As clsHeaderBtnEvents
Private mHdr3 As clsHeaderBtnEvents
Private mScrollOnce_347 As Boolean
Private mQuitExcelRequested As Boolean
Private Enum QuitMode
    qmNone = 0
    qmAsk   '閉じるボタン：保存確認してExcel終了
End Enum

Private mQuitMode As QuitMode
Private mHdrArchiveHook As clsHdrBtnHook
Private mHdrLoadPrevHook As clsHdrBtnHook
Private mPrintBtnHook As clsPrintBtnHook
Private mHdrNameSink As cHdrNameSink
Private mNameSuggestSink As cNameSuggestSink
Private mDupNameWarned As Boolean



Public Sub SyncAgeFromBirth()
    On Error GoTo EH

    Dim s As String
    s = Trim$(Me.Controls("txtBirth").Text)
    If Len(s) = 0 Then Exit Sub

    ' yyyy/mm/dd, yyyy-mm-dd を想定（DateValueに寄せる）
   s = Replace$(s, "-", "/")

' ★ここに入れる
Dim d As String
d = Replace$(Replace$(s, "/", ""), ".", "")
If Len(d) = 8 And IsNumeric(d) Then
    s = Left$(d, 4) & "/" & Mid$(d, 5, 2) & "/" & Right$(d, 2)
    Me.Controls("txtBirth").Text = s
End If

Dim dt As Date
dt = DateValue(s)


    Dim today As Date
    today = Date

    Dim age As Long
    age = DateDiff("yyyy", dt, today)
    If DateSerial(Year(today), Month(dt), day(dt)) > today Then
        age = age - 1
    End If

    Me.Controls("txtAge").Text = CStr(age)

#If APP_DEBUG Then
    Debug.Print "[SyncAgeFromBirth] birth=", Format$(dt, "yyyy/mm/dd"), " age=", age
#End If

    Exit Sub
EH:
#If APP_DEBUG Then
    Debug.Print "[SyncAgeFromBirth][ERR]", Err.Number, Err.Description
#End If
End Sub




'=== 共通ヘルパー：フレームの高さを子コントロールの一番下＋余白まで伸ばす ===
Private Sub FitFrameHeightToChildren(f As MSForms.Frame, Optional margin As Single = 6)
    
    'FitFrameHeightToChildren Me.Controls("Frame7")
    
    Dim c As Control
    Dim maxBottom As Single

    If f Is Nothing Then Exit Sub

    For Each c In f.Controls
        If c.Top + c.Height > maxBottom Then
            maxBottom = c.Top + c.Height
        End If
    Next c

    If maxBottom + margin > f.Height Then
        f.Height = maxBottom + margin
    End If
End Sub







'--- 列解決の安全ラッパ（あれば ResolveColumnLocal、なければ見出し名で探す）---
Private Function RCol(ws As Worksheet, look As Object, ParamArray headers()) As Long
    Dim c As Long, i As Long
    On Error Resume Next
    c = ResolveColumnLocal(look, CStr(headers(LBound(headers))))  ' 存在しない環境でもOK
    On Error GoTo 0
    If c > 0 Then
        RCol = c
        Exit Function
    End If
    For i = LBound(headers) To UBound(headers)
        c = modEvalIOEntry.FindColByHeaderExact(ws, CStr(headers(i)))

        If c > 0 Then RCol = c: Exit Function
    Next i
End Function



' 既存コードで参照されるが未定義だったものを最小限だけ用意
Private Sub SetupLayout()
    ' ※数値は一般的な既定サイズ。UI見た目は変えない範囲で安全な値。
    FWIDTH = 800
    lblW = 70
    COL_LX = 12
    COL_RX = 380
    pad = 6
    rowH = 24
End Sub

' 既存コードで Call されるが中身が不要な互換ダミー（No-op）
Private Sub BuildBliadlControls(ByVal mp As MSForms.MultiPage)
    ' no-op（互換用）
End Sub
Private Sub BuildBIPage(ByVal mpADL As MSForms.MultiPage)
    ' no-op（互換用）
End Sub
Private Sub BuildIADLPage(ByVal mpADL As MSForms.MultiPage)
    ' no-op（互換用）
End Sub

' 既存の保存処理で参照されるラッパー（見た目と挙動は変えない）
Private Function CtrlText(ByVal ctrlName As String) As String
    On Error Resume Next
    CtrlText = Trim$(Me.Controls(ctrlName).Text & "")
End Function
Private Sub SetCtrlText(ByVal ctrlName As String, ByVal v As String)
    On Error Resume Next
    Me.Controls(ctrlName).Text = v
End Sub

'=== ここから 画面作成ヘルパーの最小実装 =========================
Private Function CreateFrameP(parent As MSForms.Frame, title As String, _
                              Optional minHeight As Single = 120) As MSForms.Frame
    Dim f As MSForms.Frame
    Set f = parent.Controls.Add("Forms.Frame.1")
    With f
        .caption = title
        .Left = 6
        .Top = nextTop
        .Width = parent.InsideWidth - 12
        .Height = minHeight
        .ScrollBars = fmScrollBarsNone
    End With
    nextTop = f.Top + f.Height + 6
    Set CreateFrameP = f
End Function

Private Sub ResizeFrameToContent(f As MSForms.Frame, contentBottom As Single)
    If contentBottom + 22 > f.Height Then f.Height = contentBottom + 22
    nextTop = f.Top + f.Height + 6
End Sub

Private Sub nL(ByRef Y As Single, Optional ByVal rows As Long = 1)
    Y = Y + rows * 26
End Sub



' 既存呼び出し互換：caption → x → y → [w] → [name]
Function CreateLabel( _
    parent As MSForms.Frame, _
    ByVal caption As String, _
    ByVal x As Single, _
    ByVal Y As Single, _
    Optional ByVal w As Single = 160, _
    Optional ByVal nm As String = "" _
) As MSForms.label

    Dim lb As MSForms.label
    If nm <> "" Then
        Set lb = parent.Controls.Add("Forms.Label.1", nm)
    Else
        Set lb = parent.Controls.Add("Forms.Label.1")
    End If

    With lb
        .caption = caption
        .Left = x
        .Top = Y
        .AutoSize = False
        .Width = w
    End With

    Set CreateLabel = lb
End Function



Function CreateLabelXY( _
    parent As MSForms.Frame, _
    ByVal x As Single, _
    ByVal Y As Single, _
    Optional ByVal caption As String = "", _
    Optional ByVal nm As String = "", _
    Optional ByVal w As Single = 160 _
) As MSForms.label
    Set CreateLabelXY = CreateLabel(parent, caption, x, Y, w, nm)
End Function




Private Function CreateTextBox(parent As MSForms.Frame, x As Single, Y As Single, _
                               w As Single, h As Single, multiline As Boolean, _
                               Optional name As String = "", Optional tag As String = "") As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = parent.Controls.Add("Forms.TextBox.1", IIf(name = "", vbNullString, name))
    With tb
        .Left = x
        .Top = Y
        .Width = w
        .Height = IIf(h > 0, h, 20)
        .multiline = multiline
        .EnterKeyBehavior = multiline
        .tag = tag
    End With
    Set CreateTextBox = tb
End Function

Private Function CreateCombo(parent As MSForms.Frame, x As Single, Y As Single, _
                             w As Single, Optional name As String = "", Optional tag As String = "") As MSForms.ComboBox
    Dim cB As MSForms.ComboBox
    Set cB = parent.Controls.Add("Forms.ComboBox.1", IIf(name = "", vbNullString, name))
    With cB
        .Left = x
        .Top = Y
        .Width = w
        .Style = fmStyleDropDownList
        .tag = tag
    End With
    Set CreateCombo = cB
End Function

Private Function CreateCheck(parent As MSForms.Frame, caption As String, _
                             x As Single, Y As Single, Optional name As String = "", _
                             Optional tag As String = "") As MSForms.CheckBox
    Dim ck As MSForms.CheckBox
    Set ck = parent.Controls.Add("Forms.CheckBox.1", IIf(name = "", vbNullString, name))
    With ck
        .caption = caption
        .Left = x
        .Top = Y
        .tag = tag
    End With
    Set CreateCheck = ck
End Function

Private Function MakeList(csv As String) As Variant
    MakeList = Split(csv, ",")
End Function

Private Sub PositionTopRightButtons(f As MSForms.Frame)
    On Error Resume Next
    If Not btnSaveCtl Is Nothing Then
        btnSaveCtl.Top = 8
        btnSaveCtl.Left = f.Width - btnSaveCtl.Width - 12
    End If
End Sub

'=== チェックボックスを並べる汎用フレーム ===
Private Function BuildCheckFrame(parent As MSForms.Frame, _
    title As String, x As Single, Y As Single, w As Single, _
    items As Variant, Optional groupTag As String = "") As MSForms.Frame

    Dim f As MSForms.Frame
    Set f = parent.Controls.Add("Forms.Frame.1")
    With f
        .caption = title
        .Left = x
        .Top = Y
        .Width = w
        .Height = 60                ' 仮高さ。下で中身に合わせて伸ばす
        .ScrollBars = fmScrollBarsNone
    End With

    ' チェックを2列で配置（長くなりすぎない程度）
    Dim i As Long, col As Long, row As Long
    Dim colW As Single: colW = (w - 24) / 2
    Dim rowH As Single: rowH = 20
    Dim maxRow As Long: maxRow = 0

    For i = LBound(items) To UBound(items)
        col = (i - LBound(items)) Mod 2
        row = (i - LBound(items)) \ 2
        Dim ck As MSForms.CheckBox
        Set ck = f.Controls.Add("Forms.CheckBox.1", "ck_" & CStr(i))
        With ck
            .caption = CStr(items(i))
            .Left = 12 + col * colW
            .Top = 18 + row * rowH
            .tag = IIf(Len(groupTag) = 0, "", groupTag)
        End With
        If row > maxRow Then maxRow = row
    Next i

    ' 中身の高さに合わせてフレームを伸ばす
    f.Height = 18 + (maxRow + 1) * rowH + 12

    Set BuildCheckFrame = f
End Function

Private Sub BuildAssistiveChecksInWalkEval(ByVal assistiveCsv As String)
    Dim frTarget As MSForms.Frame
    Set frTarget = GetWalkEvalAssistiveTargetFrame()
    If frTarget Is Nothing Then Exit Sub

    Dim i As Long
    For i = frTarget.Controls.Count - 1 To 0 Step -1
        If TypeName(frTarget.Controls(i)) = "CheckBox" Then
            If frTarget.Controls(i).tag = "AssistiveGroup" Then
                frTarget.Controls.Remove frTarget.Controls(i).name
            End If
        End If
    Next

    Dim maxBottom As Single
    maxBottom = 0
    For i = 0 To frTarget.Controls.Count - 1
        With frTarget.Controls(i)
            If .Top + .Height > maxBottom Then maxBottom = .Top + .Height
        End With
    Next

    Dim addTop As Single
    addTop = IIf(maxBottom <= 0, 120, maxBottom + 120)

    Dim frAssist As MSForms.Frame
    Set frAssist = BuildCheckFrame(frTarget, "補助具", 8, addTop, frTarget.InsideWidth - 16, MakeList(assistiveCsv), "AssistiveGroup")

    If frAssist.Top + frAssist.Height + 8 > frTarget.Height Then
        frTarget.Height = frAssist.Top + frAssist.Height + 8
    End If
End Sub

Private Function GetWalkEvalAssistiveTargetFrame() As MSForms.Frame
    On Error GoTo EH
    Set GetWalkEvalAssistiveTargetFrame = Me.Controls("MultiPage1") _
        .Pages("Page6").Controls("Frame6") _
        .Controls("MultiPage2").Pages("Page8") _
        .Controls("Frame24").Controls("Frame25")
    Exit Function
EH:
    Set GetWalkEvalAssistiveTargetFrame = Nothing
End Function

'=== 関節拘縮：左右チェックを1行生成するヘルパー ======================
Private Sub CreateContractureRLRow(parent As MSForms.Frame, _
                                   ByRef Y As Single, _
                                   ByVal partCaption As String, _
                                   ByVal baseTag As String)
    ' ガイド：先頭で見出しを作ったレイアウトに合わせる
    '   部位: COL_LX
    '   右  : COL_LX + 90 + 20
    '   左  : （右の列）+ 60

    ' 見出し（部位名）
    Call CreateLabel(parent, partCaption, COL_LX, Y)

    ' 右チェック
    Call CreateCheck(parent, "右", COL_LX + 90 + 20, Y, , baseTag & ".右")

    ' 左チェック
    Call CreateCheck(parent, "左", COL_LX + 90 + 20 + 60, Y, , baseTag & ".左")

    ' 次の行へ
    nL Y
End Sub




'=== RLAチェック群を作る最小実装 ======================================
Private Sub Build_RLA_ChecksPart(f As MSForms.Frame, ByVal kind As String)
    Dim phases As Variant
    If LCase$(kind) = "stance" Then
        phases = Array("IC", "LR", "MSt", "TSt")        ' 立脚期
    Else
        phases = Array("PSw", "ISw", "MSw", "TSw")      ' 遊脚期
    End If

    Dim i As Long, Y As Single
    Y = 22

    For i = LBound(phases) To UBound(phases)
        Dim key As String: key = CStr(phases(i))

        ' 左端：フェーズ名ラベル
        Call CreateLabel(f, RLAPhaseCaption(key), 12, Y, 90)

        ' 簡易チェック（4つ）※名前は "RLA_<key>_<番号>" とする
        Call AddRLAChk(f, key, "可動域不足", 120, Y)
        Call AddRLAChk(f, key, "筋力低下", 300, Y)
        Y = Y + 22
        Call AddRLAChk(f, key, "疼痛/不安定", 120, Y)
        Call AddRLAChk(f, key, "協調不良", 300, Y)

        ' 右端：レベル選択（OptionButton, GroupName=key）
        Call AddRLAOpt(f, key, "軽度", f.Width - 180, Y - 22)
        Call AddRLAOpt(f, key, "中等度", f.Width - 120, Y - 22)
        Call AddRLAOpt(f, key, "高度", f.Width - 60, Y - 22)

        Y = Y + 30
    Next

    ' フレームの高さは呼び出し元で ResizeFrameToContent しているためここでは触らない
End Sub

' フェーズ名（見出し用）
Private Function RLAPhaseCaption(ByVal key As String) As String
    Select Case key
        Case "IC":  RLAPhaseCaption = "IC"
        Case "LR":  RLAPhaseCaption = "LR"
        Case "MSt": RLAPhaseCaption = "MSt"
        Case "TSt": RLAPhaseCaption = "TSt"
        Case "PSw": RLAPhaseCaption = "PSw"
        Case "ISw": RLAPhaseCaption = "ISw"
        Case "MSw": RLAPhaseCaption = "MSw"
        Case "TSw": RLAPhaseCaption = "TSw"
        Case Else:  RLAPhaseCaption = key
    End Select
End Function

' チェックボックス追加（名前は RLA_<key>_n）
Private Sub AddRLAChk(f As MSForms.Frame, ByVal key As String, ByVal caption As String, _
                      ByVal x As Single, ByVal Y As Single)
    Dim ck As MSForms.CheckBox
    Set ck = f.Controls.Add("Forms.CheckBox.1", "RLA_" & key & "_" & Replace(caption, "/", "_"))
    With ck
        .caption = caption
        .Left = x
        .Top = Y
    End With
End Sub

' レベル用オプションボタン（GroupName=key）
Private Sub AddRLAOpt(f As MSForms.Frame, ByVal key As String, ByVal caption As String, _
                      ByVal x As Single, ByVal Y As Single)
    Dim ob As MSForms.OptionButton
    Set ob = f.Controls.Add("Forms.OptionButton.1")
    With ob
        .caption = caption
        .groupName = key
        .Left = x
        .Top = Y
    End With
End Sub

' 呼び出し元で参照しているための互換ダミー（既定選択は特に設定しない）
Private Sub InitRLAdefaults()
    ' no-op
End Sub
'======================================================================

'=== ここから：ヘッダ検索系の互換ラッパー =========================
' Local実装を既に入れている前提（BuildHeaderLookupLocal / ResolveColumnLocal / EnsureHeaderColumnLocal）
' 呼び出し元とシグネチャを合わせるための薄いラッパーだけ用意

' 1) BuildHeaderLookup の別名ラッパー
Private Function BuildHeaderLookup(ByVal ws As Worksheet) As Object
    Set BuildHeaderLookup = BuildHeaderLookupLocal(ws)
End Function

' 2) ResolveColumn の別名ラッパー
Private Function ResolveColumn(ByVal look As Object, ByVal key As String) As Long
    ResolveColumn = ResolveColumnLocal(look, key)
End Function

' 3) ResolveColOrCreate
'    第1優先キーが無ければエイリアスを順に探し、それでも無ければ第1優先キーで列を新規作成
Private Function ResolveColOrCreate(ByVal ws As Worksheet, ByVal look As Object, _
                                    ByVal primaryKey As String, ParamArray aliases() As Variant) As Long
    Dim col As Long, i As Long
    col = ResolveColumnLocal(look, primaryKey)
    If col = 0 Then
        For i = LBound(aliases) To UBound(aliases)
            col = ResolveColumnLocal(look, CStr(aliases(i)))
            If col <> 0 Then Exit For
        Next
    End If
    If col = 0 Then
        col = EnsureHeaderColumnLocal(ws, look, primaryKey)
    End If
    ResolveColOrCreate = col
End Function

'=== ComboBox に配列を確実に流し込む最小ヘルパー ===
Private Sub SetComboItems(ByRef cbo As MSForms.ComboBox, ByVal items As Variant)
    Dim k As Long
    cbo.Clear
    If IsArray(items) Then
        For k = LBound(items) To UBound(items)
            cbo.AddItem CStr(items(k))
        Next
    End If
    On Error Resume Next
    cbo.ListIndex = -1
End Sub



'=== レイアウト自動フィット（安全ガード付き） ===
Private Sub FitLayout()
    On Error Resume Next

    ' ルートの有効寸法（下限でクランプ）
    Dim iw As Single, iH As Single
    iw = Me.InsideWidth - 12
    iH = Me.InsideHeight - 60
    If iw < 240 Then iw = 240          ' 幅の下限
    If iH < 180 Then iH = 180          ' 高さの下限（←ココが 0 以下になると 380）

    ' ルート MultiPage
    If Not mp Is Nothing Then
        mp.Left = 6: mp.Top = 6
        mp.Width = iw
        mp.Height = iH
    End If

    ' 各ホストフレームも同様にクランプ
    Dim hosts As Variant, h As MSForms.Frame, i As Long
    hosts = Array(hostBasic, hostPost, hostMove, hostTests, hostWalk, hostCog)
    For i = LBound(hosts) To UBound(hosts)
        Set h = hosts(i)
        If Not h Is Nothing Then
            h.Left = 0: h.Top = 0
            h.Width = iw - 12: If h.Width < 120 Then h.Width = 120
            h.Height = iH - 12: If h.Height < 100 Then h.Height = 100
            h.ScrollBars = fmScrollBarsNone
        End If
    Next i

    ' 「閉じる」ボタン
    If Not btnCloseCtl Is Nothing Then
        btnCloseCtl.Left = Me.InsideWidth - btnCloseCtl.Width - 10
        btnCloseCtl.Top = 6 + iH + 8
    End If

    Me.ScrollBars = fmScrollBarsNone
    Me.ScrollHeight = Me.InsideHeight
End Sub


'=== 互換ダミー：タブ順リセット（何もしない） =================
Private Sub ResetTabOrder()
End Sub
'============================================================

'=== 氏名で候補行を集めて Variant 配列にして返す =====================
Private Function CollectCandidatesByNameLocal(ByVal ws As Worksheet, _
                                              ByVal look As Object, _
                                              ByVal pname As String) As Variant
    Dim nameCol As Long
    nameCol = ResolveColumnLocal(look, "Basic.Name")
    If nameCol = 0 Then nameCol = ResolveColumnLocal(look, "氏名")
    If nameCol = 0 Then nameCol = ResolveColumnLocal(look, "Name")
    If nameCol = 0 Then Exit Function

    Dim key As String: key = NormName(pname)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, nameCol).End(xlUp).row
    Dim tmp As New Collection
    Dim r As Long, nm As String
    For r = 2 To lastRow
        nm = CStr(ws.Cells(r, nameCol).value)
        If NormName(nm) = key Then tmp.Add r
    Next

    If tmp.Count = 0 Then Exit Function

    Dim a() As Long: ReDim a(1 To tmp.Count)
    For r = 1 To tmp.Count
        a(r) = CLng(tmp(r))
    Next
    CollectCandidatesByNameLocal = a
End Function





' ※既に同名の関数があればそちらを使ってください
Private Function NormName(ByVal s As String) As String
    s = Replace(s, vbCrLf, "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    On Error Resume Next
    s = StrConv(s, vbNarrow)    ' 全角→半角（環境により失敗してもOK）
    On Error GoTo 0
    NormName = LCase$(s)
End Function





'=== 文字列の正規化：半角/全角・余分スペース・大小を吸収して照合用キーにする ===
Private Function KeyNormalize(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)

    ' 改行や全角スペースを通常スペースに
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(&H3000), " ") ' 全角スペース→半角

    ' 連続スペース圧縮＆前後トリム
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)

    ' 全角→半角（ASCII/数字/カタカナなど対象）
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    ' よくあるハイフン類の統一（??－‐→-）
    s = Replace(s, "－", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "‐", "-")

    ' 大文字化（英字ゆらぎ対策）
    s = UCase$(s)

    ' 照合時はスペース無視
    s = Replace(s, " ", "")

    KeyNormalize = s
End Function

'=== 入力モード：日本語/半角の自動切替 ===============================

' 画面全体のコントロールに IME モードを適用（再帰）
Private Sub SetupInputModesJP()
    ApplyInputModeJP Me

    If FnHasControl("txtAge") Then
    Debug.Print "[IME] txtAge =", Me.Controls("txtAge").IMEMode
Else
    Debug.Print "[IME] txtAge = (not found)"
End If

If FnHasControl("txtEDate") Then
    Debug.Print "[IME] txtEDate =", Me.Controls("txtEDate").IMEMode
Else
    Debug.Print "[IME] txtEDate = (not found)"
End If

If FnHasControl("txtPost_Note") Then
    Debug.Print "[IME] txtPost_Note =", Me.Controls("txtPost_Note").IMEMode
Else
    Debug.Print "[IME] txtPost_Note = (not found)"
End If

End Sub

'=== ヘルパー：コントロール存在チェック（衝突回避の別名） ===
Private Function FnHasControl(ByVal nm As String) As Boolean
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If StrComp(c.name, nm, vbTextCompare) = 0 Then
            FnHasControl = True
            Exit Function
        End If
    Next
    FnHasControl = False
End Function








'=== IME切替：コンテナを再帰的に処理（MultiPage対応版） ==============
Private Sub ApplyInputModeJP(container As Object)
    Dim typ As String
    On Error Resume Next
    typ = TypeName(container)
    On Error GoTo 0

    If typ = "MultiPage" Then
    Dim pg As MSForms.Page
    For Each pg In container.Pages
        ApplyInputModeJP pg
    Next
    Exit Sub
      End If


    ' Controls を持たないものは終了（Frame/Page 以外の安全対策）
    If Not HasControls(container) Then Exit Sub

    Dim c As MSForms.Control
    For Each c In container.Controls
        Select Case TypeName(c)
            Case "TextBox", "ComboBox"
                On Error Resume Next
                If ShouldBeNumericField(c) Then
                    c.IMEMode = fmIMEModeDisable     ' 半角英数
                Else
                    c.IMEMode = fmIMEModeHiragana     ' ひらがな
                End If
                On Error GoTo 0

            Case "Frame", "Page"
                ApplyInputModeJP c                   ' 子へ再帰

            Case "MultiPage"
                ' 子に MultiPage がぶら下がっている場合は Pages を回す
                Dim p As MSForms.Page
                For Each p In c.Pages
                    ApplyInputModeJP p
                Next
        End Select
    Next c
End Sub


' このフォームで「数字欄」と見なす判定
Private Function ShouldBeNumericField(c As MSForms.Control) As Boolean
    Dim nm As String: nm = LCase$(c.name & "")
    Dim tg As String: tg = LCase$(c.tag & "")

    ' 名前で判定（必要ならここに追記）
    If nm = "txtage" Or nm = "txtedate" Or nm = "txtonset" _
       Or nm = "txttenmwalk" Or nm = "txttug" Or nm = "txtfivests" _
       Or nm = "txtsemi" Or nm = "txtgripr" Or nm = "txtgripl" _
       Or nm = "txtbi" Or nm = "txtpid" Then
        ShouldBeNumericField = True: Exit Function
    End If

    ' Tagで判定（既存のTagを活用）
    If InStr(tg, "evaldate") > 0 Or InStr(tg, "onsetdate") > 0 Then ShouldBeNumericField = True: Exit Function
    If Left$(tg, 5) = "test." Or Left$(tg, 5) = "grip." Then ShouldBeNumericField = True: Exit Function
    If Right$(tg, 4) = ".age" Or tg = "bi.total" Then ShouldBeNumericField = True: Exit Function
End Function

' （任意）保存前に数字欄を半角に統一しておく
Private Sub NormalizeNumericInputsToHalfwidth()
    Dim c As MSForms.Control
    For Each c In Me.Controls
        Call NormalizeNumericInContainer(c)
    Next
End Sub

'=== 数値欄の半角統一：MultiPage対応版 ================================
Private Sub NormalizeNumericInContainer(container As Object)
    Dim typ As String
    On Error Resume Next
    typ = TypeName(container)
    On Error GoTo 0

    If typ = "MultiPage" Then
        Dim pg As MSForms.Page
        For Each pg In container.Pages
            NormalizeNumericInContainer pg
        Next
        Exit Sub
    End If

    If Not HasControls(container) Then Exit Sub

    Dim c As MSForms.Control
    For Each c In container.Controls
        Select Case TypeName(c)
            Case "TextBox", "ComboBox"
                If ShouldBeNumericField(c) Then
                    On Error Resume Next
                    c.Text = StrConv(c.Text, vbNarrow) ' 全角→半角
                    On Error GoTo 0
                End If

            Case "Frame", "Page"
                NormalizeNumericInContainer c

            Case "MultiPage"
                Dim p As MSForms.Page
                For Each p In c.Pages
                    NormalizeNumericInContainer p
                Next
        End Select
    Next c
End Sub


Private Function HasControls(o As Object) As Boolean
    On Error Resume Next
    Dim t As Long: t = o.Controls.Count
    HasControls = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
'====================================================================

'=== ComboBox に配列を確実に流し込む最小ヘルパー ===
Private Sub FillComboItems(ByRef cbo As MSForms.ComboBox, ByVal items As Variant)
    Dim k As Long
    cbo.Clear
    If IsArray(items) Then
        For k = LBound(items) To UBound(items)
            cbo.AddItem CStr(items(k))
        Next
    End If
    ' 既定選択なし
    On Error Resume Next
    cbo.ListIndex = -1
End Sub



'=== BI/IADL を強制再構築（必ず中身を表示） ===============================
Private Function EnsureBI_IADL() As MSForms.MultiPage
    On Error Resume Next

    Trace "EnsureBI_IADL start", "BI/IADL"   ' ←①ここ




    
    Dim mpADL As MSForms.MultiPage
    Dim nextTop As Single
    

 

 ' 1) ルート MultiPage と「日常生活動作」ページを特定
    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Function

    Dim pgMove As MSForms.Page: Set pgMove = FindPageByCaption(mp, "日常生活動作")
    If pgMove Is Nothing Then
        If mp.Pages.Count >= 3 Then
            Set pgMove = mp.Pages(2) ' フォールバック
        Else
            Exit Function
        End If
    End If

    ' 2) ページ内のホスト Frame を取得（無ければ作成）
    Dim host As MSForms.Frame, c As Control
    For Each c In pgMove.Controls
        If TypeName(c) = "Frame" Then Set host = c: Exit For
    Next
    If host Is Nothing Then
        Set host = pgMove.Controls.Add("Forms.Frame.1", "frMoveHost")
        host.Left = 6: host.Top = 6
        host.Width = pgMove.InsideWidth - 12
        host.Height = pgMove.InsideHeight - 12
    End If

    ' 3) host 内の MultiPage だけを全消去（ほかは触らない）
    Dim i As Long
    For i = host.Controls.Count - 1 To 0 Step -1
        If TypeName(host.Controls(i)) = "MultiPage" Then
            host.Controls.Remove host.Controls(i).name
        End If
    Next

    ' 4) mpADL を作成＆3枚保証（0:BI / 1:IADL / 2:起居動作）
    Set mpADL = host.Controls.Add("Forms.MultiPage.1", "mpADL")
    Trace "mpADL ready; pages=" & mpADL.Pages.Count, "BI/IADL"

    With mpADL
        .Left = 12
        .Top = pad
        .Width = host.InsideWidth - 24
        .Height = 300
        .Style = fmTabStyleTabs
        AttachMPHook mpADL
    End With
    Do While mpADL.Pages.Count < 3: mpADL.Pages.Add: Loop
    mpADL.Pages(0).caption = "バーサルインデックス"
    mpADL.Pages(1).caption = "IADL"
    mpADL.Pages(2).caption = "起居動作"
    
    
Trace "EnsureBI_IADL end; pages=" & mpADL.Pages.Count, "BI/IADL"
Set EnsureBI_IADL = mpADL

    ' 5) 起居動作タブのUI
    mp.value = 0
    BuildKyoOnADL mpADL.Pages(2)


    '======================== BI（10項目） ========================
    Dim pBI As MSForms.Page: Set pBI = mpADL.Pages(0)
    ' 一旦クリアしてから作成（空／重複どちらにも対応）
    Dim iCtl As Long
    For iCtl = pBI.Controls.Count - 1 To 0 Step -1
        pBI.Controls.Remove pBI.Controls(iCtl).name
    Next

    Dim biItems As Variant, biChoices As Variant
    biItems = Array("摂食", "車いす-ベッド移乗", "整容", "トイレ動作", "入浴", _
                    "歩行/車いす移動", "階段昇降", "更衣", "排便コントロール", "排尿コントロール")
    biChoices = Array()


    Dim yBI As Single, idx As Long
    yBI = 18

    Dim lblBI As MSForms.label, txtBI As MSForms.TextBox
    Set lblBI = pBI.Controls.Add("Forms.Label.1", "lblBIHeader")
    With lblBI
        .caption = "バーサルインデックス（点）"
        .Left = 12: .Top = yBI: .Width = 160
    End With
    Set txtBI = pBI.Controls.Add("Forms.TextBox.1", "txtBITotal")
    With txtBI
        .tag = "BI.Total"
        .Left = lblBI.Left + lblBI.Width + 8
        .Top = yBI - 3
        .Width = 60
    End With
    yBI = yBI + rowH

    For idx = LBound(biItems) To UBound(biItems)
        Dim lb As MSForms.label, cB As MSForms.ComboBox
        Set lb = pBI.Controls.Add("Forms.Label.1", "lblBI_" & CStr(idx))
        With lb: .caption = CStr(biItems(idx)): .Left = 12: .Top = yBI: .Width = 160: End With

        Set cB = pBI.Controls.Add("Forms.ComboBox.1", "cmbBI_" & CStr(idx))
        AttachBIHook cB
        With cB
            .tag = "BI." & CStr(biItems(idx))
            .Left = 190
            .Top = yBI - 3
            .Width = 200
            .Style = fmStyleDropDownList
        End With
               ' 項目ごとにバーサル標準の点数候補を設定
        Select Case idx
            ' 0: 摂食
            ' 3: トイレ動作
            ' 6: 階段昇降
            ' 7: 更衣
            ' 8: 排便コントロール
            ' 9: 排尿コントロール
            ' → 0 / 5 / 10 点
            Case 0, 3, 6, 7, 8, 9
                AddItemsToCombo cB, Array("0", "5", "10")

            ' 2: 整容
            ' 4: 入浴
            ' → 0 / 5 点
            Case 2, 4
                AddItemsToCombo cB, Array("0", "5")

            ' 1: 車いす-ベッド移乗
            ' 5: 歩行/車いす移動
            ' → 0 / 5 / 10 / 15 点
            Case 1, 5
                AddItemsToCombo cB, Array("0", "5", "10", "15")
        End Select

        yBI = yBI + rowH
    Next idx

    pBI.ScrollBars = fmScrollBarsNone
    pBI.ScrollHeight = yBI + 8

    '======================== IADL（9項目） ========================
    Dim pIADL As MSForms.Page: Set pIADL = mpADL.Pages(1)
    For iCtl = pIADL.Controls.Count - 1 To 0 Step -1
        pIADL.Controls.Remove pIADL.Controls(iCtl).name
    Next

    Dim iadlItems As Variant, iadlChoices As Variant
    iadlItems = Array("調理", "洗濯", "掃除", "買い物", "金銭管理", "服薬管理", _
                      "趣味・余暇活動", "社会参加（外出・地域活動）", "コミュニケーション（電話・会話）")
    iadlChoices = Array("自立", "見守り（監視下）", "一部介助", "全介助")

    Dim iadlCount As Long, iadlCols As Long, iadlRows As Long
    iadlCount = UBound(iadlItems) - LBound(iadlItems) + 1
    iadlCols = 2
    iadlRows = (iadlCount + iadlCols - 1) \ iadlCols

    Dim colWIADL As Single: colWIADL = (mpADL.Width - 60) / iadlCols
    Dim j As Long, rowI As Long, colI As Long
    Dim xIADL As Single, yIADL As Single

    For j = LBound(iadlItems) To UBound(iadlItems)
        colI = (j - LBound(iadlItems)) \ iadlRows
        rowI = (j - LBound(iadlItems)) Mod iadlRows
        xIADL = 12 + colI * colWIADL
        yIADL = 18 + rowI * rowH

        Dim lb2 As MSForms.label, cb2 As MSForms.ComboBox
        Set lb2 = pIADL.Controls.Add("Forms.Label.1", "lblIADL_" & CStr(j))
        With lb2: .caption = CStr(iadlItems(j)): .Left = xIADL: .Top = yIADL: .Width = 120: End With

        Set cb2 = pIADL.Controls.Add("Forms.ComboBox.1", "cmbIADL_" & CStr(j))
        With cb2
            .tag = "IADL." & CStr(iadlItems(j))
            .Left = xIADL + 120
            .Top = yIADL - 3
            .Width = colWIADL - 150
            .Style = fmStyleDropDownList
        End With
        AddItemsToCombo cb2, iadlChoices
    Next j

    Dim gridBottom As Single: gridBottom = 18 + iadlRows * rowH
    Dim lblINote As MSForms.label, txtINote As MSForms.TextBox
    Set lblINote = pIADL.Controls.Add("Forms.Label.1", "lblIADLNote")
    With lblINote: .caption = "備考": .Left = 12: .Top = gridBottom + 12: .Width = 40: End With
    Set txtINote = pIADL.Controls.Add("Forms.TextBox.1", "txtIADLNote")
    With txtINote
        .tag = "IADL.備考"
        .Left = 60
        .Top = lblINote.Top - 3
        .Width = mpADL.Width - 84
        .Height = 60
        .multiline = True
        .EnterKeyBehavior = True
    End With

    '---- 高さ更新 ----
    Dim bottomI As Single: bottomI = txtINote.Top + txtINote.Height + 18
    pIADL.ScrollBars = fmScrollBarsNone
    pIADL.ScrollHeight = bottomI

    If mpADL.Height < bottomI + 42 Then mpADL.Height = bottomI + 42
    
   
    nextTop = mpADL.Top + mpADL.Height + 10
    If Not hostMove Is Nothing Then hostMove.ScrollHeight = nextTop + 10

    Set EnsureBI_IADL = mpADL
End Function
'=================================================================

'=== BIコンボにイベントフックを張る ===
Private Sub AttachBIHook(ByRef cB As MSForms.ComboBox)
    If BIHooks Is Nothing Then Set BIHooks = New Collection
    Dim h As CboBIHook
    Set h = New CboBIHook
    h.Init Me, cB
    BIHooks.Add h
End Sub

'=== BI：項目×選択肢 → 点数（Barthel標準に沿う） ==================
Private Function BIItemScore(ByVal itemName As String, ByVal level As String) As Long
    Select Case itemName
        Case "摂食"                         ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "車いす-ベッド移乗"            ' 15 / 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 15
                Case "見守り（監視下）":       BIItemScore = 10
                Case "一部介助":               BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "整容"                         ' 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "トイレ動作"                   ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "入浴"                         ' 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "歩行/車いす移動"              ' 15 / 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 15
                Case "見守り（監視下）":       BIItemScore = 10
                Case "一部介助":               BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "階段昇降"                     ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "更衣"                         ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "排便コントロール"             ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "排尿コントロール"             ' 10 / 5 / 0
            Select Case level
                Case "自立":                   BIItemScore = 10
                Case "見守り（監視下）", "一部介助": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case Else
            BIItemScore = 0
    End Select
End Function
'================================================================

Public Sub RecalcBI()
    Dim mpADL As MSForms.MultiPage
    Dim ctrl As MSForms.Control
    Dim pBI As MSForms.Page
    Dim idx As Long
    Dim cB As MSForms.ComboBox
    Dim total As Long
    Dim v As String
    Dim txt As MSForms.TextBox

    On Error Resume Next

    ' --- 「バーサルインデックス」タブを持つ MultiPage を探す ---
    Set mpADL = Nothing
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.MultiPage Then
            Set mpADL = ctrl
            If mpADL.Pages.Count > 0 Then
                If mpADL.Pages(0).caption = "バーサルインデックス" Then
                    Exit For
                End If
            End If
            Set mpADL = Nothing
        End If
    Next ctrl

    If mpADL Is Nothing Then Exit Sub

    Set pBI = mpADL.Pages(0)   ' 0ページ目がバーサルインデックス

    ' --- 10項目分の点数を単純に合計する（コンボの値がそのまま点数） ---
    total = 0
    For idx = 0 To 9
        Set cB = Nothing
        Set cB = pBI.Controls("cmbBI_" & CStr(idx))
        If Not cB Is Nothing Then
            v = Trim$(CStr(cB.value))
            If Len(v) > 0 Then
                total = total + CLng(val(v))
            End If
        End If
    Next idx

    ' --- 合計を txtBITotal に反映 ---
    Set txt = pBI.Controls("txtBITotal")
    If Not txt Is Nothing Then
        txt.value = total
    End If
End Sub


'=== 任意のTextBoxにIMEひらがなフックを張る ===
Private Sub AttachImeHiragana(tb As MSForms.TextBox)
    If ImeHooks Is Nothing Then Set ImeHooks = New Collection
    Dim h As TxtImeHook
    Set h = New TxtImeHook
    h.Init tb
    ImeHooks.Add h
    ' 念のため直ちに反映
    On Error Resume Next
    tb.IMEMode = fmIMEModeHiragana
End Sub

'=== MultiPage にフックを張る ===
Private Sub AttachMPHook(mp As MSForms.MultiPage)
    If MPHs Is Nothing Then Set MPHs = New Collection
    Dim h As MPHook
    Set h = New MPHook
    h.Init Me, mp
    MPHs.Add h
End Sub

'=== IADL備考にIMEひらがなを再適用（都度呼ぶ） ===
Public Sub ApplyImeToIADLNote()
    On Error Resume Next
     Dim mpA As MSForms.MultiPage, c As Control
     If hostMove Is Nothing Then Exit Sub

    For Each c In hostMove.Controls
        If TypeName(c) = "MultiPage" Then
            If c.name = "mpADL" Then Set mpA = c: Exit For
        End If
    Next c
    If mpA Is Nothing Then Exit Sub

    Dim tb As MSForms.TextBox
    Set tb = mpA.Pages(1).Controls("txtIADLNote") ' Page(1) = IADL
    If Not tb Is Nothing Then tb.IMEMode = fmIMEModeHiragana
End Sub

'=== どこかに残っている mpADL を全部消す ===
Private Sub RemoveAllMpADL()
    Dim i As Long, c As Control
    ' フォーム直下
    For i = Me.Controls.Count - 1 To 0 Step -1
        If TypeName(Me.Controls(i)) = "MultiPage" Then
            If Me.Controls(i).name = "mpADL" Then
                Me.Controls.Remove Me.Controls(i).name
            End If
        End If
    Next i

    ' ルート MultiPage（mp）の各ページ内
    Dim mp As MSForms.MultiPage, p As MSForms.Page
    For Each c In Me.Controls
        If TypeName(c) = "MultiPage" Then Set mp = c: Exit For
    Next c
    If Not mp Is Nothing Then
        For i = 0 To mp.Pages.Count - 1
            For Each c In mp.Pages(i).Controls
                If TypeName(c) = "MultiPage" Then
                    If c.name = "mpADL" Then mp.Pages(i).Controls.Remove c.name
                End If
            Next c
        Next i
    End If
End Sub


'=== ADLタブ内の3枚目「起居動作」ページを組み立てる ======================
Private Sub BuildKyoOnADL(pg As MSForms.Page)

    Dim fr As MSForms.Frame
    ' 既存があれば再利用、無ければ作成（FindOrAddFrameは既存ヘルパー）
    Set fr = FindOrAddFrame(pg, "frKyo")
    fr.caption = "起居動作"
    ClearChildren fr
    
     ' ★追加：位置とサイズ（ページいっぱいに広げる）
    With fr
        .Left = 12
        .Top = 12
        .Width = pg.parent.Width - 24   ' ← MultiPageの幅に追従
    End With
    

    Dim Y As Single: Y = 22
    Dim lb As MSForms.label, cB As MSForms.ComboBox, txt As MSForms.TextBox
    Dim choices As Variant

    ' 候補：既存の PostureChoices() があれば利用、無ければデフォルト
    On Error Resume Next
    choices = PostureChoices()
    If Err.Number <> 0 Then
        choices = Array("自立", "見守り（監視下）", "一部介助", "全介助")
        Err.Clear
    End If
    On Error GoTo 0
    
    
' 寝返り
Set lb = CreateLabel(fr, "寝返り", COL_LX, Y)
Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbKyo_Roll", True)
With cB: .Left = COL_LX + lblW + 60: .Top = Y - 3: .Width = 120: End With
AddItemsToCombo cB, choices
Y = Y + rowH

' 起き上がり
Set lb = CreateLabel(fr, "起き上がり", COL_LX, Y)
Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbKyo_SitUp", True)
With cB: .Left = COL_LX + lblW + 60: .Top = Y - 3: .Width = 120: End With
AddItemsToCombo cB, choices
Y = Y + rowH

' 座位保持
Set lb = CreateLabel(fr, "座位保持", COL_LX, Y)
Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbKyo_SitHold", True)
With cB: .Left = COL_LX + lblW + 60: .Top = Y - 3: .Width = 120: End With
AddItemsToCombo cB, choices
Y = Y + rowH

    
  ' 右列：立ち上がり / 立位保持（左列と同じ行Y=22,46／オフセット+60／幅120で揃える）
CreateLabel fr, "立ち上がり", COL_RX, 22
Dim cboUp As MSForms.ComboBox
Set cboUp = CreateCombo(fr, COL_RX + lblW + 60, 22, 120, , "POSTURE|立ち上がり")
cboUp.List = MakeList("自立,見守り（監視下）,一部介助,全介助")

CreateLabel fr, "立位保持", COL_RX, 46
Dim cboStand As MSForms.ComboBox
Set cboStand = CreateCombo(fr, COL_RX + lblW + 60, 46, 120, , "POSTURE|立位保持")
cboStand.List = MakeList("自立,見守り（監視下）,一部介助,全介助")

    ' 備考
    Set lb = CreateLabel(fr, "備考", COL_LX, Y)
    Set txt = fr.Controls.Add("Forms.TextBox.1", "txtKyoNote")
    With txt
        .Left = COL_LX + 40
        .Top = Y - 3
        .Width = fr.InsideWidth - (COL_LX + 40) - 12
        .Height = 120
        .multiline = True
        .EnterKeyBehavior = True
    End With
    Y = Y + txt.Height + 12

    fr.Height = Y + 12
End Sub

'ADL-起居動作用：コンボに候補をセット
Private Sub AddItemsToCombo(cB As MSForms.ComboBox, items As Variant)
    Dim k As Long
    On Error Resume Next
    cB.Clear
    cB.Style = fmStyleDropDownList
    For k = LBound(items) To UBound(items)
        cB.AddItem CStr(items(k))
    Next k
End Sub











'======================================================================
' 起居動作（「身体機能＋起居動作」ページ）? 安定テンプレ（備考あり）
' ・途中にタブを挿入しても壊れない：Caption検索
' ・生成（Build）と整列（Layout）を分離：拡張時の事故を最小化
' ・保存/読込は Tag="POSTURE|…" で既存ロジックにそのまま乗る
'======================================================================



'―― フォーム内の MultiPage を自動検出（名前に依存しない）
Private Function FindMainMultiPage() As MSForms.MultiPage
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If TypeOf c Is MSForms.MultiPage Then
            Set FindMainMultiPage = c
            Exit Function
        End If
    Next
End Function

'―― Caption に指定文字列を含むページを返す（無ければ Nothing）
Private Function FindPageByCaption(mp As MSForms.MultiPage, cap As String) As MSForms.Page
    Dim pg As MSForms.Page
    For Each pg In mp.Pages
        If InStr(pg.caption, cap) > 0 Then
            Set FindPageByCaption = pg
            Exit Function
        End If
    Next
End Function

'―― Page 内で Frame を取得（無ければ作成）
Private Function FindOrAddFrame(pg As MSForms.Page, nm As String) As MSForms.Frame
    Dim c As MSForms.Control
    For Each c In pg.Controls
        If TypeOf c Is MSForms.Frame Then
            If StrComp(c.name, nm, vbTextCompare) = 0 Then
                Set FindOrAddFrame = c
                Exit Function
            End If
        End If
    Next
    Set FindOrAddFrame = pg.Controls.Add("Forms.Frame.1", nm, True)
End Function

'―― 子コントロール全削除（生成前に一度だけ）
Private Sub ClearChildren(fr As MSForms.Frame)
    Dim i As Long
    For i = fr.Controls.Count - 1 To 0 Step -1
        fr.Controls.Remove fr.Controls(i).name
    Next
End Sub


Private Function PostureItems() As Variant
    PostureItems = Array("寝返り", "起き上がり", "座位保持", "立ち上がり", "立位保持")
End Function



Private Function PostureChoices() As Variant
    PostureChoices = Array("自立", "見守り（監視下）", "一部介助", "全介助")
End Function
'======================================================================

'―― 初回だけ“生成”する（Activate から呼ばれる）
Private Sub BuildPostureUI()

Debug.Print "BuildPostureUI CALLED", Join(PostureItems, " / ")

    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Sub

    ' 該当ページを Caption で取得（無ければ作る）
    Dim pg As MSForms.Page
    Set pg = FindPageByCaption(mp, CAP_POSTURE_PAGE)
    If pg Is Nothing Then
        Set pg = mp.Pages.Add
        pg.caption = CAP_POSTURE_PAGE
    End If
    mp.value = pg.Index   ' このタブを前面へ

    ' フレーム取得
    Dim fr As MSForms.Frame: Set fr = FindOrAddFrame(pg, "frPosture")

    ' 既存をクリア → 行生成
    ClearChildren fr
    CreatePostureRows fr    ' ラベル/コンボ/備考を作る（座標は Layout で調整）

    ' 生成直後に一度レイアウト
    LayoutPosture
End Sub

'―― ラベル＆コンボ行＋備考欄を作る（位置はここでは決めない）
Private Sub CreatePostureRows(fr As MSForms.Frame)
    Dim items As Variant:   items = PostureItems()
    Dim choices As Variant: choices = PostureChoices()

    Dim i As Long
    For i = LBound(items) To UBound(items)
        ' ラベル
        Dim lb As MSForms.label
        Set lb = fr.Controls.Add("Forms.Label.1", "lblPost_" & CStr(i), True)
        lb.caption = CStr(items(i))

        ' コンボ（保存/読込に使う Tag を付与）
        Dim cB As MSForms.ComboBox
        Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbPost_" & CStr(i), True)
        cB.Style = fmStyleDropDownList
        cB.tag = POSTURE_TAG_PREFIX & CStr(items(i))

        ' 選択肢の設定：共通関数があれば優先、無ければフォールバック
        On Error Resume Next
        SetComboItems cB, choices
        If Err.Number <> 0 Then
            Dim k As Long: Err.Clear
            For k = LBound(choices) To UBound(choices)
                cB.AddItem CStr(choices(k))
            Next
        End If
        On Error GoTo 0
    Next i

    '―― 備考ラベル＋テキストボックス（5行相当）
    Dim lbNote As MSForms.label
    Set lbNote = fr.Controls.Add("Forms.Label.1", "lblPost_Note", True)
    lbNote.caption = "備考"

    Dim txtNote As MSForms.TextBox
    Set txtNote = fr.Controls.Add("Forms.TextBox.1", "txtPost_Note", True)
    With txtNote
        .multiline = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .IMEMode = fmIMEModeHiragana   ' ← 日本語入力（全角）を明示
        .tag = POSTURE_TAG_PREFIX & "備考"
    End With
End Sub

'―― 位置・サイズ調整（Resize 毎に呼ぶ／再生成しない）
Private Sub LayoutPosture()
    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Sub

    Dim pg As MSForms.Page: Set pg = FindPageByCaption(mp, CAP_POSTURE_PAGE)
    If pg Is Nothing Then Exit Sub

    Dim fr As MSForms.Frame: Set fr = FindOrAddFrame(pg, "frPosture")

    ' フレームの基準寸法（MultiPage が 0 の場合は Form からフォールバック）
    fr.Visible = True
    fr.Left = 6: fr.Top = 6
    fr.Width = Application.Max(120, IIf(mp.Width > 0, mp.Width, Me.InsideWidth) - 12)
    fr.Height = Application.Max(120, IIf(mp.Height > 0, mp.Height, Me.InsideHeight) - 30)
    fr.ZOrder 0

    ' レイアウト（2列）
    Dim items As Variant: items = PostureItems()
    Dim cols As Long: cols = POSTURE_COLS
    Dim colW As Single: colW = Application.Max(120, (fr.Width - 24) / cols)
    Dim rows As Long: rows = (UBound(items) - LBound(items) + 1 + cols - 1) \ cols

    Dim startY As Single: startY = 12
    Dim rowH As Single:   rowH = 28
    Dim labelW As Single: labelW = Application.Max(60, colW - 110)

    Dim i As Long, c As Long, r As Long, x As Single, Y As Single
    For i = LBound(items) To UBound(items)
        c = (i - LBound(items)) \ rows
        r = (i - LBound(items)) Mod rows
        x = 12 + c * colW
        Y = startY + r * rowH

        With fr.Controls("lblPost_" & CStr(i))
            .Left = x
            .Top = Y + 3
            .Width = labelW
            .Visible = True
        End With
        With fr.Controls("cmbPost_" & CStr(i))
            .Left = x + labelW + 6
            .Top = Y
            .Width = 100
            .Visible = True
        End With
    Next i

    '―― 備考の位置とサイズ（5行分 ≒ 約90px）
    Dim noteTop As Single: noteTop = startY + rows * rowH + 10
    With fr.Controls("lblPost_Note")
        .Left = 12
        .Top = noteTop + 2
        .Width = 40
        .Visible = True
    End With
    With fr.Controls("txtPost_Note")
        .Left = 12 + 40 + 6
        .Top = noteTop
        .Width = fr.Width - 24 - 46
        .Height = 90          ' ← だいたい5行相当
        .Visible = True
    End With

    ' スクロール（備考の下端までをカバー）
    fr.ScrollBars = fmScrollBarsNone
    fr.ScrollHeight = noteTop + 90 + 12
    If fr.Height < fr.ScrollHeight Then fr.ScrollBars = fmScrollBarsVertical
    
    ' Call UserForm_Resize  ' NOTE: 直呼びすると mpPhys の高さが再計算され「全体が短い」再発源になるため禁止

    
    Debug.Print "[init] call phys root"
    'Call modPhysEval.EnsurePhysicalFunctionTabs_Under(Me, mp)
    

End Sub



'======================================================================
Public Sub RegisterMPHook(h As MPHook)
    mMPHooks.Add h
End Sub

Public Sub RegisterTxtHook(h As TxtImeHook)
    mTxtHooks.Add h
End Sub





'====================
' ヘルパー関数
'====================

Private Function RequiredOk() As Boolean
    On Error Resume Next
    RequiredOk = (Len(Trim$(Me.Controls("txtPID").Text)) > 0) _
        And (Len(Trim$(Me.Controls("txtName").Text)) > 0) _
        And IsNumeric(Me.Controls("txtAge").Text) _
        And (val(Me.Controls("txtAge").Text) >= 0)
End Function

Private Sub RefreshSaveEnabled()
    If Not btnSaveCtl Is Nothing Then btnSaveCtl.Enabled = RequiredOk()
End Sub

Private Sub txtPID_Change():  RefreshSaveEnabled: End Sub

Private Sub txtHdrName_Change()
     EnsureNameSuggestList

     Me.Controls("txtName").Text = Me.Controls("frHeader").Controls("txtHdrName").Text
     Me.UpdateNameSuggest

     UpdateNameSuggest
End Sub


Public Sub UpdateNameSuggest()

    Dim host As Object
    Dim tb As MSForms.TextBox
    Dim lb As MSForms.ListBox
    Dim ws As Worksheet
    Dim cName As Long, cID As Long
    Dim lastRow As Long, r As Long
    Dim key As String, keyN As String
    Dim nm As String, idv As String
    Dim hit As Long
    Dim seen As Object



    Set host = Me.Controls("frHeader")
    Set tb = host.Controls("txtHdrName")



    ' 候補リスト確保
    On Error Resume Next
        Dim i As Long
        Set lb = Nothing
           For i = Me.Controls.Count - 1 To 0 Step -1
           If Me.Controls(i).name = "lstNameSuggest" Then
        Set lb = Me.Controls(i)
        Exit For
    End If
Next i

    On Error GoTo 0

    If lb Is Nothing Then
        EnsureNameSuggestList
        Set lb = Me.Controls("lstNameSuggest")
    End If


    key = Trim$(tb.Text)
    keyN = NormalizeName(key)

    lb.Clear
    lb.Visible = False
    If Len(keyN) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(EVAL_SHEET_NAME)

    ' 氏名列（既存ロジックと同じ探し方）
    cName = FindHeaderColLocal(ws, "Basic.Name")
    If cName = 0 Then cName = FindHeaderColLocal(ws, "氏名")
    If cName = 0 Then cName = FindHeaderColLocal(ws, "Name")
    If cName = 0 Then Exit Sub

    ' ID列（あれば併記）
    cID = FindHeaderColLocal(ws, "Basic.ID")
    If cID = 0 Then cID = FindHeaderColLocal(ws, "ID")
    If cID = 0 Then cID = FindHeaderColLocal(ws, "PID")


    lastRow = ws.Cells(ws.rows.Count, cName).End(xlUp).row

    ' 2列にして、2列目（ID）は非表示運用（表示文字列に併記する）
    lb.ColumnCount = 2
    lb.ColumnWidths = CStr(lb.Width - 2) & ";0"
    
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare


    For r = 2 To lastRow
        nm = CStr(ws.Cells(r, cName).value)
        If Len(nm) > 0 Then
            If InStr(1, NormalizeName(nm), keyN, vbTextCompare) > 0 Then
                idv = ""
                If cID > 0 Then idv = CStr(ws.Cells(r, cID).value)
                
                
                Dim nmKey As String
                nmKey = NormalizeName(nm)

                If Not seen.exists(nmKey) Then
                seen.Add nmKey, True
                

                lb.AddItem nm
                lb.List(lb.ListCount - 1, 1) = idv

                hit = hit + 1
                If hit >= 20 Then Exit For
                End If

                
            End If
        End If
    Next r

    If hit > 0 Then lb.Visible = True
    
    If lb.ListCount <= 1 Then
        mDupNameWarned = False
    ElseIf lb.ListCount > 1 Then
       If Not mDupNameWarned Then
           MsgBox "同姓同名の候補が複数あります。必要ならIDで絞り込みしてください。", vbInformation
           mDupNameWarned = True
      End If
    End If



End Sub

Private Function FindHeaderColLocal(ws As Worksheet, headerText As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If CStr(ws.Cells(1, c).value) = headerText Then
            FindHeaderColLocal = c
            Exit Function
        End If
    Next c
End Function



Private Sub txtAge_Change():  RefreshSaveEnabled: End Sub





'========================
' スクロール無しのホストフレーム
'========================
Private Function CreateScrollHost(pg As MSForms.Page) As MSForms.Frame
    Dim host As MSForms.Frame
    Set host = pg.Controls.Add("Forms.Frame.1")

    With host
        .caption = ""
        .Left = 0
        .Top = 0
        .Width = pg.parent.Width - 12
        .Height = pg.parent.Height - 12
        .ScrollBars = fmScrollBarsNone
        .ScrollWidth = .Width
    End With

    Set CreateScrollHost = host
End Function


'========================
' シート（EvalData） ※frmEvalローカル版
'========================
Private Function EnsureEvalData() As Worksheet
    Const sh As String = "EvalData"
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sh)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = sh
    End If

    Set EnsureEvalData = ws
End Function

Private Sub EnsureJapaneseHeaderRow(ws As Worksheet)
    If Application.WorksheetFunction.CountA(ws.rows(2)) = 0 Then
        Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim c As Long
        For c = 1 To lastCol
            Select Case CStr(ws.Cells(1, c).value)
                Case "PatientID":        ws.Cells(2, c).value = "ID"
                Case "EvalDate":         ws.Cells(2, c).value = "評価日"
                Case "Basic.Name":       ws.Cells(2, c).value = "氏名"
                Case "Basic.Age":        ws.Cells(2, c).value = "年齢"
                Case "Basic.Gender":     ws.Cells(2, c).value = "性別"
                Case "Basic.PrimaryDx":  ws.Cells(2, c).value = "主診断"
                Case "Basic.OnsetDate":  ws.Cells(2, c).value = "発症日"
                Case "Basic.Living":     ws.Cells(2, c).value = "生活状況"
                Case "Basic.CareLevel":  ws.Cells(2, c).value = "要介護度"
                ' 必要に応じて追加
            End Select
        Next
    End If
End Sub

'========================
' 直近行探索（PID）
'========================
Private Function FindLastRowByPID(ByVal pid As String, ByVal ws As Worksheet) As Long
    Dim colPID As Variant, colDate As Variant, colTS As Variant
    colPID = Application.Match("PatientID", ws.rows(1), 0)
    colDate = Application.Match("EvalDate", ws.rows(1), 0)
    colTS = Application.Match("Timestamp", ws.rows(1), 0)

    If IsError(colPID) Or IsError(colDate) Then Exit Function

    Dim last As Long: last = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    Dim bestRow As Long
    Dim bestD As Date, bestHasTs As Boolean, bestTs As Date

    Dim r As Long
    For r = 2 To last
        If CStr(ws.Cells(r, CLng(colPID)).value) = pid Then
            If IsDate(ws.Cells(r, CLng(colDate)).value) Then
                Dim d As Date: d = CDate(ws.Cells(r, CLng(colDate)).value)

                Dim hasTs As Boolean, t As Date
                hasTs = (Not IsError(colTS)) And IsDate(ws.Cells(r, CLng(colTS)).value)
                If hasTs Then t = CDate(ws.Cells(r, CLng(colTS)).value)

                If bestRow = 0 Then
                    bestRow = r: bestD = d: bestHasTs = hasTs: If hasTs Then bestTs = t
                ElseIf d > bestD Then
                    bestRow = r: bestD = d: bestHasTs = hasTs: If hasTs Then bestTs = t
                ElseIf d = bestD Then
                    If hasTs And Not bestHasTs Then
                        bestRow = r: bestHasTs = True: bestTs = t
                    ElseIf hasTs And bestHasTs Then
                        If t > bestTs Then
                            bestRow = r: bestTs = t
                        ElseIf t = bestTs Then
                            If r > bestRow Then bestRow = r
                        End If
                    ElseIf (Not hasTs) And (Not bestHasTs) Then
                        If r > bestRow Then bestRow = r
                    End If
                End If
            End If
        End If
    Next

    FindLastRowByPID = bestRow
End Function

'========================
' RLA レベル取得
'========================
Private Function GetRLAGroupLevel(ByVal grp As String) As String
    Dim c As MSForms.Control, p As MSForms.Page, fr As MSForms.Control, ob As MSForms.Control
    For Each c In hostWalk.Controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each fr In p.Controls
                    If TypeName(fr) = "Frame" Then
                        For Each ob In fr.Controls
                            If TypeName(ob) = "OptionButton" Then
                                If ob.groupName = grp And ob.value Then GetRLAGroupLevel = ob.caption: Exit Function
                            End If
                        Next
                    End If
                Next
            Next
        End If
    Next
End Function

'========================
' 収集
'========================
Private Function CollectFormData() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As MSForms.Control, oc As MSForms.Control, ic As MSForms.Control

    Dim j As Long
    For Each c In Me.Controls
        Select Case TypeName(c)
            Case "MultiPage"
                For j = 0 To c.Pages.Count - 1
                    Dim p As MSForms.Page: Set p = c.Pages(j)
                    Dim co As MSForms.Control
                    For Each co In p.Controls
                        CollectOne d, co
                        If TypeName(co) = "Frame" Then
                            For Each ic In co.Controls: CollectOne d, ic: Next
                        ElseIf TypeName(co) = "MultiPage" Then
                            Dim p2 As MSForms.Page, fr As MSForms.Control, it As MSForms.Control
                            For Each p2 In co.Pages
                                For Each fr In p2.Controls
                                    CollectOne d, fr
                                    If TypeName(fr) = "Frame" Then
                                        For Each it In fr.Controls: CollectOne d, it: Next
                                    End If
                                Next
                            Next
                        End If
                    Next
                Next
            Case Else
                CollectOne d, c
        End Select
    Next

    d("Basic.Assistive") = AggregateChecks("AssistiveGroup")
    d("Basic.Risks") = AggregateChecks("RiskGroup")

    Dim frames As Collection: Set frames = FindAllFramesByCaptionPart("RLA歩行周期")
    Dim keys As Variant: keys = Array("IC", "LR", "MSt", "TSt", "PSw", "ISw", "MSw", "TSw")
    Dim f As MSForms.Frame, k As Variant, s As String
    For Each k In keys
        s = ""
        For Each f In frames
            Dim part As String: part = BuildRLAString(f, CStr(k))
            If Len(part) > 0 Then s = part
        Next
        d("RLA." & CStr(k)) = s
        d("RLA." & CStr(k) & ".Level") = GetRLAGroupLevel(CStr(k))
    Next

    Set CollectFormData = d
End Function

Private Sub CollectOne(ByRef d As Object, ByVal ctl As MSForms.Control)
    If Len(ctl.tag & "") = 0 Then Exit Sub
    Select Case TypeName(ctl)
        Case "TextBox", "ComboBox": d(ctl.tag) = ctl.Text
        Case "CheckBox"
            If ctl.tag <> "AssistiveGroup" And ctl.tag <> "RiskGroup" Then
                d(ctl.tag) = IIf(ctl.value, "有", "無")
            End If
    End Select
End Sub

Private Function AggregateChecks(ByVal groupTag As String) As String
    Dim picks As String, c As MSForms.Control, p As MSForms.Page, fr As MSForms.Control, cc As MSForms.Control
    For Each c In Me.Controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each fr In p.Controls
                    If TypeName(fr) = "Frame" Then
                        For Each cc In fr.Controls
                            If TypeName(cc) = "CheckBox" Then
                                If cc.tag = groupTag And cc.value Then picks = IIf(Len(picks) = 0, cc.caption, picks & "/" & cc.caption)
                            End If
                        Next
                    End If
                Next
            Next
        End If
    Next
    AggregateChecks = picks
End Function

Private Function BuildRLAString(ByVal f As MSForms.Frame, ByVal key As String) As String
    Dim c As MSForms.Control, acc As String
    For Each c In f.Controls
        If TypeName(c) = "CheckBox" Then
            If Left$(c.name, 4) = "RLA_" And Mid$(c.name, 5, Len(key)) = key Then
                If c.value Then acc = IIf(Len(acc) = 0, c.caption, acc & "/" & c.caption)
            End If
        End If
    Next
    BuildRLAString = acc
End Function

Private Function FindAllFramesByCaptionPart(ByVal part As String) As Collection
    Dim col As New Collection
    Dim c As MSForms.Control, p As MSForms.Page, oc As MSForms.Control
    For Each c In Me.Controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each oc In p.Controls
                    If TypeName(oc) = "Frame" Then
                        If InStr(1, oc.caption, part, vbTextCompare) > 0 Then col.Add oc
                    ElseIf TypeName(oc) = "MultiPage" Then
                        Dim p2 As MSForms.Page, oc2 As MSForms.Control
                        For Each p2 In oc.Pages
                            For Each oc2 In p2.Controls
                                If TypeName(oc2) = "Frame" Then
                                    If InStr(1, oc2.caption, part, vbTextCompare) > 0 Then col.Add oc2
                                End If
                            Next
                        Next
                    End If
                Next
            Next
        End If
    Next
    Set FindAllFramesByCaptionPart = col
End Function

'========================
' 入力チェック
'========================
Private Function CheckRange(frm As Object, ByVal nm As String, ByVal lo As Double, ByVal hi As Double, ByVal message As String, ByRef sb As String) As Boolean
    If Not FnHasControl(nm) Then CheckRange = True: Exit Function
    Dim t As String: t = Trim$(frm.Controls(nm).Text & "")
    If t = "" Then CheckRange = True: Exit Function
    If Not IsNumeric(t) Then sb = sb & "・" & message & vbCrLf: CheckRange = False: Exit Function
    Dim v As Double: v = CDbl(t)
    If v < lo Or v > hi Then sb = sb & "・" & message & vbCrLf: CheckRange = False: Exit Function
    CheckRange = True
End Function

Private Function ValidateForm(ByRef errmsg As String) As Boolean
    Dim ok As Boolean: ok = True
    Dim sb As String: sb = ""

    If FnHasControl("txtName") Then If Trim$(Me.Controls("txtName").Text) = "" Then ok = False: sb = sb & "・氏名を入力してください。" & vbCrLf
    If FnHasControl("txtAge") Then
        If Trim$(Me.Controls("txtAge").Text) = "" Or Not IsNumeric(Me.Controls("txtAge").Text) Then
            ok = False: sb = sb & "・年齢を数値で入力してください。" & vbCrLf
        ElseIf val(Me.Controls("txtAge").Text) < 0 Or val(Me.Controls("txtAge").Text) > 120 Then
            ok = False: sb = sb & "・年齢は0～120で入力してください。" & vbCrLf
        End If
    End If

    ' 評価日チェック
    If FnHasControl("txtEDate") Then
        If Not IsDate(Me.Controls("txtEDate").Text) Then
            ok = False: sb = sb & "・評価日を正しい日付（yyyy/mm/dd 等）で入力してください。" & vbCrLf
        End If
    End If

    ok = ok And CheckRange(Me, "txtTenMWalk", 0, 300, "10m歩行（秒）は0～300で入力してください。", sb)
    ok = ok And CheckRange(Me, "txtTUG", 0, 300, "TUG（秒）は0～300で入力してください。", sb)
    ok = ok And CheckRange(Me, "txtFiveSTS", 0, 300, "5回立ち上がり（秒）は0～300で入力してください。", sb)
    ok = ok And CheckRange(Me, "txtSemi", 0, 300, "セミタンデム（秒）は0～300で入力してください。", sb)
    ok = ok And CheckRange(Me, "txtGripR", 0, 120, "握力 右（kg）は0～120で入力してください。", sb)
    ok = ok And CheckRange(Me, "txtGripL", 0, 120, "握力 左（kg）は0～120で入力してください。", sb)

    errmsg = IIf(ok, "", sb)
    ValidateForm = ok
End Function

Private Sub btnSaveCtl_Click()
    Call SyncAgeFromBirth
    Me.Controls("txtName").Text = Me.Controls("txtHdrName").Text
    SaveEvaluation_Append_From Me
End Sub



'=== frmEval：前回読込ボタン 完全貼り替え ============================
Private Sub btnLoadPrevCtl_Click()

Me.Controls("txtName").Text = Me.Controls("txtHdrName").Text

Dim pname As String
On Error Resume Next
pname = Trim$(Me.Controls("txtName").Text)
On Error GoTo 0
If Len(pname) = 0 Then
    MsgBox "氏名を入力してください。", vbExclamation
    Exit Sub
End If

Dim ws As Worksheet: Set ws = EnsureEvalData()
Dim look As Object: Set look = BuildHeaderLookup(ws)

look("IO_Sensory") = HeaderCol_Compat("IO_Sensory", ws)
look("IO_ADL") = HeaderCol_Compat("IO_ADL", ws)
look("IO_MMT") = HeaderCol_Compat("IO_MMT", ws)
look("IO_Tone") = HeaderCol_Compat("IO_Tone", ws)

' ROMは複数列のため次手で別処理
Dim r As Long
ws.Activate
Dim cName As Long: cName = HeaderCol_Compat("氏名", ws)
If cName = 0 Then cName = HeaderCol_Compat("利用者名", ws)
If cName = 0 Then cName = HeaderCol_Compat("名前", ws)
If cName = 0 Then
    MsgBox "氏名列が見つかりません。", vbExclamation
    Exit Sub
End If

Dim rr As Long
For rr = ws.Cells(ws.rows.Count, cName).End(xlUp).row To 2 Step -1
    If Trim$(CStr(ws.Cells(rr, cName).value)) = pname Then
        r = rr
        Exit For
    End If
Next
If r = 0 Then
    MsgBox "前回データが見つかりません（氏名一致なし）: " & pname, vbExclamation
    Exit Sub
End If



'If r <= 0 Then Exit Sub


If r > 0 Then
    Application.Run "LoadSensoryFromSheet", ThisWorkbook.Worksheets("EvalData"), r, Me
    Application.Run "LoadMMTFromSheet", ThisWorkbook.Worksheets("EvalData"), r, Me
    Application.Run "LoadLatestPainNow"
    Application.Run "Load_ADL_Latest"
End If

    Call modEvalIOEntry.LoadEvaluation_ByName_From(Me)

' [Pain-Load] manual wrapper entrypoint (DO NOT REMOVE)
Call LoadLatestPainNow

Me.Repaint

    Exit Sub

End Sub

Public Sub HandleHdrLoadPrevClick()
    
    Call btnLoadPrevCtl_Click
End Sub

Private Sub cmdHdrLoadPrev_Click()
    Call btnLoadPrevCtl_Click
End Sub


' 下から遡って氏名一致の「最新?最大5件」を集め、
' 件数=1ならそれを返し、2以上なら番号選択のInputBoxを出す
Private Function FindRowByNameWithPickLocal(ws As Worksheet, nameText As String, Optional maxCount As Long = 5) As Long
    Dim colName As Long, colDate As Long
    colName = modEvalIOEntry.FindColByHeaderExact(ws, "氏名")
    If colName = 0 Then colName = modEvalIOEntry.FindColByHeaderExact(ws, "利用者名")
    If colName = 0 Then colName = modEvalIOEntry.FindColByHeaderExact(ws, "名前")
    If colName = 0 Then Exit Function

    colDate = modEvalIOEntry.FindColByHeaderExact(ws, "評価日")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "記録日")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "更新日")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "作成日")

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, colName).End(xlUp).row
    Dim rows() As Long, cnt As Long, r As Long
    ReDim rows(1 To maxCount)

    For r = lastRow To 2 Step -1
        If ws.Cells(r, colName).value = nameText Then
            cnt = cnt + 1
            rows(cnt) = r
            If cnt = maxCount Then Exit For
        End If
    Next r

    If cnt = 0 Then Exit Function
    If cnt = 1 Then
        FindRowByNameWithPickLocal = rows(1)
        Exit Function
    End If

    Dim disp As String, i As Long, d As Variant
    For i = 1 To cnt
        If colDate > 0 Then
            d = ws.Cells(rows(i), colDate).value
            If IsDate(d) Then
                disp = disp & i & ") " & Format$(d, "yyyy-mm-dd") & "   Row:" & rows(i) & vbCrLf
            Else
                disp = disp & i & ") （日付なし）  Row:" & rows(i) & vbCrLf
            End If
        Else
            disp = disp & i & ") Row:" & rows(i) & vbCrLf
        End If
    Next i

    Dim idx As Variant
    idx = Application.InputBox( _
            Prompt:="同姓同名の直近" & cnt & "件（最新→古い）:" & vbCrLf & disp & vbCrLf & _
                    "番号を入力（1-" & cnt & "、キャンセル=中止）", _
            Type:=1)
    If idx = False Then Exit Function
    If idx >= 1 And idx <= cnt Then
        FindRowByNameWithPickLocal = rows(CLng(idx))
    End If
End Function



'=== 候補一覧を表示して番号入力で1つ選んでもらう ======================
Public Function PickCandidateRowByNameLocal(ByVal ws As Worksheet, _
                                             ByVal look As Object, _
                                             ByVal candidates As Variant, _
                                             ByVal pname As String) As Long
    
    Debug.Print "[ENTER] PickCandidateRowByNameLocal", Timer

    
    If IsEmpty(candidates) Then Exit Function
    If Not IsArray(candidates) Then Exit Function

    Dim pidCol As Long, ageCol As Long, dtCol As Long
    pidCol = RCol(ws, look, "Basic.ID", "ID", "個人ID")
ageCol = RCol(ws, look, "Basic.Age", "年齢")
dtCol = RCol(ws, look, "Basic.EvalDate", "評価日", "記録日", "更新日", "作成日")

    If dtCol = 0 Then dtCol = ResolveColumnLocal(look, "評価日")
    If dtCol = 0 Then dtCol = ResolveColumnLocal(look, "EvalDate")

    Dim lb As Long, ub As Long, cnt As Long
    lb = LBound(candidates): ub = UBound(candidates): cnt = ub - lb + 1
    Debug.Print "[CANDS] name=", pname, " cnt=", cnt, " range=", lb & "-" & ub



    Dim i As Long, r As Long, msg As String, disp As Long
    msg = "同姓同名が見つかりました。読み込むデータを選択してください（番号を入力）。" & vbCrLf & vbCrLf
    For i = lb To ub
        r = candidates(i)
        disp = i - lb + 1
        msg = msg & CStr(disp) & ") "
        If dtCol > 0 Then
            If IsDate(ws.Cells(r, dtCol).value) Then
                msg = msg & Format$(CDate(ws.Cells(r, dtCol).value), "yyyy/mm/dd")
            Else
                msg = msg & CStr(ws.Cells(r, dtCol).value)
            End If
        Else
            msg = msg & "(評価日なし)"
        End If
        If ageCol > 0 Then msg = msg & " | 年齢:" & Trim$(CStr(ws.Cells(r, ageCol).value))
        If pidCol > 0 Then msg = msg & " | ID:" & Trim$(CStr(ws.Cells(r, pidCol).value))
        msg = msg & vbCrLf
    Next i

   ' 入力取得（数値のみ）
Dim sel As Variant
sel = Application.InputBox(msg, "候補選択 - " & pname, Type:=1)

' Cancel / エラー / 非数値 / 空 を弾く（短絡で判定）
If VarType(sel) = vbBoolean Then Exit Function   ' Cancel
If IsError(sel) Then Exit Function               ' まれに CVErr
If Not IsNumeric(sel) Then Exit Function
If Len(CStr(sel)) = 0 Then Exit Function


    Dim n As Long: n = CLng(sel)
    If n < 1 Or n > cnt Then Exit Function

    PickCandidateRowByNameLocal = candidates(lb + n - 1)
End Function


Private Function NzTxt(tb As MSForms.TextBox) As String
    On Error Resume Next
    NzTxt = ""
    If Not tb Is Nothing Then NzTxt = Trim$(tb.Text)
End Function

Private Sub SetComboByValue(ByVal cbo As MSForms.ComboBox, ByVal v As String)
    Dim i As Long, idx As Long: idx = -1
    For i = 0 To cbo.ListCount - 1
        If CStr(cbo.List(i)) = CStr(v) Then idx = i: Exit For
    Next
    If idx >= 0 Then
        cbo.ListIndex = idx
    Else
        If Len(v) > 0 Then cbo.AddItem CStr(v)
        If cbo.ListCount > 0 Then cbo.ListIndex = cbo.ListCount - 1
    End If
End Sub



'=== ここから補助関数群（frmEval ローカル） ============================
Private Function FxGetText(ByVal ctrlName As String) As String
    On Error Resume Next
    FxGetText = Trim$(Me.Controls(ctrlName).Text)
End Function

Private Sub FxSetText(ByVal ctrlName As String, ByVal value As String)
    On Error Resume Next
    Me.Controls(ctrlName).Text = value
End Sub

Private Function GetOrCreateEvalSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.name = "EvalData"
    End If
    Set GetOrCreateEvalSheet = ws
End Function

Private Function BuildHeaderLookupLocal(ByVal ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 'TextCompare

    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, k1 As String, k2 As String
    For c = 1 To IIf(lastCol > 0, lastCol, 1)
        k1 = NormalizeKeyLocal(CStr(ws.Cells(1, c).value))
        k2 = NormalizeKeyLocal(CStr(ws.Cells(2, c).value))
        If Len(k1) > 0 Then If Not dict.exists(k1) Then dict.Add k1, c
        If Len(k2) > 0 Then If Not dict.exists(k2) Then dict.Add k2, c
    Next
    Set BuildHeaderLookupLocal = dict
End Function

Private Function ResolveColumnLocal(ByVal look As Object, ByVal key As String) As Long
    Dim k As String: k = NormalizeKeyLocal(key)
    If Len(k) = 0 Then Exit Function
    If look.exists(k) Then ResolveColumnLocal = CLng(look(k))
End Function

Private Function EnsureHeaderColumnLocal(ByVal ws As Worksheet, ByVal look As Object, ByVal key As String) As Long
    Dim col As Long: col = ResolveColumnLocal(look, key)
    If col > 0 Then EnsureHeaderColumnLocal = col: Exit Function

    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 And Len(CStr(ws.Cells(1, 1).value)) = 0 Then lastCol = 0

    col = lastCol + 1
    ws.Cells(1, col).value = key
    ws.Cells(2, col).value = key
    look(NormalizeKeyLocal(key)) = col
    EnsureHeaderColumnLocal = col
End Function

Private Sub WriteCellByKey(ByVal ws As Worksheet, ByVal look As Object, ByVal row As Long, _
                           ByVal key As String, ByVal val As Variant, ByVal createIfMissing As Boolean)
    Dim col As Long: col = ResolveColumnLocal(look, key)
    If col = 0 And createIfMissing Then
        col = EnsureHeaderColumnLocal(ws, look, key)
    End If
    If col > 0 Then ws.Cells(row, col).value = val
End Sub

Private Function FindLastRowByPIDLocal(ByVal ws As Worksheet, ByVal look As Object, ByVal pid As String) As Long
    Dim idCol As Long: idCol = EnsureHeaderColumnLocal(ws, look, "PatientID")
    Dim tsCol As Long: tsCol = EnsureHeaderColumnLocal(ws, look, "Timestamp")
    Dim dtCol As Long: dtCol = EnsureHeaderColumnLocal(ws, look, "Basic.EvalDate")

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, idCol).End(xlUp).row
    Dim r As Long, pick As Long, best As Date, curTs As Date, curDt As Date

    For r = 3 To lastRow
        If CStr(ws.Cells(r, idCol).value) = pid Then
            curTs = 0: curDt = 0
            On Error Resume Next
            curTs = CDate(ws.Cells(r, tsCol).value)
            curDt = CDate(ws.Cells(r, dtCol).value)
            On Error GoTo 0
            If curTs > 0 Then
                If curTs > best Then best = curTs: pick = r
            ElseIf curDt > 0 Then
                If curDt > best Then best = curDt: pick = r
            End If
        End If
    Next
    FindLastRowByPIDLocal = pick
End Function

'=== 氏名で最新行を探す（ID版の完全互換ロジック） =====================
Private Function FindLastRowByNameLocal(ByVal ws As Worksheet, _
                                        ByVal look As Object, _
                                        ByVal pname As String) As Long
    ' 見出しの列位置（無ければ作成して揃える：ID版と同じ方針）
    Dim nameCol As Long: nameCol = ResolveColOrCreate(ws, look, "Basic.Name", "氏名", "Name")
    Dim tsCol As Long:   tsCol = ResolveColOrCreate(ws, look, "Timestamp")
    Dim dtCol As Long:   dtCol = ResolveColOrCreate(ws, look, "Basic.EvalDate", "評価日", "EvalDate")

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, nameCol).End(xlUp).row
    Dim r As Long, pick As Long, best As Date, curTs As Date, curDt As Date

    For r = 3 To lastRow
        If KeyNormalize(ws.Cells(r, nameCol).value) = KeyNormalize(pname) Then

            curTs = 0: curDt = 0
            On Error Resume Next
            curTs = CDate(ws.Cells(r, tsCol).value)
            curDt = CDate(ws.Cells(r, dtCol).value)
            On Error GoTo 0
            If curTs > 0 Then
                If curTs > best Then best = curTs: pick = r
            ElseIf curDt > 0 Then
                If curDt > best Then best = curDt: pick = r
            End If
        End If
    Next
    FindLastRowByNameLocal = pick
End Function
'====================================================================


Private Function NextDataRowLocal(ByVal ws As Worksheet) As Long
    Dim r As Long: r = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If r < 3 Then r = 2
    NextDataRowLocal = r + 1
End Function

Private Function GeneratePIDLocal() As String
    GeneratePIDLocal = "PID-" & Format(Now, "yyyymmdd-hhnnss")
End Function

Private Function NormalizeKeyLocal(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "　", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeKeyLocal = LCase$(s)
End Function
'=== 補助関数群 ここまで ===============================================

Private Sub btnCloseCtl_Click()

    RequestQuitExcelAskAndCloseForm
    Unload Me
End Sub


Private Sub SetImeRecursive(container As Object)
    Dim ctl As Control
    For Each ctl In container.Controls
        If TypeName(ctl) = "TextBox" Then ctl.IMEMode = fmIMEModeHiragana

        Select Case TypeName(ctl)
            Case "Frame", "UserForm"
                SetImeRecursive ctl

            Case "MultiPage"
                Dim p As MSForms.Page
                For Each p In ctl.Pages          ' ★ Pages を列挙するのが重要
                    SetImeRecursive p
                Next
        End Select
    Next
End Sub

Private Sub FixRestNRS_Once()
    Dim l As Control, c As Control
    On Error Resume Next
    Set c = Me.Controls("cmbNRS_Move")                  ' 動作時NRSのコンボ
    If c Is Nothing Then Exit Sub

    ' 近い高さにあるラベルを探す（見つからなければ新規に作る）
    For Each l In Me.Controls
        If TypeName(l) = "Label" Then
            If l.caption = "安静時NRS" Or (Abs(l.Top - c.Top) <= 20 And l.Left < c.Left) Then
                Exit For
            End If
        End If
    Next
    If l Is Nothing Then
        Set l = c.parent.Controls.Add("Forms.Label.1", "lblNRS_Rest", True)
    End If

    ' 表示・位置・サイズを確定
    l.caption = "安静時NRS"
    l.Visible = True
    l.WordWrap = False
    l.Width = 72
    l.Top = c.Top
    l.Left = c.Left - (l.Width + 6)
    l.ZOrder 0
End Sub










'=== 追加：MultiPageの既定ページを掃除（"Page*" を全部削除） ===
Private Sub CleanDefaultPages(mp As MSForms.MultiPage)
    On Error Resume Next
    Dim i As Long
    ' キャプションが "Page" で始まるページを後ろから削除
    For i = mp.Pages.Count - 1 To 0 Step -1
        If Left$(mp.Pages(i).caption, 4) = "Page" Then
            mp.Pages.Remove i
        End If
    Next
End Sub


' 指定キャプションのボタンを再帰で探す
Private Function FindButtonByCaption(container As Object, ByVal cap As String) As MSForms.CommandButton
    Dim c As Object, hit As MSForms.CommandButton
    For Each c In container.Controls
        If TypeName(c) = "CommandButton" Then
            If CStr(c.caption) = cap Then
                Set FindButtonByCaption = c
                Exit Function
            End If
        End If
        If TypeOf c Is MSForms.Frame Or TypeOf c Is MSForms.MultiPage Or TypeName(c) = "Page" Then
            Set hit = FindButtonByCaption(c, cap)
            If Not hit Is Nothing Then Set FindButtonByCaption = hit: Exit Function
        End If
    Next
End Function

Private Sub HookHeaderButtons()
    On Error Resume Next
    Set btnHdrSave = Nothing
    Set btnHdrLoadPrev = Nothing

    ' まずキャプション部分一致で
    '=== まずキャプション一致（定数で厳密に） ===
Const CAP_SAVE As String = "シートへ保存"          ' ← cap=[ ] の中身だけ
Const CAP_LOAD As String = "前回の値を読み込む"    ' ← cap=[ ] の中身だけ

Debug.Print "[TEST] 保存like:", TypeName(FindButtonByCaptionLike(Me, "保存"))
Debug.Print "[TEST] 読み込like:", TypeName(FindButtonByCaptionLike(Me, "読み込"))


Set btnHdrSave = FindButtonByCaptionLike(Me, CAP_SAVE)
Set btnHdrLoadPrev = FindButtonByCaptionLike(Me, CAP_LOAD)


    If btnHdrLoadPrev Is Nothing Then Set btnHdrLoadPrev = FindButtonByCaptionLike(Me, "読")

    ' 見つからなければ最上段の右側２つを採用
    If (btnHdrSave Is Nothing) Or (btnHdrLoadPrev Is Nothing) Then
        FindHeaderButtonsByPosition Me, btnHdrSave, btnHdrLoadPrev
    End If

    Debug.Print "[HOOK] save=", IIf(btnHdrSave Is Nothing, "NG", "OK"), _
                "    load=", IIf(btnHdrLoadPrev Is Nothing, "NG", "OK")
End Sub





'=== ボタンをキャプション部分一致で探す（Frame/MultiPage/Page を再帰）===
Public Function FindButtonByCaptionLike(container As Object, _
                                        ByVal needle As String) As Object
    On Error GoTo SafeExit

    Dim c  As Object
    Dim pg As MSForms.Page
    Dim hit As Object

    ' MultiPage は Pages 経由で潜る（ここが重要）
    If TypeName(container) = "MultiPage" Then
        For Each pg In container.Pages
            Set hit = FindButtonByCaptionLike(pg, needle)
            If Not hit Is Nothing Then
                Set FindButtonByCaptionLike = hit
                Exit Function
            End If
        Next
        Exit Function
    End If

    ' それ以外は Controls を走査
    For Each c In container.Controls
        If TypeName(c) = "CommandButton" Then
            If InStr(Replace$(c.caption, vbCrLf, ""), needle) > 0 Then
                Set FindButtonByCaptionLike = c
                Exit Function
            End If
        ElseIf TypeName(c) = "Frame" Or TypeName(c) = "Page" Then
            Set hit = FindButtonByCaptionLike(c, needle)
            If Not hit Is Nothing Then
                Set FindButtonByCaptionLike = hit
                Exit Function
            End If
        End If
    Next

SafeExit:
End Function


'=== 位置でヘッダ行のボタンを拾う（最上段の右側 2 個を採用）===
Private Sub GatherButtons(container As Object, ByRef arr As Collection)
    Dim c  As Object
    Dim pg As MSForms.Page

    ' MultiPage は Pages を再帰
    If TypeName(container) = "MultiPage" Then
        For Each pg In container.Pages
            GatherButtons pg, arr
        Next
        Exit Sub
    End If

    ' それ以外は Controls
    For Each c In container.Controls
        If TypeName(c) = "CommandButton" Then arr.Add c
        If TypeName(c) = "Frame" Or TypeName(c) = "Page" Then
            GatherButtons c, arr
        End If
    Next
End Sub


Private Sub FindHeaderButtonsByPosition(container As Object, _
        ByRef bSave As MSForms.CommandButton, ByRef bLoad As MSForms.CommandButton)
    Dim arr As New Collection, i As Long
    GatherButtons container, arr
    If arr.Count = 0 Then Exit Sub

    Dim minTop As Single: minTop = 1E+20
    For i = 1 To arr.Count
        If arr(i).Top < minTop Then minTop = arr(i).Top
    Next

    Dim cand As New Collection
    For i = 1 To arr.Count
        If arr(i).Top <= minTop + 10 Then cand.Add arr(i)   '最上段±10px
    Next
    ' キーワードで優先判定
    For i = 1 To cand.Count
        Dim cap As String: cap = Replace(CStr(cand(i).caption), vbCrLf, "")
        If InStr(cap, "保存") > 0 Then Set bSave = cand(i)
        If InStr(cap, "読") > 0 Or InStr(cap, "込") > 0 Then Set bLoad = cand(i)
    Next
    ' まだ空いていたら右側２つを割り当て
    Dim right1 As MSForms.CommandButton, right2 As MSForms.CommandButton
    For i = 1 To cand.Count
        If right1 Is Nothing Or cand(i).Left > right1.Left Then
            Set right2 = right1
            Set right1 = cand(i)
        ElseIf right2 Is Nothing Or cand(i).Left > right2.Left Then
            Set right2 = cand(i)
        End If
    Next
    If bSave Is Nothing Then Set bSave = right1
    If bLoad Is Nothing Then Set bLoad = right2
End Sub



'=== 一覧を出す入口（公開） ===
Public Sub Debug_ListButtons()
    DumpButtonsProc Me
End Sub

'=== ボタン列挙（MultiPage対応版） ===
Private Sub DumpButtonsProc(container As Object)
    Dim c As Control
    Dim pg As MSForms.Page

    If TypeName(container) = "MultiPage" Then
        'MultiPage は Pages 配下を回す
        For Each pg In container.Pages
            For Each c In pg.Controls
                If TypeName(c) = "CommandButton" Then
                    Debug.Print "Type=CommandButton, cap=[" & c.caption & _
                                "], Top=" & c.Top & ", Left=" & c.Left
                End If
                If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
                    DumpButtonsProc c
                End If
            Next
        Next
    Else
        '通常のコンテナ（UserForm / Frame / Page など）
        For Each c In container.Controls
            If TypeName(c) = "CommandButton" Then
                Debug.Print "Type=CommandButton, cap=[" & c.caption & _
                            "], Top=" & c.Top & ", Left=" & c.Left
            End If
            If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
                DumpButtonsProc c
            End If
        Next
    End If
End Sub



' ========= ここから貼る（cmdLoadPrev_Click 全体）=========
Private Sub cmdLoadPrev_Click()
Debug.Print "[ENTER] cmdLoadPrev_Click", Timer


    Dim ws As Worksheet: Set ws = modSchema.GetEvalDataSheet()

    ' 1) 氏名をフォームから取得（ラベル拾い）
    Dim frm As MSForms.Frame
    Dim pname As String
    Set frm = modEvalIOEntry.GetBasicInfoFrame(Me)
    pname = GetTextByLabelInFrame(frm, "氏名")
    If Trim$(pname) = "" Then
        MsgBox "氏名が空です。読み込めません。", vbExclamation
        Exit Sub
    End If

    ' 2) 参照列を準備（存在すれば使う。なければスキップでOK）
    Dim look As Object: Set look = CreateObject("Scripting.Dictionary")
    look.CompareMode = 1
   ' --- 見出しの列位置を look に直書きで入れる（未定義関数は使わない） ---
look.RemoveAll
look.CompareMode = 1  ' vbTextCompare

' 氏名
look("Basic.Name") = modEvalIOEntry.FindColByHeaderExact(ws, "Basic.Name")
If look("Basic.Name") = 0 Then look("Basic.Name") = modEvalIOEntry.FindColByHeaderExact(ws, "氏名")
If look("Basic.Name") = 0 Then look("Basic.Name") = modEvalIOEntry.FindColByHeaderExact(ws, "Name")

' ID（候補表示用に使うだけ）
look("Basic.ID") = modEvalIOEntry.FindColByHeaderExact(ws, "Basic.ID")
If look("Basic.ID") = 0 Then look("Basic.ID") = modEvalIOEntry.FindColByHeaderExact(ws, "BasicInfo_ID")
If look("Basic.ID") = 0 Then look("Basic.ID") = modEvalIOEntry.FindColByHeaderExact(ws, "ID")

' 年齢（候補表示用）
look("Basic.Age") = modEvalIOEntry.FindColByHeaderExact(ws, "Basic.Age")
If look("Basic.Age") = 0 Then look("Basic.Age") = modEvalIOEntry.FindColByHeaderExact(ws, "年齢")
If look("Basic.Age") = 0 Then look("Basic.Age") = modEvalIOEntry.FindColByHeaderExact(ws, "Age")

' 評価日（候補表示用）
look("Basic.EvalDate") = modEvalIOEntry.FindColByHeaderExact(ws, "Basic.EvalDate")
If look("Basic.EvalDate") = 0 Then look("Basic.EvalDate") = modEvalIOEntry.FindColByHeaderExact(ws, "評価日")
If look("Basic.EvalDate") = 0 Then look("Basic.EvalDate") = modEvalIOEntry.FindColByHeaderExact(ws, "EvalDate")


    ' 3) 名前で探索（複数なら候補番号を選ばせる）
  ' ③ 名前で行番号を特定
Dim r As Long
r = FindRowByNameWithPickLocal(ws, pname)


If r <= 0 Then
    MsgBox "前回の値が見つかりません。", vbInformation
    Exit Sub
End If

' ★氏名一致ガード（選択行が入力名と一致しなければ中止）
Dim cName As Long
cName = modEvalIOEntry.FindColByHeaderExact(ws, "氏名")
If cName = 0 Then cName = modEvalIOEntry.FindColByHeaderExact(ws, "利用者名")
If cName = 0 Then cName = modEvalIOEntry.FindColByHeaderExact(ws, "名前")
If cName = 0 Then
    MsgBox "氏名列が見つかりません。", vbExclamation
    Exit Sub
End If
If StrComp(CStr(ws.Cells(r, cName).value), pname, vbTextCompare) <> 0 Then
    MsgBox "選択行の氏名が「" & pname & "」と一致しません。読み込みを中止します。", vbExclamation
    Exit Sub
End If


' 見つかったら黙って読み込み
    Call modEvalIOEntry.LoadEvaluation_ByName_From(Me)



End Sub





Private Sub UserForm_Activate()
   

   
    Static done As Boolean
    Dim scrH As Single
    Dim h As Single
    
    Application.WindowState = xlMaximized
    
    If done Then Exit Sub
    done = True
   
    Me.Controls("txtHdrName").SetFocus


End Sub




'==== UserForm コードモジュール（fraDynPainBox が載っているフォーム）====

Private Sub UserForm_Initialize()

Me.StartUpPosition = 0
Me.Width = 1072
Me.Height = 632.15
Me.Left = Application.Left + (Application.Width - Me.Width) / 2: Me.Top = Application.Top + (Application.Height - Me.Height) / 2



 Dim scrH As Single
    Dim h As Single
    scrH = Application.UsableHeight
    If scrH < 500 Then
        h = 530
    Else
        h = 690
    End If

    Me.Height = h
    DoEvents

    Call LegacyInit
#If APP_DEBUG Then
    Debug.Print "[PostInit] CtlCount=" & Me.Controls.Count
#End If

TidyBaseLayout_Once

Call AddPainQualUI
Call AddPainFactorsUI
Call AddVASUI: Call WireVAS
Call AddPainCourseUI
Call AddPainSiteUI
Call ArrangePainLayout
Call RemoveLegacyPainUI


  If Not mPainLayoutDone Then
        TidyPainBoxes
        mPainLayoutDone = True
    End If
    
    TidyPainBoxes        ' ← 右列(誘因・軽減因子)の恒久配置
    TidyPainCourse       ' ← ★追加：経過・時間の変化の恒久配置
    Me.WidenAndTidyPainCourse
 
   
    

    Me.TidyPainUI_Once

If Not mPainTidyBusy Then
    'TidyPainUI_Once
    Me.Height = h
    DoEvents
    ClearPainUI Me   ' ← 起動時は空で開始（読み込みは手動で）
End If

        BuildWalkUI_All

    Me.Controls("Frame31").caption = ""
    BuildCogMentalUI_Simple
    BuildCog_CognitionCore      '← 認知6項目を生成
    BuildCog_DementiaBlock      '← 認知症の種類＋備考を生成
    BuildCog_BPSD
    BuildCog_MentalBlock
    BuildDailyLogTab
    Me.Controls("txtDailyStaff").IMEMode = fmIMEModeHiragana
    Me.Controls("txtDailyNote").IMEMode = fmIMEModeHiragana
    
    Set mDailyList = New clsDailyLogList
    'Set mDailyList.lb = Me.Controls("lstDailyLogList")


    
    BuildDailyLog_ExtractButton Me
    BuildDailyLog_SaveButton Me
       Me.Controls("txtEDate").value = Date
    Me.Controls("txtDailyDate").value = Date

    If Not mPlacedGlobalSave Then
    PlaceGlobalSaveButton_Once
    mPlacedGlobalSave = True
    End If
    


    
    On Error Resume Next
    Me.Controls("btnSaveCtl").Visible = False
    On Error GoTo 0
    
    'Me.Height = Application.UsableHeight - 40
    
    




If mBaseLayoutDone Then
    'Apply_AlignRoot_All
End If


'--- Fix: 子MultiPage見切れ対策（2025-12-13 OKスナップショット固定）

'Me.Controls("mpPhys").Height = 520#
'Me.Controls("cmdSaveGlobal").Top = Me.InsideHeight - Me.Controls("cmdSaveGlobal").Height - 12

Me.Controls("MultiPage2").parent.Height = Me.Controls("MultiPage2").Height
Me.Controls("Frame12").Height = 508.1
Me.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Height = Me.Controls("mpPhys").Pages(0).Controls("Frame8").InsideHeight - Me.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Top
'Me.Controls("mpPhys").Height = Me.InsideHeight - 80
Me.Controls("mpPhys").Pages(0).Controls("Frame8").Height = Me.Controls("mpPhys").Height
Me.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Height = Me.Controls("mpPhys").Pages(0).Controls("Frame8").InsideHeight - Me.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Top



    'frmEval.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Height = _
    'frmEval.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Pages(0).Controls("Frame19").Top + _
    'frmEval.Controls("mpPhys").Pages(0).Controls("Frame8").Controls("mpROM").Pages(0).Controls("Frame19").Height + 24





    Call BuildEvalShell_Once
    Call CreateHeaderButtons_Once

    Tidy_DailyLog_Once

    
    

    Fix_Page8_DailyLog_Once
    Fix_Page6_Walk_FrameScroll_Once
    ApplyScroll_MP1_Page3_7_Once
    
    Call Preview_NameToHeader
    Me.Controls("txtName").Visible = False
    Me.Controls("txtPID").Visible = False
    Me.Controls("Label115").Visible = False
    Me.Controls("Label112").Visible = False

    
    AddHeaderArchiveDeleteButton

   
   
  

AddPrintButton_TestEval

Call Ensure_MonthlyDraftBox_UnderFraDailyLog

Set mHdrNameSink = New cHdrNameSink
mHdrNameSink.Hook Me.Controls("frHeader").Controls("txtHdrName")


Call Align_LoadPrevButton_NextToHdrKana(Me)
Call Ensure_LoadPrevButton_Once(Me)

'--- hook header "LoadPrev" button (MUST be after Ensure_LoadPrevButton_Once) ---
Set mHdrLoadPrevHook = New clsHdrBtnHook
Set mHdrLoadPrevHook.btn = Me.Controls("frHeader").Controls("cmdHdrLoadPrev")
mHdrLoadPrevHook.tag = "LoadPrev"
Set mHdrLoadPrevHook.owner = Me
DoEvents

On Error Resume Next
Me.Controls("MultiPage1").Pages(0).Controls("Frame32").Controls("btnLoadPrevCtl").Visible = False
On Error GoTo 0
    
End Sub


   
   
   
   Private Sub FillNRS(ByVal cmb As MSForms.ComboBox)
    Dim i As Long
    cmb.Clear
    For i = 0 To 10
        cmb.AddItem CStr(i)
    Next
    cmb.ListIndex = 0
    cmb.Style = fmStyleDropDownList
    cmb.MatchRequired = True
End Sub
'=== frmEval のコードモジュールに貼る（ここまで） ===





Private Sub LegacyInit()



    Set mMPHooks = New Collection
    Set mTxtHooks = New Collection

SetupLayout

    Me.caption = "評価フォーム"
Me.ScrollBars = fmScrollBarsNone

' 画面サイズに合わせて上限をかける（最小変更）
Const MARGIN_W As Single = 80
Const MARGIN_H As Single = 140

Dim initW As Single, initH As Single
initW = Application.UsableWidth - 80
initH = Application.UsableHeight - 140
If initW < FWIDTH + 60 Then initW = FWIDTH + 60
If initW > 1200 Then initW = 1200
If initH < 640 Then initH = 640
If initH > 820 Then initH = 820
Me.Width = initW
'Me.Height = initH
Me.StartUpPosition = 1      ' CenterOwner（任意：中央に出したい場合）



    ' ルート MultiPage
    Set mp = Me.Controls.Add("Forms.MultiPage.1")
    With mp
        .Left = 6: .Top = 6
        .Width = Me.InsideWidth - 12
        .Height = Me.InsideHeight - 60
        .Style = fmTabStyleTabs
    End With
    ' 安全策：最低2ページを保証（環境差対策）
    If mp.Pages.Count = 0 Then mp.Pages.Add: mp.Pages.Add

    ' 先頭2枚は既存
mp.Pages(0).caption = "基本情報"
mp.Pages(1).caption = "姿勢評価"

' ←ここで「身体機能評価」を先に追加
Dim pgPhys  As MSForms.Page: Set pgPhys = mp.Pages.Add: pgPhys.caption = "身体機能評価"

' 残りの親タブ
Dim pgMove  As MSForms.Page: Set pgMove = mp.Pages.Add: pgMove.caption = "日常生活動作"
Dim pgTests As MSForms.Page: Set pgTests = mp.Pages.Add: pgTests.caption = "テスト・評価"
Dim pgWalk  As MSForms.Page: Set pgWalk = mp.Pages.Add: pgWalk.caption = "歩行評価"
Dim pgCog   As MSForms.Page: Set pgCog = mp.Pages.Add: pgCog.caption = "認知・精神"

' --- ホスト変数の宣言（ここを追加）---
' --- ホスト変数の宣言 ---
Dim hostBasic As MSForms.Frame
Dim hostPost  As MSForms.Frame   ' ← 旧 hostBody を使っていたら置換
Dim hostPhys  As MSForms.Frame
Dim hostMove  As MSForms.Frame
Dim hostTests As MSForms.Frame
Dim hostWalk  As MSForms.Frame
Dim hostCog   As MSForms.Frame


' --- ここから既存のホスト生成 ---
Set hostBasic = CreateScrollHost(mp.Pages(0))
Set hostPost = CreateScrollHost(mp.Pages(1))
Set hostPhys = CreateScrollHost(pgPhys)
Set hostMove = CreateScrollHost(pgMove)
Set hostTests = CreateScrollHost(pgTests)
Set hostWalk = CreateScrollHost(pgWalk)
Set hostCog = CreateScrollHost(pgCog)



    
 Call modPhysEval.EnsurePhysicalFunctionTabs_Under(Me, mp)
    
    
    
    
' --- 迷子の mpADL を全消去（保険） ---
Trace "RemoveAllMpADL()", "Initialize"
RemoveAllMpADL
Trace "RemoveAllMpADL() done", "Initialize"


' --- mpADL の器を用意（ページ0と1を作る） ---
Trace "EnsureBI_IADL() call", "Initialize"
Dim mpADL As MSForms.MultiPage
Set mpADL = EnsureBI_IADL()
Trace "EnsureBI_IADL() returned; pages=" & mpADL.Pages.Count, "Initialize"


' --- 起居動作ページのUIを作る ---
Trace "BuildKyoOnADL Page(2) start", "Initialize"
BuildKyoOnADL mpADL.Pages(2)
Trace "BuildKyoOnADL Page(2) end", "Initialize"



' 初期表示はBIでOKなら 0 のまま
mpADL.value = 0


' （任意）IADL備考のIME再設定
ApplyImeToIADLNote

' --- 下に続く nextTop の更新など ---
Trace "update nextTop start", "Initialize"

nextTop = mpADL.Top + mpADL.Height + 10
hostMove.ScrollHeight = nextTop + 10

Trace "update nextTop done; mpADL.Top=" & mpADL.Top & _
      ", mpADL.Height=" & mpADL.Height & _
      ", nextTop=" & nextTop, "Initialize"


Dim Y As Single, cboTmp As MSForms.ComboBox



    '================ テスト・評価 ================
    Trace "TESTS start", "Init"
    
    nextTop = pad
    Dim fTests As MSForms.Frame: Set fTests = CreateFrameP(hostTests, "テスト・評価（秒=小数2桁・握力=小数1桁・0以上）", 150)
    Y = 22
    CreateLabel fTests, "10m歩行（秒）", COL_LX, Y: CreateTextBox fTests, COL_LX + lblW + 20, Y, 80, 0, False, "txtTenMWalk", "Test.10m": nL Y
    CreateLabel fTests, "TUG（秒）", COL_LX, Y: CreateTextBox fTests, COL_LX + lblW + 20, Y, 80, 0, False, "txtTUG", "Test.TUG"
    CreateLabel fTests, "5回立ち上がり（秒）", COL_RX, 22: CreateTextBox fTests, COL_RX + lblW + 60, 22, 80, 0, False, "txtFiveSTS", "Test.5xSTS"
    CreateLabel fTests, "セミタンデム（秒）", COL_RX, 46: CreateTextBox fTests, COL_RX + lblW + 60, 46, 80, 0, False, "txtSemi", "Test.Semi"
    CreateLabel fTests, "握力 右（kg）", COL_LX, 94: CreateTextBox fTests, COL_LX + lblW + 20, 94, 80, 0, False, "txtGripR", "Grip.R"
    CreateLabel fTests, "握力 左（kg）", COL_RX, 94: CreateTextBox fTests, COL_RX + lblW + 40, 94, 80, 0, False, "txtGripL", "Grip.L"
    ResizeFrameToContent fTests, 120

    Trace "TESTS end", "Init"

    
    '================ 歩行評価（自立度 / RLA） ================
Trace "WALK start", "Init"

Set mpWalk = hostWalk.Controls.Add("Forms.MultiPage.1")

' 変数宣言は With の外でOK
Dim w As Single, h As Single

With mpWalk
    .Left = 0
    .Top = 0

    ' 内寸ベースで算出し、下限を付けてクランプ
    w = hostWalk.InsideWidth - 12:  If w < 200 Then w = 200
    h = hostWalk.InsideHeight - 12: If h < 150 Then h = 150

    .Width = w
    .Height = h
    .Style = fmTabStyleTabs
End With

' ページ(0),(1)を触る前に2枚保証
Do While mpWalk.Pages.Count < 2
    mpWalk.Pages.Add
Loop
mpWalk.Pages(0).caption = "自立度"
mpWalk.Pages(1).caption = "RLA"


    Dim hostWalkGait As MSForms.Frame
    Set hostWalkGait = mpWalk.Pages(0).Controls.Add("Forms.Frame.1")
    With hostWalkGait
    .caption = ""
    .Left = 0: .Top = 0
    .Width = mpWalk.Width - 12
    Dim tGait As Single: tGait = mpWalk.Height - 30
    If tGait < 120 Then tGait = 120     ' ← ここを t → tGait に
    .Height = tGait                      ' ← ここも t → tGait に
    .ScrollBars = fmScrollBarsNone
End With


    nextTop = pad
    Dim fGait As MSForms.Frame: Set fGait = CreateFrameP(hostWalkGait, "歩行評価（自立度）", 90)
    Y = 22
    Dim rowH As Single: rowH = 28
    
    CreateLabel fGait, "歩行自立度", COL_LX, Y
    Dim cboGait As MSForms.ComboBox: Set cboGait = CreateCombo(fGait, COL_LX + lblW, Y, 500, , "Gait.自立度")
    cboGait.List = MakeList("完全自立,修正自立（補助具使用）,監視・見守り,軽介助（25%未満）,中等度介助（25-50%）,重介助（50%以上）,全介助")
    ResizeFrameToContent fGait, Y + rowH

    Dim hostWalkRLA As MSForms.Frame
    Set hostWalkRLA = mpWalk.Pages(1).Controls.Add("Forms.Frame.1")
    With hostWalkRLA: .caption = "": .Left = 0: .Top = 0: .Width = mpWalk.Width - 12: .Height = mpWalk.Height - 30: .ScrollBars = fmScrollBarsNone: End With

    Dim mpRLA As MSForms.MultiPage
    Set mpRLA = hostWalkRLA.Controls.Add("Forms.MultiPage.1")
    With mpRLA
        .Left = 0: .Top = 0
        .Width = hostWalkRLA.Width - 6
        .Height = hostWalkRLA.Height - 6
        .Style = fmTabStyleTabs
    End With
    mpRLA.Pages(0).caption = "立脚期（IC-TSt）"
    mpRLA.Pages(1).caption = "遊脚期（PSw-TSw）"

    Dim hostRLAStance As MSForms.Frame
    Set hostRLAStance = mpRLA.Pages(0).Controls.Add("Forms.Frame.1")
With hostRLAStance
    .caption = ""
    .Left = 0: .Top = 0
    .Width = mpRLA.Width - 12
    Dim tStance As Single: tStance = mpRLA.Height - 30
    If tStance < 120 Then tStance = 120
    .Height = tStance
    .ScrollBars = fmScrollBarsNone
End With

    nextTop = pad
    Dim fRLA1 As MSForms.Frame: Set fRLA1 = CreateFrameP(hostRLAStance, "RLA歩行周期（IC / LR / MSt / TSt）", 280)
    Build_RLA_ChecksPart fRLA1, "stance": ResizeFrameToContent fRLA1, 260

    Dim hostRLASwing As MSForms.Frame
    Set hostRLASwing = mpRLA.Pages(1).Controls.Add("Forms.Frame.1")
With hostRLASwing
    .caption = ""
    .Left = 0: .Top = 0
    .Width = mpRLA.Width - 12
    Dim tSwing As Single: tSwing = mpRLA.Height - 30
    If tSwing < 120 Then tSwing = 120
    .Height = tSwing
    .ScrollBars = fmScrollBarsNone
End With


    nextTop = pad
    Dim fRLA2 As MSForms.Frame: Set fRLA2 = CreateFrameP(hostRLASwing, "RLA歩行周期（PSw / ISw / MSw / TSw）", 280)
    Build_RLA_ChecksPart fRLA2, "swing": ResizeFrameToContent fRLA2, 260

    
    
    Trace "WALK end", "Init"



    '================ 認知・精神 ================
   Trace "COG start", "Init"
    
    nextTop = pad
    Dim fCog As MSForms.Frame: Set fCog = CreateFrameP(hostCog, "認知機能・精神面", 110)
    Y = 22
    CreateLabel fCog, "認知機能レベル", COL_LX, Y
    Dim cboCog As MSForms.ComboBox: Set cboCog = CreateCombo(fCog, COL_LX + lblW, Y, 160, , "Cognition.Level")
    cboCog.List = MakeList("正常,軽度低下,中等度低下,高度低下")
    CreateLabel fCog, "精神面", COL_RX, Y
    Dim cboPsy As MSForms.ComboBox: Set cboPsy = CreateCombo(fCog, COL_RX + lblW, Y, 160, , "Psych.Status")
    cboPsy.List = MakeList("安定,不安傾向,抑うつ傾向,その他")
    CreateLabel fCog, "備考", COL_LX, Y + 28
    CreateTextBox fCog, COL_LX + lblW, Y + 26, 610, 50, True, , "Cognition.備考"
    ResizeFrameToContent fCog, Y + 26 + 50
    
    Trace "COG end", "Init"




    '--- 下部の「閉じる」 ---
    Trace "CLOSE start", "Init"

Set btnCloseCtl = Me.Controls.Add("Forms.CommandButton.1")
With btnCloseCtl
    .caption = "閉じる"
    .Width = 60
    .Height = 26
    .Left = Me.InsideWidth - .Width - 20
    .Top = nextTop + 20
    .name = "btnCloseCtl"
End With
Me.ScrollBars = fmScrollBarsVertical
Me.ScrollHeight = nextTop + 120

Trace "CLOSE end", "Init"



RecalcBI



    
    

    '================ 基本情報：全面レイアウト ================
    nextTop = pad
  Dim fBasic As MSForms.Frame: Set fBasic = CreateFrameP(hostBasic, "基本情報", 360)
    Set fBasicRef = fBasic
    Y = 22

    ' --- 最上段：ユーティリティ行 ---
    Dim chkDelta As MSForms.CheckBox
    Set chkDelta = CreateCheck(fBasic, "変更点のみ保存（空欄は前回値を引継ぎ）", COL_LX, 6, "chkDeltaOnly", "Delta.Only")
    chkDelta.AutoSize = True: chkDelta.WordWrap = False

    Set btnLoadPrevCtl = fBasic.Controls.Add("Forms.CommandButton.1")
    With btnLoadPrevCtl
        .caption = "前回の値を読み込む": .Accelerator = "L"
        .Width = 180: .Height = 24: .name = "btnLoadPrevCtl"
    End With
    Set btnSaveCtl = fBasic.Controls.Add("Forms.CommandButton.1")
    With btnSaveCtl
        .caption = "シートへ保存": .Accelerator = "S"
        .Width = 120: .Height = 24: .name = "btnSaveCtl"
    End With
    PositionTopRightButtons fBasic
    nL Y, 1

    ' 行1：ID / 評価日 / 評価者（小さめ）
    CreateLabel fBasic, "ID", COL_LX, Y
    Dim tbPID As MSForms.TextBox
    Set tbPID = CreateTextBox(fBasic, COL_LX + lblW, Y, 120, 0, False, "txtPID", "PatientID")
    btnLoadPrevCtl.Left = tbPID.Left + tbPID.Width + 12
    btnLoadPrevCtl.Top = tbPID.Top
    CreateLabel fBasic, "評価日", COL_RX, Y
    Dim tbED As MSForms.TextBox: Set tbED = CreateTextBox(fBasic, COL_RX + lblW, Y, 120, 0, False, "txtEDate", "EvalDate")
    tbED.Text = Format(Date, "yyyy/mm/dd")
    CreateLabel fBasic, "評価者", COL_RX + lblW + 130, Y
    Dim tbEva As MSForms.TextBox: Set tbEva = CreateTextBox(fBasic, COL_RX + lblW + 180, Y, 90, 0, False, "txtEvaluator", "Basic.Evaluator")
    tbEva.Font.Size = 8
    nL Y

    ' 行2：氏名 / 年齢 / 性別
    CreateLabel fBasic, "氏名", COL_LX, Y
    CreateTextBox fBasic, COL_LX + lblW, Y, 200, 0, False, "txtName", "Basic.Name"
    CreateLabel fBasic, "年齢", COL_RX, Y
    CreateTextBox fBasic, COL_RX + lblW, Y, 60, 0, False, "txtAge", "Basic.Age"
    CreateLabel fBasic, "性別", COL_RX + lblW + 70, Y
    Dim cboSex As MSForms.ComboBox: Set cboSex = CreateCombo(fBasic, COL_RX + lblW + 110, Y, 90, "cboSex", "Basic.Gender")
    cboSex.List = MakeList("男性,女性,その他,不明")
    nL Y
    
    ' 行2.5：生年月日
    CreateLabel fBasic, "生年月日", COL_RX, Y
    Dim tbBirth As MSForms.TextBox
    Set tbBirth = CreateTextBox(fBasic, COL_RX + lblW, Y, 120, 0, False, "txtBirth", "Basic.BirthDate")
    tbBirth.IMEMode = fmIMEModeOff


    nL Y
    
    CreateLabel fBasic, "※生年月日は 19990804 の形式で入力してください（年齢は自動計算されます）", COL_RX + lblW + 130, Y
    fBasic.Controls(fBasic.Controls.Count - 1).Font.Size = 8
    fBasic.Controls(fBasic.Controls.Count - 1).Top = Y - 21



    ' 行3：主診断 / 発症日
    CreateLabel fBasic, "主診断", COL_LX, Y
    CreateTextBox fBasic, COL_LX + lblW, Y, 260, 0, False, "txtDx", "Basic.PrimaryDx"
    CreateLabel fBasic, "発症日", COL_RX, Y
    CreateTextBox fBasic, COL_RX + lblW, Y, 120, 0, False, "txtOnset", "Basic.OnsetDate"
    nL Y

    ' 行4：生活状況 / 要介護度
    CreateLabel fBasic, "生活状況", COL_LX, Y
    CreateTextBox fBasic, COL_LX + lblW, Y, 220, 0, False, "txtLiving", "Basic.Living"
    CreateLabel fBasic, "要介護度", COL_RX, Y
    Dim cboLev As MSForms.ComboBox: Set cboLev = CreateCombo(fBasic, COL_RX + lblW, Y, 150, "cboCare", "Basic.CareLevel")
    cboLev.List = MakeList("要支援1,要支援2,要介護1,要介護2,要介護3,要介護4,要介護5")
    nL Y

    ' 行5：障害高齢者／認知症高齢者（ラベル個別幅）
    Dim LBLW_LONG_LEFT As Long:  LBLW_LONG_LEFT = 150
    Dim LBLW_LONG_RIGHT As Long: LBLW_LONG_RIGHT = 170

    CreateLabel fBasic, "障害高齢者の日常生活自立度", COL_LX, Y, LBLW_LONG_LEFT
    Dim cboEL As MSForms.ComboBox
    Set cboEL = CreateCombo(fBasic, COL_LX + LBLW_LONG_LEFT + 6, Y, 180, "cboElder", "Basic.ElderlyLevel")
    cboEL.List = MakeList("自立,J1,J2,A1,A2,B1,B2,C1,C2")

    CreateLabel fBasic, "認知症高齢者の日常生活自立度", COL_RX, Y, LBLW_LONG_RIGHT
    Dim cboDL As MSForms.ComboBox
    Set cboDL = CreateCombo(fBasic, COL_RX + LBLW_LONG_RIGHT + 6, Y, 160, "cboDementia", "Basic.DementiaLevel")
    cboDL.List = MakeList("自立,I,IIa,IIb,IIIa,IIIb,IV,M")
    nL Y, 1

    ' 補助具／リスク（2列）
    nL Y, 1

    Dim frRisk As MSForms.Frame
    Dim ASSISTIVE_CSV As String: ASSISTIVE_CSV = "杖,シルバーカー,歩行器,車いす,短下肢装具,介助ベルト,手すり,スロープ"
    Dim RISK_CSV As String: RISK_CSV = "転倒,誤嚥,失禁,褥瘡,低栄養,徘徊,せん妄,ADL低下"

    Set frRisk = BuildCheckFrame(fBasic, "リスク", COL_RX, Y, 370, MakeList(RISK_CSV), "RiskGroup")

    Dim nextY As Single
    nextY = frRisk.Top + frRisk.Height
    Y = nextY + 8

    BuildAssistiveChecksInWalkEval ASSISTIVE_CSV

    ' Needs（左右2カラム）
    Dim needsH As Single: needsH = 36
    CreateLabel fBasic, "患者Needs", COL_LX, Y
    CreateTextBox fBasic, COL_LX + lblW, Y, 270, needsH, True, "txtNeedsPt", "Needs.Patient"
    CreateLabel fBasic, "家族Needs", COL_RX, Y
    CreateTextBox fBasic, COL_RX + lblW, Y, 270, needsH, True, "txtNeedsFam", "Needs.Family"
    Y = Y + needsH + 10

    ResizeFrameToContent fBasic, Y
    
    
    ' レイアウト＆タブ順
    FitLayout
    If Not fBasicRef Is Nothing Then PositionTopRightButtons fBasicRef
    ResetTabOrder

    ' 初期の保存ボタン活性状態を反映
    RefreshSaveEnabled
    Me.Controls("btnSaveCtl").Enabled = True  ' ← 一旦、保存ボタンを常時有効化
    
    

    '================ 姿勢評価（変形要因は備考へ） ================
    nextTop = pad
    Dim fPost As MSForms.Frame
    Set fPost = CreateFrameP(hostPost, "姿勢評価（変形要因は備考へ）", 200)

    Dim yP As Single: yP = 22

    ' 1行目
    CreateCheck fPost, "頭部前方突出", COL_LX, yP, , "Posture.頭部前方突出"
    CreateCheck fPost, "円背", COL_LX + 150, yP, , "Posture.円背"
    CreateCheck fPost, "側弯", COL_LX + 300, yP, , "Posture.側弯"
    CreateLabel fPost, "骨盤傾斜", COL_RX, yP - 2
    Dim cboPel As MSForms.ComboBox
    Set cboPel = CreateCombo(fPost, COL_RX + lblW, yP - 2, 120, , "Posture.骨盤傾斜")
    cboPel.List = MakeList("前傾,後傾,正常,不明")
    nL yP

    ' 2行目
    CreateCheck fPost, "体幹回旋", COL_LX, yP, , "Posture.体幹回旋"
    CreateCheck fPost, "反張膝", COL_LX + 150, yP, , "Posture.反張膝"

    ' 備考（右上）
    CreateLabel fPost, "備考", COL_RX, yP - 2
    CreateTextBox fPost, COL_RX + lblW + 10, yP - 4, 190, 50, True, "", "Posture.備考"

    nL yP, 3
    ResizeFrameToContent fPost, yP

    nextTop = fPost.Top + fPost.Height + 6

    '================ 関節拘縮（右・左）＋備考 ================
    Dim fCon As MSForms.Frame: Set fCon = CreateFrameP(hostPost, "関節拘縮（右・左）＋備考", 180)
    Dim y0 As Single: y0 = 22
    Y = y0

    ' ガイド見出し
    CreateLabel fCon, "部位", COL_LX, Y
    CreateLabel fCon, "右", COL_LX + 90 + 20, Y
    CreateLabel fCon, "左", COL_LX + 90 + 20 + 60, Y
    nL Y

    CreateCheck fCon, "頸部（左右なし）", COL_LX, Y, "", "Contracture.頸部": nL Y

    ' 部位ごとにR/Lチェック
    CreateContractureRLRow fCon, Y, "肩関節", "Contracture.肩"
    CreateContractureRLRow fCon, Y, "肘関節", "Contracture.肘"
    CreateContractureRLRow fCon, Y, "手関節", "Contracture.手関節"
    CreateContractureRLRow fCon, Y, "股関節", "Contracture.股関節"
    CreateContractureRLRow fCon, Y, "膝関節", "Contracture.膝関節"
    CreateContractureRLRow fCon, Y, "足関節", "Contracture.足関節"

    ' 備考（右側上部に）
    CreateLabel fCon, "備考", COL_RX, y0 - 2
    CreateTextBox fCon, COL_RX + lblW, y0 - 4, 250, 80, True, "", "Contracture.備考"

    ' 高さ調整
    ResizeFrameToContent fCon, Application.WorksheetFunction.Max(Y, y0 + 80)


  SetupInputModesJP
  
'== ROM内の空ページ(Page14/15)を起動時に自動削除 ==
Dim ctlZ As Object, mpZ As MSForms.MultiPage, iZ As Long, capZ As String
For Each ctlZ In Me.Controls
    If TypeName(ctlZ) = "MultiPage" Then
        Set mpZ = ctlZ
        For iZ = mpZ.Pages.Count - 1 To 0 Step -1
            capZ = CStr(mpZ.Pages(iZ).caption)
            If capZ = "Page14" Or capZ = "Page15" Then mpZ.Pages.Remove iZ
        Next iZ
    End If
Next ctlZ


'== 備考欄（ラベル＋大きいTextBox）を非表示 ==

Dim mpN As Object, pgN As Object, pN As Object, cN As Object

' --- ROMページ特定（"ROM" または "主要関節"） ---
Set mpN = Nothing: Set pgN = Nothing
For Each cN In Me.Controls
    If TypeName(cN) = "MultiPage" Then
        Set mpN = cN
        Exit For
    End If
Next cN

If Not mpN Is Nothing Then
    For Each pN In mpN.Pages
    If InStr(1, CStr(pN.caption), "ROM", vbTextCompare) > 0 _
       Or InStr(1, CStr(pN.caption), "主要関節", vbTextCompare) > 0 Then
        Set pgN = pN: Exit For
    End If
Next pN

End If



If Not pgN Is Nothing Then
    Dim stk As Collection: Set stk = New Collection
    Dim parent As Object, ctl As Object
    Dim nLbl As Object, nTB As Object, isML As Boolean

    ' 子コンテナを含めて深さ優先で探索
    stk.Add pgN
    Do While stk.Count > 0
        Set parent = stk(1): stk.Remove 1
        On Error Resume Next
        For Each ctl In parent.Controls
            On Error GoTo 0
            Select Case TypeName(ctl)
                Case "Frame", "MultiPage", "Page"
                    stk.Add ctl                      ' 子をたどる
                Case "Label"
                    If InStr(1, CStr(ctl.caption), "備考", vbTextCompare) > 0 Then
                        Set nLbl = ctl              ' 「備考/備考欄」を捕捉
                    End If
                Case "TextBox"
                    isML = False
                    On Error Resume Next: isML = ctl.multiline: On Error GoTo 0
                    If isML Or ctl.Height >= 80 Or ctl.Width >= 400 Then
                        If nTB Is Nothing Then
                            Set nTB = ctl
                        ElseIf ctl.Height * ctl.Width > nTB.Height * nTB.Width Then
                            Set nTB = ctl          ' 最大サイズを備考として採用
                        End If
                    End If
            End Select
        Next ctl
    Loop

    ' 見つかったら非表示
    If Not nLbl Is Nothing Then nLbl.Visible = False
    If Not nTB Is Nothing Then nTB.Visible = False
End If


'=== 備考ラベルと備考テキストを再帰的に非表示（全コンテナ対応） ===
Dim qH As New Collection, parentH As Object, ctlH As Object, iH As Long
Dim noteTB As Object, areaMax As Double




' ROMページをその場で特定してから開始
Dim rootPg As Object, c0 As Object, mp0 As Object, i0 As Long
Set rootPg = Nothing
For Each c0 In Me.Controls
    If TypeName(c0) = "MultiPage" Then
        Set mp0 = c0
        For i0 = 0 To mp0.Pages.Count - 1
            If InStr(1, CStr(mp0.Pages(i0).caption), "ROM", vbTextCompare) > 0 _
               Or InStr(1, CStr(mp0.Pages(i0).caption), "主要関節", vbTextCompare) > 0 Then
                Set rootPg = mp0.Pages(i0): Exit For
            End If
        Next i0
        If Not rootPg Is Nothing Then Exit For
    End If
Next c0
If rootPg Is Nothing Then Exit Sub
qH.Add rootPg

Do While qH.Count > 0
    Set parentH = qH(1): qH.Remove 1

    If TypeName(parentH) = "MultiPage" Then
        ' MultiPage は Controls を持たないので Pages を個別に辿る
        For iH = 0 To parentH.Pages.Count - 1
            qH.Add parentH.Pages(iH)
        Next iH
    Else
        On Error Resume Next
        For Each ctlH In parentH.Controls
            On Error GoTo 0
           Select Case TypeName(ctlH)
    Case "Frame", "Page"
        qH.Add ctlH                      ' 子コンテナを辿る
    Case "MultiPage"
        For iH = 0 To ctlH.Pages.Count - 1
            qH.Add ctlH.Pages(iH)        ' ページを個別に辿る
        Next iH
    Case "Label"
        If InStr(1, CStr(ctlH.caption), "備考", vbTextCompare) > 0 Then
            Set nLbl = ctlH
            Set noteTB = ctlH
            ctlH.Visible = False         ' 備考ラベルを消す
        End If
    Case "TextBox"
        Dim mlH As Boolean: mlH = False
        On Error Resume Next: mlH = ctlH.multiline: On Error GoTo 0
        If mlH Or ctlH.Height >= 80 Or ctlH.Width >= 400 Then
            ctlH.Visible = False         ' 大きいテキストBOXを消す
        End If
End Select

        Next ctlH
    End If
Loop

If Not noteTB Is Nothing Then noteTB.Visible = False
'Debug.Print "[ROM_Note] hidden lbl=" & (Not (nLbl Is Nothing)) & "  tb=" & (Not (noteTB Is Nothing))

'=== /備考非表示 ===

Call ROM_Fix_TextBoxHeight_Recursive_OnROM_Once
Call ROM_CheckBoxes_Up12_OnROM_Recursive_Once_V2
Call MMT_BuildChildTabs_Direct


Dim c As Control
For Each c In Me.Controls
    If TypeName(c) = "Label" Then
        If c.caption = "NRS" Then
            c.caption = "安静時NRS"
            Exit For
        End If
    End If
Next

'--- 動作時NRSを自動追加（1回だけ） ---
Dim srcLbl As MSForms.label, srcCmb As MSForms.ComboBox
Dim lbl As MSForms.label, cmb As MSForms.ComboBox
Dim ct As Control

' 安静時NRSラベルを特定
For Each ct In Me.Controls
    If TypeName(ct) = "Label" Then
        If ct.caption = "安静時NRS" Then
            Set srcLbl = ct
            Exit For
        End If
    End If
Next

If Not srcLbl Is Nothing Then
    ' 安静時NRSの右側にある既存Comboを推定（同じ高さ±6）
    For Each ct In Me.Controls
        If TypeName(ct) = "ComboBox" Then
            If Abs(ct.Top - srcLbl.Top) <= 20 And ct.Left > srcLbl.Left Then
                Set srcCmb = ct: Exit For
                srcLbl.Left = srcLbl.parent.InsideWidth - (srcLbl.Width + srcCmb.Width + 12)
                srcCmb.Left = srcLbl.Left + srcLbl.Width + 8
            End If
        End If
    Next

    ' 既に作成済みなら何もしない
    For Each ct In Me.Controls
        If TypeName(ct) = "Label" And ct.name = "lblNRS_Move" Then Set lbl = ct
        If TypeName(ct) = "ComboBox" And ct.name = "cmbNRS_Move" Then Set cmb = ct
    Next

    If lbl Is Nothing Then
       Set lbl = srcLbl.parent.Controls.Add("Forms.Label.1", "lblNRS_Move", True)
        lbl.caption = "動作時NRS"
    End If

    If cmb Is Nothing Then
       Set cmb = srcLbl.parent.Controls.Add("Forms.ComboBox.1", "cmbNRS_Move", True)
        cmb.Style = fmStyleDropDownList
        Dim i As Long
        For i = 0 To 10: cmb.AddItem CStr(i): Next i
    End If

    ' 位置決め（安静時NRSの「下」に配置）
Dim baseLeft As Single, baseTop As Single, gap As Single
gap = 8
baseLeft = 12
If Not srcCmb Is Nothing Then
    baseTop = srcCmb.Top + srcCmb.Height + gap
Else
    baseTop = srcLbl.Top + srcLbl.Height + gap
End If
lbl.Left = baseLeft:  lbl.Top = baseTop
cmb.Left = lbl.Left + lbl.Width + 12
cmb.Top = lbl.Top - 2
cmb.Width = 42

End If
  
  srcLbl.Left = srcLbl.parent.InsideWidth - (srcLbl.Width + srcCmb.Width + 12)
srcCmb.Left = srcLbl.Left + srcLbl.Width + 8
End Sub


Sub ShowFrame12()
    Dim f As Control, t As Control
    For Each f In frmEval.Controls
        If TypeName(f) = "Frame" Then
            For Each t In f.Controls
                If TypeName(t) = "TextBox" And t.name = "TextBox2" Then
                    f.ZOrder 0                 '一番手前に
                    f.caption = "★これがFrame12★" '見つけやすくする
                    Beep
                    Exit Sub
                End If
            Next
        End If
    Next
    MsgBox "TextBox2 の親フレームが見つかりません。"
End Sub













Public Sub AddPainQualUI()
    Dim host As MSForms.Frame
    Dim cap As MSForms.label
    Dim lb  As MSForms.ListBox
    Dim items As Variant
    Dim i As Long

    ' 既存チェック（重複生成防止）
    Dim t As Object
    Set t = FindCtlDeep(Me, "lstPainQual")
    If Not t Is Nothing Then Exit Sub


    ' 既存NRSの親フレームに追加する
    Set host = Me.Controls("cmbNRS_Move").parent

    ' 見出し
    Set cap = host.Controls.Add("Forms.Label.1", "lblPainQual", True)
    cap.caption = "痛みの性質（複数選択可）"
    cap.Left = 12
    cap.Top = 12
    cap.AutoSize = True

    ' リストボックス（複数選択）
    Set lb = host.Controls.Add("Forms.ListBox.1", "lstPainQual", True)
    lb.Left = 12
    lb.Top = cap.Top + cap.Height + 6
    lb.Width = 240
    lb.Height = 96
    lb.MultiSelect = fmMultiSelectMulti
    lb.IntegralHeight = False

    ' 選択肢
    items = Array("鈍痛", "刺すような痛み", "しびれ", "灼熱感", "ズキズキ", _
                  "締め付け感", "圧迫感", "引きつり", "電撃痛", "こわばり", "けいれん", "その他")
    For i = LBound(items) To UBound(items)
        lb.AddItem items(i)
    Next i
End Sub
















Public Sub AddPainFactorsUI()
    Dim host As MSForms.Frame
    Dim fr As MSForms.Frame
    Dim cap As MSForms.label
    Dim i As Long, Y As Single

    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainFactors")
    If Not t Is Nothing Then Exit Sub

    ' 追加先は疼痛タブのフレーム（NRSの親）
    Set host = Me.Controls("cmbNRS_Move").parent

    ' 見出し
    Set cap = host.Controls.Add("Forms.Label.1", "lblPainFactors", True)
    cap.caption = "誘因・軽減因子"
    cap.Left = 270
    cap.Top = 280
    cap.AutoSize = True

    ' コンテナフレーム
    Set fr = host.Controls.Add("Forms.Frame.1", "fraPainFactors", True)
    fr.Left = cap.Left
    fr.Top = cap.Top + cap.Height + 4
    fr.Width = 360
    fr.Height = 120

    ' 左列：誘因（Provoking）
    Dim provItems As Variant, relItems As Variant
    provItems = Array( _
        Array("chkPainProv_Move", "動作"), _
        Array("chkPainProv_Posture", "姿勢"), _
        Array("chkPainProv_Walk", "歩行"), _
        Array("chkPainProv_Lift", "持ち上げ"), _
        Array("chkPainProv_Cough", "咳/くしゃみ") _
    )

    ' 右列：軽減（Relieving）
    relItems = Array( _
        Array("chkPainRelief_Rest", "安静"), _
        Array("chkPainRelief_Heat", "温熱"), _
        Array("chkPainRelief_Cold", "冷却"), _
        Array("chkPainRelief_Med", "服薬"), _
        Array("chkPainRelief_Brace", "コルセット") _
    )

    ' 左列配置
    Y = 8
    For i = LBound(provItems) To UBound(provItems)
        Dim cB As MSForms.CheckBox
        Set cB = fr.Controls.Add("Forms.CheckBox.1", CStr(provItems(i)(0)), True)
        cB.caption = CStr(provItems(i)(1))
        cB.Left = 12
        cB.Top = Y
        Y = Y + cB.Height + 2
    Next i

    ' 右列配置
    Y = 8
    For i = LBound(relItems) To UBound(relItems)
        Dim cb2 As MSForms.CheckBox
        Set cb2 = fr.Controls.Add("Forms.CheckBox.1", CStr(relItems(i)(0)), True)
        cb2.caption = CStr(relItems(i)(1))
        cb2.Left = fr.Width \ 2 + 8
        cb2.Top = Y
        Y = Y + cb2.Height + 2
    Next i
End Sub

Public Sub AddVASUI()
    Dim host As MSForms.Frame
    Dim cap As MSForms.label
    Dim fr As MSForms.Frame
    Dim tb As MSForms.TextBox
    Dim sb As MSForms.ScrollBar

    ' 既存チェック（重複生成防止）
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraVAS")
    If Not t Is Nothing Then Exit Sub
    
    ' 追加先＝疼痛タブのフレーム（NRSの親）
    Set host = Me.Controls("cmbNRS_Move").parent

    ' 見出し
    Set cap = host.Controls.Add("Forms.Label.1", "lblVAS", True)
    cap.caption = "VAS（0?100）"
    cap.Left = 640
    cap.Top = 280
    cap.AutoSize = True

    ' コンテナ
    Set fr = host.Controls.Add("Forms.Frame.1", "fraVAS", True)
    fr.Left = cap.Left
    fr.Top = cap.Top + cap.Height + 4
    fr.Width = 122
    fr.Height = 64

    ' テキストボックス（数値 0?100）
    Set tb = fr.Controls.Add("Forms.TextBox.1", "txtVAS", True)
    tb.Left = 8
    tb.Top = 10
    tb.Width = 40
    tb.Text = "0"

    ' スクロールバー（横）0?100
    Set sb = fr.Controls.Add("Forms.ScrollBar.1", "sldVAS", True)
    sb.Left = tb.Left + tb.Width + 6
    sb.Top = tb.Top + 2
    sb.Width = 60
    sb.Height = 14
    sb.Min = 0
    sb.Max = 100
    sb.value = 0
    sb.Orientation = fmOrientationHorizontal
End Sub



Private Sub mVAS_Change()
    On Error Resume Next
    Dim v As Long
    v = mVAS.value
    Me.Controls("fraVAS").Controls("txtVAS").Text = CStr(v)

End Sub


Public Sub WireVAS()
    On Error Resume Next
    Set mVAS = Me.Controls("Frame12").Controls("fraVAS").Controls("sldVAS")


End Sub





Public Sub AddPainCourseUI()
    Dim host As MSForms.Frame
    Dim lb As MSForms.label
    Dim cB As MSForms.ComboBox
    Dim tb As MSForms.TextBox
    Dim i As Long

    ' 既存チェック（重複生成防止）
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainCourse")
    If Not t Is Nothing Then Exit Sub

    ' 追加先＝疼痛タブのフレーム（NRSの親）
    Set host = Me.Controls("cmbNRS_Move").parent

    ' 見出し
    Set lb = host.Controls.Add("Forms.Label.1", "lblPainCourse", True)
    lb.caption = "痛みの経過・時間変化"
    lb.Left = 12
    lb.Top = 280
    lb.AutoSize = True

    ' コンテナ
    Dim fr As MSForms.Frame
    Set fr = host.Controls.Add("Forms.Frame.1", "fraPainCourse", True)
    fr.Left = lb.Left
    fr.Top = lb.Top + lb.Height + 4
    fr.Width = 610
    fr.Height = 78

    ' 発症時期
    Set lb = fr.Controls.Add("Forms.Label.1", "lblPainOnset", True)
    lb.caption = "発症時期"
    lb.Left = 12: lb.Top = 10: lb.AutoSize = True

    Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbPainOnset", True)
    cB.Left = lb.Left + 60: cB.Top = 8: cB.Width = 140
    cB.List = Array("急性（?1週）", "亜急性（?3か月）", "慢性（3か月?）", "再燃／再発", "不明")

    ' 持続時間
    Set lb = fr.Controls.Add("Forms.Label.1", "lblPainDuration", True)
    lb.caption = "持続"
    lb.Left = 260: lb.Top = 10: lb.AutoSize = True

    Set tb = fr.Controls.Add("Forms.TextBox.1", "txtPainDuration", True)
    tb.Left = lb.Left + 36: tb.Top = 8: tb.Width = 40: tb.Text = ""

    Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbPainDurationUnit", True)
    cB.Left = tb.Left + tb.Width + 6: cB.Top = 8: cB.Width = 70
    cB.List = Array("日", "週", "か月", "年")

    ' 日内変動
    Set lb = fr.Controls.Add("Forms.Label.1", "lblPainDayPeriod", True)
    lb.caption = "日内変動"
    lb.Left = 12: lb.Top = 38: lb.AutoSize = True

    Set cB = fr.Controls.Add("Forms.ComboBox.1", "cmbPainDayPeriod", True)
    cB.Left = lb.Left + 54: cB.Top = 36: cB.Width = 260
    cB.List = Array("朝に強い", "昼に強い", "夜に強い", "入浴後に軽減", "活動後に増悪", "一定で変化なし")
End Sub








Public Sub AddPainSiteUI()
    Dim host As MSForms.Frame
    Dim lb As MSForms.label
    Dim fr As MSForms.Frame
    Dim lst As MSForms.ListBox
    Dim i As Long
    Dim items As Variant

    ' 既存チェック（重複生成防止）
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainSite")
    If Not t Is Nothing Then Exit Sub

    ' 追加先＝疼痛タブのフレーム（NRSの親）
    Set host = Me.Controls("cmbNRS_Move").parent

    ' 見出し
    Set lb = host.Controls.Add("Forms.Label.1", "lblPainSite", True)
    lb.caption = "疼痛部位（複数選択可）"
    lb.Left = 12
    lb.Top = 380
    lb.AutoSize = True

    ' フレーム
    Set fr = host.Controls.Add("Forms.Frame.1", "fraPainSite", True)
    fr.Left = lb.Left
    fr.Top = lb.Top + lb.Height + 4
    fr.Width = 360
    fr.Height = 140

    ' リスト（複数選択）
    Set lst = fr.Controls.Add("Forms.ListBox.1", "lstPainSite", True)
    lst.Left = 12
    lst.Top = 10
    lst.Width = fr.Width - 24
    lst.Height = fr.Height - 20
    lst.MultiSelect = fmMultiSelectMulti
    lst.IntegralHeight = False

    ' 部位候補
    items = Array( _
        "頭部", "頸部", "肩", "肩甲部", "上腕", "肘", "前腕", "手首", "手/指", _
        "胸部", "背部上部", "背部下部（腰）", _
        "骨盤部/仙腸部", _
        "股", "大腿", "膝", "下腿", "足首", "足/趾" _
    )
    For i = LBound(items) To UBound(items)
        lst.AddItem items(i)
    Next i
End Sub









Public Sub SummarizePainUI()
    Dim fr As MSForms.Frame, s As String, i As Long
    Dim lbQ As MSForms.ListBox, lbS As MSForms.ListBox
    Dim frF As MSForms.Frame, c As Control
    Dim onset$, dura$, unit$, day$, vas$

    Set fr = Me.Controls("Frame12")

    ' 参照取得
    On Error Resume Next
    Set lbQ = fr.Controls("lstPainQual")
    Set lbS = fr.Controls("fraPainSite").Controls("lstPainSite")
    Set frF = fr.Controls("fraPainFactors")
    vas = fr.Controls("fraVAS").Controls("txtVAS").Text
    onset = fr.Controls("fraPainCourse").Controls("cmbPainOnset").Text
    dura = fr.Controls("fraPainCourse").Controls("txtPainDuration").Text
    unit = fr.Controls("fraPainCourse").Controls("cmbPainDurationUnit").Text
    day = fr.Controls("fraPainCourse").Controls("cmbPainDayPeriod").Text
    On Error GoTo 0

    ' 痛みの性質
    s = "【疼痛まとめ】"
    s = s & " 性質: "
    If Not lbQ Is Nothing Then
        For i = 0 To lbQ.ListCount - 1
            If lbQ.Selected(i) Then s = s & lbQ.List(i) & "／"
        Next
        If Right$(s, 1) = "／" Then s = Left$(s, Len(s) - 1)
    End If

    ' 部位
    s = s & "｜部位: "
    If Not lbS Is Nothing Then
        Dim tmpS As String: tmpS = ""
        For i = 0 To lbS.ListCount - 1
            If lbS.Selected(i) Then tmpS = tmpS & lbS.List(i) & "／"
        Next
        If tmpS <> "" Then tmpS = Left$(tmpS, Len(tmpS) - 1)
        s = s & tmpS
    End If

    ' 誘因・軽減因子
    s = s & "｜因子: "
    If Not frF Is Nothing Then
        Dim tmpF As String: tmpF = ""
        For Each c In frF.Controls
            If TypeName(c) = "CheckBox" Then
                If c.value = True Then tmpF = tmpF & c.caption & "／"
            End If
        Next
        If tmpF <> "" Then tmpF = Left$(tmpF, Len(tmpF) - 1)
        s = s & tmpF
    End If

    ' 経過＋VAS
    If onset <> "" Then s = s & "｜発症: " & onset
    If dura <> "" Or unit <> "" Then s = s & "｜持続: " & dura & unit
    If day <> "" Then s = s & "｜日内: " & day
    If vas <> "" Then s = s & "｜VAS: " & vas

    ' メモへ反映
    On Error Resume Next
    Me.Controls("Frame12").Controls("txtPainMemo").Text = s
End Sub



Private Sub mBtnPainSum_Click()
    SummarizePainUI
End Sub











' 疼痛タブ( Frame12 )に残っている旧UIを除去する
Public Sub RemoveLegacyPainUI()
    Dim f As MSForms.Frame, n As Variant
    Set f = Me.Controls("Frame12")

    ' Probeで[LEGACY?]と判定されたものだけ削除（新UIやNRS/備考は残す）
    For Each n In Array("Label85", "TextBox1", "Label86", "Label87", "ComboBox39", "TextBox2", "txtPainMemo_lbl", "txtPainMemo", "lblNRS_Move", "cmbNRS_Move")
        On Error Resume Next
        f.Controls.Remove CStr(n)
        If Err.Number = 0 Then Debug.Print "[removed]", n Else Debug.Print "[skip]", n, "err", Err.Number
        Err.Clear
        On Error GoTo 0
    Next n

    Debug.Print "[done] RemoveLegacyPainUI"
   
    
End Sub




Public Sub ArrangePainLayout()
    Dim f As MSForms.Frame
    Set f = Me.Controls("Frame12")

    ' 上段：左＝性質、右＝VAS
    With f.Controls("lblPainQual")
        .Left = 12: .Top = 12: .ZOrder 0
    End With
    With f.Controls("lstPainQual")
        .Left = 12
        .Top = f.Controls("lblPainQual").Top + f.Controls("lblPainQual").Height + 4
        .Width = 360: .Height = 120
        .ZOrder 0
    End With
    With f.Controls("lblVAS")
        .Left = 420: .Top = 12: .ZOrder 0
    End With
    With f.Controls("fraVAS")
        .Left = 420
        .Top = f.Controls("lblVAS").Top + f.Controls("lblVAS").Height + 4
        .ZOrder 0
    End With

    ' 中段：左＝経過、右＝誘因・軽減
    With f.Controls("lblPainCourse")
        .Left = 12: .Top = 160: .ZOrder 0
    End With
    With f.Controls("fraPainCourse")
        .Left = 12
        .Top = f.Controls("lblPainCourse").Top + f.Controls("lblPainCourse").Height + 4
        .Width = 360
        .ZOrder 0
    End With
    With f.Controls("lblPainFactors")
        .Left = 420: .Top = 160: .ZOrder 0
    End With
    With f.Controls("fraPainFactors")
        .Left = 420
        .Top = f.Controls("lblPainFactors").Top + f.Controls("lblPainFactors").Height + 4
        .Width = 330
        .ZOrder 0
    End With

    ' 下段：左＝部位、最下段＝備考
    With f.Controls("lblPainSite")
        .Left = 12: .Top = 300: .ZOrder 0
    End With
    With f.Controls("fraPainSite")
        .Left = 12
        .Top = f.Controls("lblPainSite").Top + f.Controls("lblPainSite").Height + 4
        .Width = 360: .Height = 140
        .ZOrder 0
    End With
    With f.Controls("txtPainMemo_lbl")
        .Top = 470: .ZOrder 0
    End With
    With f.Controls("txtPainMemo")
        .Top = 492: .ZOrder 0
    End With
End Sub



Sub RemoveLegacyPainUI_Final()
    Dim fr As MSForms.Frame, c As Control
    Set fr = frmEval.Controls("Frame12")
    
    For Each c In fr.Controls
        Select Case c.name
            Case "Label85", "TextBox1", "Label86", "Label87", _
                 "ComboBox39", "TextBox2", "txtPainMemo_lbl", "txtPainMemo"
                Debug.Print "[remove]", c.name
                fr.Controls.Remove c.name
        End Select
    Next
    
    Debug.Print "[done] RemoveLegacyPainUI_Final"
End Sub









Public Sub MatchPainFrameHeights()
    Dim z As MSForms.Frame, pf As MSForms.Frame, ps As MSForms.Frame, lb As MSForms.label
    Set z = Me.Controls("Frame12")
    Set pf = z.Controls("fraPainFactors")   ' 誘因・軽減因子
    Set ps = z.Controls("fraPainSite")      ' 疼痛部位
    Set lb = z.Controls("lblPainFactors")

    Dim newH As Single, bottom As Single, availH As Single
    newH = ps.Height
    bottom = ps.Top + ps.Height

    pf.Height = newH
    pf.Top = bottom - pf.Height

    availH = z.Height - 12 - pf.Top
    If pf.Height > availH Then pf.Height = availH

    lb.Left = pf.Left
    lb.Top = pf.Top - lb.Height - 4

    Debug.Print "[MatchPF] Top=", pf.Top, "H=", pf.Height
End Sub





'== frmEval のコードに貼り付け ==
Public Sub TidyPainBoxes()
    Const gap As Single = 24

    Dim z As MSForms.Frame
    Dim ps As MSForms.Frame, pf As MSForms.Frame
    Dim lbPS As MSForms.label, lbPF As MSForms.label

    Set z = Me.Controls("Frame12")
    Set ps = z.Controls("fraPainSite")
    Set pf = z.Controls("fraPainFactors")
    Set lbPS = z.Controls("lblPainSite")
    Set lbPF = z.Controls("lblPainFactors")

    ' 疼痛部位のラベルを枠直上に揃える（位置は現状維持なら不要）
    lbPS.Left = ps.Left: lbPS.Top = ps.Top - lbPS.Height - 4

    ' 誘因・軽減因子を「疼痛部位の右隣」に配置
    pf.Top = ps.Top
    pf.Left = ps.Left + ps.Width + gap
    pf.Height = ps.Height   ' 高さ一致

    ' ラベルも枠の直上に
    lbPF.Left = pf.Left
    lbPF.Top = pf.Top - lbPF.Height - 4

   



End Sub





'--------------------------------------------
' 汎用: フレーム内コントロール取得（Nothing許容）
Private Function GetCtl(ByVal host As Object, ByVal name As String) As Object
    On Error Resume Next
    Set GetCtl = host.Controls(name)
    On Error GoTo 0
End Function

' 痛みの経過・時間ブロックを恒久レイアウト
Public Sub TidyPainCourse()
    Dim f As MSForms.Frame
    Dim frCourse As MSForms.Frame, lbCourse As MSForms.label
    Dim frSite As MSForms.Frame, frFactors As MSForms.Frame
    Dim L0 As Single, T0 As Single, M As Single, gap As Single
    Dim wLeftCol As Single, rightEdge As Single
    
    Set f = Me.Controls("Frame12")
    If f Is Nothing Then Exit Sub

    ' 参照（存在しない場合は何もしない）
    Set lbCourse = GetCtl(f, "lblPainCourse")
    Set frCourse = GetCtl(f, "fraPainCourse")
    Set frSite = GetCtl(f, "fraPainSite")
    Set frFactors = GetCtl(f, "fraPainFactors")

    If lbCourse Is Nothing Or frCourse Is Nothing Then Exit Sub

    ' レイアウト定数
    M = 12        ' フレーム内の左右マージン
    gap = 8       ' コントロール間のギャップ
    L0 = frSite.Left                        ' 左列の開始位置（疼痛部位と合わせる）
    wLeftCol = frSite.Width                 ' 左列の幅を疼痛部位と一致させる

    ' 「誘因・軽減因子」を右列に下げた前提で、左列の幅で広げる
    lbCourse.Left = L0
    frCourse.Left = L0
    frCourse.Width = wLeftCol               ' ← 左列と同じ幅に恒久化
    ' 高さは中のレイアウトに依存。必要なら最後に自動で背丈調整も可

    ' 縦位置：痛みの性質の下（現在配置されている位置から少し詰める）
    ' ここでは「痛みの性質」リストの下端＋余白で合わせる
    Dim lstQual As MSForms.ListBox
    Set lstQual = GetCtl(f, "lstPainQual")
    If Not lstQual Is Nothing Then
        lbCourse.Top = lstQual.Top + lstQual.Height + 20
    Else
        ' フォールバック（現在のTopを尊重）
        lbCourse.Top = lbCourse.Top
    End If
    frCourse.Top = lbCourse.Top + lbCourse.Height + gap

    ' --- 内部項目の並び（既存の相対配置を維持しつつ幅だけ追従） ---
    '   [発症時期] [持続 数値][単位]     ←1段目
    '   [日内変動: ＿＿＿＿＿＿＿＿＿＿ ] ←2段目 幅いっぱい
    Dim lblOn As MSForms.label, cmbOn As MSForms.ComboBox
    Dim lblDur As MSForms.label, txtDur As MSForms.TextBox, cmbUnit As MSForms.ComboBox
    Dim lblDay As MSForms.label, cmbDay As MSForms.ComboBox

    Set lblOn = GetCtl(frCourse, "lblPainOnset")
    Set cmbOn = GetCtl(frCourse, "cmbPainOnset")
    Set lblDur = GetCtl(frCourse, "lblPainDuration")
    Set txtDur = GetCtl(frCourse, "txtPainDuration")
    Set cmbUnit = GetCtl(frCourse, "cmbPainDurationUnit")
    Set lblDay = GetCtl(frCourse, "lblPainDayPeriod")
    Set cmbDay = GetCtl(frCourse, "cmbPainDayPeriod")

    ' 1段目は既存のLeftを使いつつ、右端がはみ出さないように微調整
    Dim right1 As Single
    right1 = cmbUnit.Left + cmbUnit.Width
    If right1 > frCourse.Width - M Then
        cmbUnit.Left = frCourse.Width - M - cmbUnit.Width
        ' 数値・ラベルも左へ詰める
        If Not txtDur Is Nothing Then txtDur.Left = cmbUnit.Left - gap - txtDur.Width
        If Not lblDur Is Nothing Then lblDur.Left = txtDur.Left - gap - lblDur.Width - 4   ' ← ★少し余白
    Else
        ' ★通常時も軽く左に余白をとる（被り防止）
        If Not lblDur Is Nothing Then lblDur.Left = txtDur.Left - gap - lblDur.Width - 4
    End If

    ' 2段目（「日内変動」コンボ）は枠内いっぱいに
    If Not lblDay Is Nothing Then lblDay.Left = M
    If Not cmbDay Is Nothing Then
        cmbDay.Left = IIf(lblDay Is Nothing, M, lblDay.Left + lblDay.Width + gap)
        cmbDay.Width = frCourse.Width - M - cmbDay.Left
        If cmbDay.Width < 80 Then cmbDay.Width = 80
    End If

    ' 高さを自動調整（最下部のコントロール下端＋マージン）
    Dim bottomY As Single
    bottomY = 0
    Dim c As Control
    For Each c In frCourse.Controls
        bottomY = Application.WorksheetFunction.Max(bottomY, c.Top + c.Height)
    Next
    frCourse.Height = bottomY + M
End Sub
'--------------------------------------------


Public Sub WidenAndTidyPainCourse()
    ' 相互呼び出しなし／手動実行前提
    Dim f As MSForms.Frame
   Set f = Me.Controls("fraPainCourse")

    With f
        ' 参照
        Dim cmbOnset As MSForms.ComboBox
        Dim lblDur As MSForms.label
        Dim txtDur As MSForms.TextBox
        Dim cmbUnit As MSForms.ComboBox

        Set cmbOnset = .Controls("cmbPainOnset")
        Set lblDur = .Controls("lblPainDuration")
        Set txtDur = .Controls("txtPainDuration")
        Set cmbUnit = .Controls("cmbPainDurationUnit")

        ' 横並びの確定ロジック（今回「完璧」になった式と同じ）
        lblDur.Left = cmbOnset.Left + cmbOnset.Width + 12
        txtDur.Left = lblDur.Left + lblDur.Width + 12
        cmbUnit.Left = txtDur.Left + txtDur.Width + 8
       

      

        ' 右余白を24pt確保
        .Width = cmbUnit.Left + cmbUnit.Width + 24
    End With


End Sub


Public Sub TidyPainUI_Once()
    If mPainTidyBusy Then Exit Sub
    mPainTidyBusy = True
    On Error GoTo Clean



    Me.TidyPainBoxes   '※内部で WidenAndTidyPainCourse を1回だけ呼ぶ

Clean:

    mPainTidyBusy = False
    Call FixPainCaptionsAndWidth   ' ← この行だけ追加（1）
    
    '=== Pain headings finalize (once) ===
Dim f As MSForms.Frame
Dim l As MSForms.label

On Error Resume Next
'--- 直下 ---
Set l = Me.Controls("lblVAS"):           If Not l Is Nothing Then l.WordWrap = False: l.caption = "VAS（0～100）": l.WordWrap = False: l.AutoSize = False: l.Width = 120
Set l = Me.Controls("lblPainQual"):      If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
Set l = Me.Controls("lblPainCourse"):    If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
Set l = Me.Controls("lblPainSite"):      If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150
Set l = Me.Controls("lblPainFactors"):   If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150

'--- Frame3 内 ---
Set f = Me.Controls("Frame3")
If Not f Is Nothing Then
    Set l = f.Controls("lblVAS"):         If Not l Is Nothing Then l.WordWrap = False: l.caption = "VAS（0～100）": l.WordWrap = False: l.AutoSize = False: l.Width = 120
    Set l = f.Controls("lblPainQual"):    If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
    Set l = f.Controls("lblPainCourse"):  If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
    Set l = f.Controls("lblPainSite"):    If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150
    Set l = f.Controls("lblPainFactors"): If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150
End If

'--- Frame12 内 ---
Set f = Me.Controls("Frame12")
If Not f Is Nothing Then
    Set l = f.Controls("lblVAS"):         If Not l Is Nothing Then l.WordWrap = False: l.caption = "VAS（0～100）": l.WordWrap = False: l.AutoSize = False: l.Width = 120
    Set l = f.Controls("lblPainQual"):    If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
    Set l = f.Controls("lblPainCourse"):  If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 140
    Set l = f.Controls("lblPainSite"):    If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150
    Set l = f.Controls("lblPainFactors"): If Not l Is Nothing Then l.WordWrap = False: l.AutoSize = False: l.Width = 150
End If
On Error GoTo 0
'=== /Pain headings finalize ===


End Sub
Private Sub FixPainCaptionsAndWidth()
    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            ' 左：見出し（痛みの性質…）を折り返さない幅にする（親右端?24pt）
            If InStr(c.caption, "痛") > 0 And InStr(c.caption, "性質") > 0 Then
                On Error Resume Next
                c.Width = c.parent.InsideWidth - c.Left - 24
                On Error GoTo 0
            End If
            ' 右：VASの表記を固定
            If InStr(c.caption, "VAS") > 0 Then
                c.caption = "VAS（0～100）"
            End If
        End If
    Next
End Sub

Public Sub FixPainLabels_Final()
    Dim f As Control, c As Control, l As Object

    '--- 直下のラベルを処理 ---
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then
            If c.name = "lblPainQual" Then
                Set l = c
                On Error Resume Next
                CallByName l, "AutoSize", VbLet, True   ' 必要幅を取得
                CallByName l, "AutoSize", VbLet, False  ' 固定に戻す（折返し防止）
                On Error GoTo 0
            ElseIf c.name = "lblVAS" Then
                c.caption = "VAS（0～100）"
            End If
        End If
    Next

    '--- 各Frame内のラベルを処理 ---
    For Each f In Me.Controls
        If TypeName(f) = "Frame" Then
            For Each c In f.Controls
                If TypeName(c) = "Label" Then
                    If c.name = "lblPainQual" Then
                        Set l = c
                        On Error Resume Next
                        CallByName l, "AutoSize", VbLet, True
                        CallByName l, "AutoSize", VbLet, False
                        On Error GoTo 0
                    ElseIf c.name = "lblVAS" Then
                        c.caption = "VAS（0～100）"
                    End If
                End If
            Next
        End If
    Next
End Sub


Public Sub ListToneKeyCaptions()
    Dim c As Control
    For Each c In frmEval.Controls
        On Error Resume Next
        If TypeName(c) = "CheckBox" Or TypeName(c) = "OptionButton" Or TypeName(c) = "Label" Then
            Dim cap As String: cap = CStr(c.caption)
            If InStr(cap, "MAS_") > 0 Or InStr(cap, "反射_") > 0 Then
                Debug.Print "[TONE-CTL]", TypeName(c), c.name, "|", cap
            End If
        End If
        On Error GoTo 0
    Next
End Sub

Private Sub BuildWalkIndep_DistanceOutdoor()
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim lblBase As MSForms.label   ' 自立度のラベル（Label100）
    Dim cmbBase As MSForms.ComboBox
    Dim lblDist As MSForms.label
    Dim cmbDist As MSForms.ComboBox
    Dim lblOut As MSForms.label
    Dim cmbOut As MSForms.ComboBox
    Dim top1 As Single, top2 As Single, top3 As Single

       ' 「歩行」自立度フレーム取得（共通ヘルパー経由）
       ' 「歩行」と「自立」を含むフレームを探す
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Then
            Set f = ctl
            If InStr(f.caption, "歩行") > 0 And InStr(f.caption, "自立") > 0 Then
                Exit For
            End If
            Set f = Nothing
        End If
    Next

    If f Is Nothing Then
        Exit Sub
    End If



    ' ベース（自立度）取得
    On Error Resume Next
    Set lblBase = f.Controls("Label100")
    On Error GoTo 0
    If lblBase Is Nothing Then

        Exit Sub
    End If

    ' 自立度コンボ（同じ行にある ComboBox を探す）
    For Each ctl In f.Controls
        If TypeName(ctl) = "ComboBox" Then
            If Abs(ctl.Top - lblBase.Top) < 0.5 Then
                Set cmbBase = ctl
                Exit For
            End If
        End If
    Next
    If cmbBase Is Nothing Then

        Exit Sub
    End If

 cmbBase.tag = "WalkIndepLevel"

    ' 行の高さ設定
    top1 = lblBase.Top
    top2 = top1 + 24
    top3 = top2 + 24

    ' --- 距離（2段目） ---
    On Error Resume Next
    Set lblDist = f.Controls("lblWalkDistance")
    On Error GoTo 0

    If lblDist Is Nothing Then
        Set lblDist = f.Controls.Add("Forms.Label.1", "lblWalkDistance", True)
    End If
    With lblDist
        .caption = "歩行距離"
        .Left = 12
        .Top = top2
        .Width = 60
        .Height = 18
    End With

    On Error Resume Next
    Set cmbDist = f.Controls("cmbWalkDistance")
    On Error GoTo 0

    If cmbDist Is Nothing Then
        Set cmbDist = f.Controls.Add("Forms.ComboBox.1", "cmbWalkDistance", True)
    End If
    With cmbDist
        .Left = lblDist.Left + lblDist.Width + 12
        .Top = top2
        .Width = 300
        .Height = 18
        If .ListCount = 0 Then
            .AddItem "5m未満"
            .AddItem "5～10m"
            .AddItem "10～30m"
            .AddItem "30～50m"
            .AddItem "50～100m"
            .AddItem "100m以上"
        End If
    End With

    ' --- 屋外歩行（3段目） ---
    On Error Resume Next
    Set lblOut = f.Controls("lblWalkOutdoor")
    On Error GoTo 0

    If lblOut Is Nothing Then
        Set lblOut = f.Controls.Add("Forms.Label.1", "lblWalkOutdoor", True)
    End If
    With lblOut
        .caption = "屋外歩行"
        .Left = 12
        .Top = top3
        .Width = 60
        .Height = 18
    End With

    On Error Resume Next
    Set cmbOut = f.Controls("cmbWalkOutdoor")
    On Error GoTo 0

    If cmbOut Is Nothing Then
        Set cmbOut = f.Controls.Add("Forms.ComboBox.1", "cmbWalkOutdoor", True)
    End If
    With cmbOut
        .Left = lblOut.Left + lblOut.Width + 12
        .Top = top3
        .Width = 300
        .Height = 18
        If .ListCount = 0 Then
            .AddItem "屋内のみ可"
            .AddItem "屋外も短距離なら可"
            .AddItem "屋外長距離も可"
            .AddItem "屋外歩行は原則不可"
        End If
    End With


    BuildWalkIndep_Stability
    BuildWalkIndep_Speed   '★ この行を追加

End Sub



Private Sub BuildWalkIndep_Stability()
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim lblBase As MSForms.label
    Dim top1 As Single, top2 As Single, top3 As Single, top4 As Single
    Dim chk As MSForms.CheckBox
    Dim leftPos As Single
    Dim nm As Variant

    ' 「歩行」と「自立」を含むフレームを探す
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Then
            Set f = ctl
            If InStr(f.caption, "歩行") > 0 And InStr(f.caption, "自立") > 0 Then
                Exit For
            End If
            Set f = Nothing
        End If
    Next

    If f Is Nothing Then Exit Sub

    ' ベース行（Label100）の Top を基準に行位置を決める
    On Error Resume Next
    Set lblBase = f.Controls("Label100")
    On Error GoTo 0
    If lblBase Is Nothing Then Exit Sub

    top1 = lblBase.Top
    top2 = top1 + 24          ' 距離
    top3 = top2 + 24          ' 屋外歩行
    top4 = top3 + 24          ' 安定性（新規）

    ' まず既存の安定性チェックを全部削除
    For Each nm In Array( _
        "chkWalkStab_Furatsuki", _
        "chkWalkStab_Foot", _
        "chkWalkStab_Turn", _
        "chkWalkStab_Slow", _
        "chkWalkStab_FallRisk")
        On Error Resume Next
        f.Controls.Remove CStr(nm)
        On Error GoTo 0
    Next

    ' フレーム高さが足りなければ伸ばす
    If f.Height < top4 + 24 Then
        f.Height = top4 + 24
    End If

    leftPos = 12

    ' ふらつきあり
    Set chk = f.Controls.Add("Forms.CheckBox.1", "chkWalkStab_Furatsuki", True)
    With chk
        .caption = "ふらつきあり"
        .Left = leftPos
        .Top = top4
        .Width = 90
        .Height = 18
    End With
    leftPos = leftPos + chk.Width + 12

    ' 足運び不安定
    Set chk = f.Controls.Add("Forms.CheckBox.1", "chkWalkStab_Foot", True)
    With chk
        .caption = "足運び不安定"
        .Left = leftPos
        .Top = top4
        .Width = 100
        .Height = 18
    End With
    leftPos = leftPos + chk.Width + 12

    ' 方向転換不安
    Set chk = f.Controls.Add("Forms.CheckBox.1", "chkWalkStab_Turn", True)
    With chk
        .caption = "方向転換不安"
        .Left = leftPos
        .Top = top4
        .Width = 100
        .Height = 18
    End With
    leftPos = leftPos + chk.Width + 12

    ' 速度低下
    Set chk = f.Controls.Add("Forms.CheckBox.1", "chkWalkStab_Slow", True)
    With chk
        .caption = "速度低下"
        .Left = leftPos
        .Top = top4
        .Width = 80
        .Height = 18
    End With
    leftPos = leftPos + chk.Width + 12

    ' 転倒リスク高い
    Set chk = f.Controls.Add("Forms.CheckBox.1", "chkWalkStab_FallRisk", True)
    With chk
        .caption = "転倒リスク高い"
        .Left = leftPos
        .Top = top4
        .Width = 110
        .Height = 18
    End With
End Sub


Private Sub BuildWalkIndep_Speed()
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim lblBase As MSForms.label
    Dim top1 As Single, top2 As Single, top3 As Single, top4 As Single, top5 As Single
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox

    ' 「歩行」と「自立」を含むフレームを探す
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Then
            Set f = ctl
            If InStr(f.caption, "歩行") > 0 And InStr(f.caption, "自立") > 0 Then
                Exit For
            End If
            Set f = Nothing
        End If
    Next

    If f Is Nothing Then Exit Sub

    ' ベース行（Label100）の Top から段を決める
    On Error Resume Next
    Set lblBase = f.Controls("Label100")
    On Error GoTo 0
    If lblBase Is Nothing Then Exit Sub

    top1 = lblBase.Top          ' 自立度
    top2 = top1 + 24            ' 距離
    top3 = top2 + 24            ' 屋外歩行
    top4 = top3 + 24            ' 安定性
    top5 = top4 + 24            ' ★歩行速度（新規）

    ' フレーム高さを必要に応じて伸ばす
    If f.Height < top5 + 24 Then
        f.Height = top5 + 24
    End If

    ' ラベル（歩行速度）
    On Error Resume Next
    Set lbl = f.Controls("lblWalkSpeed")
    On Error GoTo 0
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblWalkSpeed", True)
    End If
    With lbl
        .caption = "歩行速度"
        .Left = 12
        .Top = top5
        .Width = 60
        .Height = 18
    End With

    ' コンボボックス（速度区分）
    On Error Resume Next
    Set cmb = f.Controls("cmbWalkSpeed")
    On Error GoTo 0
    If cmb Is Nothing Then
        Set cmb = f.Controls.Add("Forms.ComboBox.1", "cmbWalkSpeed", True)
    End If
    With cmb
        .Left = lbl.Left + lbl.Width + 12
        .Top = top5
        .Width = 200
        .Height = 18
        If .ListCount = 0 Then
            .AddItem "速い"
            .AddItem "やや速い"
            .AddItem "ふつう"
            .AddItem "やや遅い"
            .AddItem "遅い"
        End If
    End With
End Sub

Private Sub BuildWalk_AbnormalTab()
    Dim ctl As MSForms.Control
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page

    ' 歩行評価用の MultiPage2 を探す
    For Each ctl In Me.Controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "MultiPage2" Then
                Set mp = ctl
                Exit For
            End If
        End If
    Next

    If mp Is Nothing Then Exit Sub

    ' 既に「異常歩行」タブがあれば何もしない
    On Error Resume Next
    Set pg = mp.Pages("pgWalkAbnormal")
    On Error GoTo 0
    If Not pg Is Nothing Then Exit Sub

    ' 新しいページを追加（Index=2想定：自立度, RLA の次）
    Set pg = mp.Pages.Add
    pg.name = "pgWalkAbnormal"
    pg.caption = "異常歩行"

    ' ひとまず中に空のフレームだけ置いておく（中身は後で作る）
    Dim f As MSForms.Frame
    Set f = pg.Controls.Add("Forms.Frame.1", "fraWalkAbnormal", True)
    With f
        .caption = "異常歩行パターン（チェック）"
        .Left = 6
        .Top = 6
        .Width = mp.Width - 24
        .Height = mp.Height - 24
    End With
End Sub


Private Sub BuildWalkAbnormal_Frames()
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim w As Single, h As Single

    ' MultiPage2（歩行評価）を取得
    For Each ctl In Me.Controls
        If TypeName(ctl) = "MultiPage" And ctl.name = "MultiPage2" Then
            Set mp = ctl
            Exit For
        End If
    Next
    If mp Is Nothing Then Exit Sub

    ' 異常歩行ページ取得
    On Error Resume Next
    Set pg = mp.Pages("pgWalkAbnormal")
    On Error GoTo 0
    If pg Is Nothing Then Exit Sub

    ' ページのワークエリアサイズ
    w = mp.Width - 24
    h = mp.Height - 24

    ' 既存フレーム削除（再生成用）
    For Each ctl In pg.Controls
        If TypeName(ctl) = "Frame" Then
            pg.Controls.Remove ctl.name
        End If
    Next

    ' --- A：片麻痺系 ---
    Set f = pg.Controls.Add("Forms.Frame.1", "fraWalkAbn_A", True)
    With f
        .caption = "A：片麻痺・脳血管障害パターン"
        .Left = 6
        .Top = 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- B：パーキンソン系 ---
    Set f = pg.Controls.Add("Forms.Frame.1", "fraWalkAbn_B", True)
    With f
        .caption = "B：パーキンソン関連パターン"
        .Left = w / 2 + 6
        .Top = 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- C：整形・高齢者不安定歩行 ---
    Set f = pg.Controls.Add("Forms.Frame.1", "fraWalkAbn_C", True)
    With f
        .caption = "C：整形・高齢者不安定歩行"
        .Left = 6
        .Top = h / 2 + 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- D：協調障害・失調 ---
    Set f = pg.Controls.Add("Forms.Frame.1", "fraWalkAbn_D", True)
    With f
        .caption = "D：協調障害・失調パターン"
        .Left = w / 2 + 6
        .Top = h / 2 + 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With
End Sub


Private Sub BuildWalkAbnormal_Checks()
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim chk As MSForms.CheckBox
    Dim items As Variant
    Dim i As Long, topPos As Single
    
    ' MultiPage2 を取得
    For Each ctl In Me.Controls
        If TypeName(ctl) = "MultiPage" And ctl.name = "MultiPage2" Then
            Set mp = ctl
            Exit For
        End If
    Next
    If mp Is Nothing Then Exit Sub

    ' 異常歩行ページを取得
    On Error Resume Next
    Set pg = mp.Pages("pgWalkAbnormal")
    On Error GoTo 0
    If pg Is Nothing Then Exit Sub

    ' 4つのフレームを順番に処理
    Dim fNames As Variant
    fNames = Array("fraWalkAbn_A", "fraWalkAbn_B", "fraWalkAbn_C", "fraWalkAbn_D")
    
    Dim A_items As Variant
    Dim B_items As Variant
    Dim C_items As Variant
    Dim D_items As Variant

    ' ----------------------
    ' A：片麻痺・脳血管障害（1～10）
    ' ----------------------
    A_items = Array( _
        "すり足歩行", _
        "ぶん回し歩行", _
        "反張膝歩行（膝過伸展）", _
        "トレンデレンブルグ歩行", _
        "デュシェンヌ歩行", _
        "下垂足（フットスラップ）", _
        "共同運動パターンの強さ", _
        "骨盤後傾／体幹後傾の立脚", _
        "片脚立脚の著しい短縮", _
        "足部クリアランス不良" _
    )

    ' ----------------------
    ' B：パーキンソン系（1～9）
    ' ----------------------
    B_items = Array( _
        "小刻み歩行", _
        "前傾姿勢歩行", _
        "フリーズ（突然停止）", _
        "歩行開始困難（スタート hesitation）", _
        "突進歩行（フェスティネーション）", _
        "歩幅減少", _
        "手の振り消失", _
        "方向転換困難", _
        "リズム性消失" _
    )

    ' ----------------------
    ' C：整形・高齢者不安定歩行（1～10）
    ' ----------------------
    C_items = Array( _
        "よちよち歩行（筋力低下）", _
        "膝折れ（knee buckling）", _
        "股OAの疼痛性歩行", _
        "体幹左右揺れ（ヤコビー徴候様）", _
        "前足部荷重がしにくい", _
        "歩幅のばらつき", _
        "靴の引きずり", _
        "杖・歩行器への強い依存", _
        "片脚荷重回避", _
        "疼痛回避性の異常歩容" _
    )

    ' ----------------------
    ' D：協調障害・失調（1～8）
    ' ----------------------
    D_items = Array( _
        "失調性歩行（ワイドベース）", _
        "千鳥足歩行", _
        "ステッピング歩行", _
        "ぎくしゃくした歩行", _
        "方向転換時の大きな揺れ", _
        "上肢との協調不良", _
        "ふらつき大", _
        "足の位置決めが不正確" _
    )

    ' --------- ここから生成処理 ---------

    Dim listSets As Variant
    listSets = Array(A_items, B_items, C_items, D_items)

    Dim idx As Long, arr As Variant

    For idx = 0 To 3
        ' 対象フレーム取得
        Set f = pg.Controls(fNames(idx))
        If f Is Nothing Then GoTo ContinueNext
        
        ' 既存チェック削除
        For Each ctl In f.Controls
            If TypeName(ctl) = "CheckBox" Then
                f.Controls.Remove ctl.name
            End If
        Next ctl
        
                ' 追加
        arr = listSets(idx)
        topPos = 24

        Dim maxBottom As Single
        maxBottom = 0

        For i = LBound(arr) To UBound(arr)
            Set chk = f.Controls.Add("Forms.CheckBox.1", fNames(idx) & "_chk" & CStr(i), True)
            With chk
                .caption = arr(i)
                .Left = 12
                .Top = topPos
                .Width = f.Width - 24
                .Height = 18
            End With

            ' このチェックボックスの下端を記録
            If chk.Top + chk.Height > maxBottom Then
                maxBottom = chk.Top + chk.Height
            End If

            topPos = topPos + 20
        Next i
        ' 中身に合わせてフレーム高さを最低限まで伸ばす（ログ付き）
        If maxBottom + 12 > f.Height Then
#If APP_DEBUG Then
            Debug.Print "[ABN-RESIZE]", f.name, _
                        "oldH=", f.Height, _
                        "maxBottom=", maxBottom, _
                        "newH=", maxBottom + 12
#End If
            f.Height = maxBottom + 12
        End If


ContinueNext:
    Next idx
    
    


End Sub

Private Sub BuildWalkUI_All()
    ' 歩行 自立度タブ（距離・屋外・安定性・速度）
    BuildWalkIndep_DistanceOutdoor
    
    ' 「異常歩行」タブ＋中身（4分類フレーム＋チェック群）
    BuildWalk_AbnormalTab
    BuildWalkAbnormal_Frames
    BuildWalkAbnormal_Checks
    FixWalkRootFrameHeight
End Sub



Public Sub BuildCogMentalUI_Simple()
    Dim f As MSForms.Frame
    Dim c As MSForms.Control
    Dim mp As MSForms.MultiPage
    Dim fw As MSForms.Frame   '★ 歩行タブのフレーム

    ' 親フレーム（認知機能・精神面）
    On Error Resume Next
    Set f = Me.Controls("Frame31")
    On Error GoTo 0

    If f Is Nothing Then
        MsgBox "Frame31（認知機能・精神面）が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' ★歩行タブ(Frame6)と同じ位置・サイズに合わせる
    On Error Resume Next
    Set fw = Me.Controls("Frame6")
    On Error GoTo 0
    If Not fw Is Nothing Then
        f.Left = fw.Left
        f.Top = fw.Top
        f.Width = fw.Width
        'f.Height = fw.Height
    End If

    ' いったん中身を全部クリア（元のラベル／コンボ／既存マルチページも含めて）
    Do While f.Controls.Count > 0
        f.Controls.Remove f.Controls(0).name
    Loop

    ' 子MultiPageを追加（認知機能 / 精神面 の2タブのみ）
    Set mp = f.Controls.Add("Forms.MultiPage.1", "mpCogMental", True)
           With mp
        .Left = 6
        .Top = 0          ' ← ここを 6 → 0 に
        .Width = f.Width - 12
        .Height = f.Height - 12
        .Style = fmTabStyleTabs
        .TabOrientation = fmTabOrientationTop
    End With



    ' 既存ページを全部消してから2ページ作成
    Do While mp.Pages.Count > 0
        mp.Pages.Remove 0
    Loop

    mp.Pages.Add
    mp.Pages.Add

    mp.Pages(0).caption = "認知機能"
    mp.Pages(0).name = "pgCognition"

    mp.Pages(1).caption = "精神面"
    mp.Pages(1).name = "pgMental"
End Sub




Public Sub BuildCog_CognitionCore()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim c As MSForms.Control
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    
    ' 親フレーム（認知）
    On Error Resume Next
    Set f = Me.Controls("Frame31")
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "Frame31 が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 子マルチページ
    On Error Resume Next
    Set mp = f.Controls("mpCogMental")
    On Error GoTo 0
    If mp Is Nothing Then
        MsgBox "mpCogMental が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 認知機能ページ
    On Error Resume Next
    Set pg = mp.Pages("pgCognition")
    On Error GoTo 0
    If pg Is Nothing Then
        MsgBox "pgCognition ページが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 既存コントロールをクリア（やり直し用）
    Do While pg.Controls.Count > 0
        pg.Controls.Remove pg.Controls(0).name
    Loop
    
    ' レイアウト基準
    Dim rowTop As Single, rowGap As Single
    Dim col1Left As Single, col2Left As Single
    Dim lblW As Single, cmbW As Single
    
    rowTop = 18
    rowGap = 24
    col1Left = 12
    col2Left = 260
    lblW = 60
    cmbW = 140
    
    ' 共通で使う評価リスト
    Dim i As Long
    Dim arr4()
    
    arr4 = Array("正常", "やや低下", "低下", "著明に低下")
    
    '―― 1行目：記憶／注意 ――
    ' 記憶
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogMemory", True)
    With lbl
        .caption = "記憶"
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogMemory", True)
    With cmb
        .Left = col1Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        For i = LBound(arr4) To UBound(arr4)
            .AddItem arr4(i)
        Next i
    End With
    
    ' 注意
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogAttention", True)
    With lbl
        .caption = "注意"
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogAttention", True)
    With cmb
        .Left = col2Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        For i = LBound(arr4) To UBound(arr4)
            .AddItem arr4(i)
        Next i
    End With
    
    '―― 2行目：見当識／判断 ――
    rowTop = rowTop + rowGap
    
    ' 見当識
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogOrientation", True)
    With lbl
        .caption = "見当識"
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogOrientation", True)
    With cmb
        .Left = col1Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        For i = LBound(arr4) To UBound(arr4)
            .AddItem arr4(i)
        Next i
    End With
    
    ' 判断
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogJudgement", True)
    With lbl
        .caption = "判断"
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogJudgement", True)
    With cmb
        .Left = col2Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "良好"
        .AddItem "やや不安定"
        .AddItem "不安定"
    End With
    
    '―― 3行目：遂行／言語 ――
    rowTop = rowTop + rowGap
    
    ' 遂行
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogExecutive", True)
    With lbl
        .caption = "遂行"
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogExecutive", True)
    With cmb
        .Left = col1Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "良好"
        .AddItem "やや不安定"
        .AddItem "不安定"
    End With
    
    ' 言語
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblCogLanguage", True)
    With lbl
        .caption = "言語"
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbCogLanguage", True)
    With cmb
        .Left = col2Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "問題なし"
        .AddItem "やや障害"
        .AddItem "障害顕著"
    End With
End Sub



Public Sub BuildCog_DementiaBlock()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim fraTop As Single
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    Dim txt As MSForms.TextBox
    
    ' 親フレーム
    On Error Resume Next
    Set f = Me.Controls("Frame31")
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "Frame31 が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 子マルチページ
    On Error Resume Next
    Set mp = f.Controls("mpCogMental")
    On Error GoTo 0
    If mp Is Nothing Then
        MsgBox "mpCogMental が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 認知機能ページ
    On Error Resume Next
    Set pg = mp.Pages("pgCognition")
    On Error GoTo 0
    If pg Is Nothing Then
        MsgBox "pgCognition ページが見つかりません。", vbExclamation
        Exit Sub
    End If
    
       ' いったん、既存の認知症ブロックを消す（やり直し用）
    Dim i As Long
    For i = pg.Controls.Count - 1 To 0 Step -1
        With pg.Controls(i)
            If .name = "lblDementiaType" _
               Or .name = "cmbDementiaType" _
               Or .name = "lblDementiaNote" _
               Or .name = "txtDementiaNote" Then
                pg.Controls.Remove .name
            End If
        End With
    Next i

    
    ' 上の認知6項目ブロックのすぐ下に配置（だいたい3行＋余白ぶん下げる）
    fraTop = 18 + 3 * 24 + 18   ' 18(最初) + 3行*24 + 余白
    
    ' 見出しラベル「認知症の種類」
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblDementiaType", True)
    With lbl
        .caption = "認知症の種類"
        .Left = 12
        .Top = fraTop
        .Width = 90
        .Height = 18
    End With
    
    ' 診断名コンボ
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbDementiaType", True)
    With cmb
        .Left = lbl.Left + lbl.Width + 6
        .Top = fraTop - 2
        .Width = 160
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "なし / 不明"
        .AddItem "アルツハイマー型"
        .AddItem "血管性"
        .AddItem "レビー小体型"
        .AddItem "前頭側頭型(FTD)"
        .AddItem "混合型"
        .AddItem "その他"
    End With
    
    ' 備考ラベル
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblDementiaNote", True)
    With lbl
        .caption = "備考"
        .Left = cmb.Left + cmb.Width + 12
        .Top = fraTop
        .Width = 40
        .Height = 18
    End With
    
    ' 備考テキスト
    Set txt = pg.Controls.Add("Forms.TextBox.1", "txtDementiaNote", True)
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = fraTop - 2
        .Width = mp.Width - .Left - 12
        .Height = 18
    End With
End Sub



Public Sub BuildCog_BPSD()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim topY As Single
    Dim lbl As MSForms.label
    Dim chk As MSForms.CheckBox
    Dim i As Long
    
    ' 親フレーム
    On Error Resume Next
    Set f = Me.Controls("Frame31")
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "Frame31 が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 子マルチページ
    On Error Resume Next
    Set mp = f.Controls("mpCogMental")
    On Error GoTo 0
    If mp Is Nothing Then
        MsgBox "mpCogMental が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 認知機能ページ
    On Error Resume Next
    Set pg = mp.Pages("pgCognition")
    On Error GoTo 0
    If pg Is Nothing Then
        MsgBox "pgCognition が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' --- 既存BPSDコントロール削除 ---
    Dim c As MSForms.Control
    For i = pg.Controls.Count - 1 To 0 Step -1
        If TypeName(pg.Controls(i)) = "CheckBox" _
           Or pg.Controls(i).name Like "lblBPSD*" Then
            pg.Controls.Remove pg.Controls(i).name
        End If
    Next i
    
    ' --- 追加位置（認知症の種類ブロックの下） ---
    topY = 18 + 3 * 24 + 18 + 24   '6項目ブロック + 余白
    topY = topY + 24               '認知症種類の行
    
    ' 見出し
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblBPSD_Title", True)
    With lbl
        .caption = "認知症の周辺症状（BPSD）"
        .Left = 12
        .Top = topY
        .Width = 180
        .Height = 18
    End With
    
    topY = topY + 24
    
    ' BPSD項目
    Dim items
    items = Array("抑うつ", "不安", "焦燥", "幻覚", "妄想", _
                  "徘徊", "暴言", "暴力", "不穏", "睡眠障害", "昼夜逆転")
    
    Dim col As Long, row As Long
    col = 0: row = 0
    
    For i = LBound(items) To UBound(items)
        Set chk = pg.Controls.Add("Forms.CheckBox.1", "chkBPSD" & CStr(i), True)
        With chk
            .caption = items(i)
            .Left = 12 + (col * 140)
            .Top = topY + (row * 22)
            .Width = 130
            .Height = 18
        End With
        
        col = col + 1
        If col = 3 Then
            col = 0
            row = row + 1
        End If
    Next i
End Sub



Public Sub BuildCog_MentalBlock()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    Dim txt As MSForms.TextBox
    Dim topY As Single
    
    ' 親フレーム（Frame31）
    On Error Resume Next
    Set f = Me.Controls("Frame31")
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "Frame31 が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 子マルチページ mpCogMental
    On Error Resume Next
    Set mp = f.Controls("mpCogMental")
    On Error GoTo 0
    If mp Is Nothing Then
        MsgBox "mpCogMental が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 精神面ページ
    On Error Resume Next
    Set pg = mp.Pages("pgMental")
    On Error GoTo 0
    If pg Is Nothing Then
        MsgBox "pgMental が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 既存クリア（やり直し用）
    Dim i As Long
    For i = pg.Controls.Count - 1 To 0 Step -1
        pg.Controls.Remove pg.Controls(i).name
    Next i
    
    ' レイアウト
    Dim rowGap As Single: rowGap = 26
    Dim lblW As Single: lblW = 90
    Dim cmbW As Single: cmbW = 150
    Dim left1 As Single: left1 = 12
    Dim left2 As Single: left2 = 260
    topY = 18
    
    ' --- 気分 ---
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblMood", True)
    With lbl
        .caption = "気分"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbMood", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "安定"
        .AddItem "やや不安定"
        .AddItem "不安定"
    End With
    
    ' --- 意欲 ---
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblMotivation", True)
    With lbl
        .caption = "意欲"
        .Left = left2
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbMotivation", True)
    With cmb
        .Left = left2 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "高い"
        .AddItem "普通"
        .AddItem "低い"
        .AddItem "ほとんどなし"
    End With
    
    ' --- 不安 ---
    topY = topY + rowGap
    
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblAnxiety", True)
    With lbl
        .caption = "不安"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbAnxiety", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "なし"
        .AddItem "軽度"
        .AddItem "中等度"
        .AddItem "強い"
    End With
    
    ' --- 対人 ---
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblRelation", True)
    With lbl
        .caption = "対人関係"
        .Left = left2
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbRelation", True)
    With cmb
        .Left = left2 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "良好"
        .AddItem "おおむね良好"
        .AddItem "やや問題"
        .AddItem "問題あり"
    End With
    
    ' --- 睡眠 ---
    topY = topY + rowGap
    
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblSleep", True)
    With lbl
        .caption = "睡眠"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.Controls.Add("Forms.ComboBox.1", "cmbSleep", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "良好"
        .AddItem "入眠困難"
        .AddItem "中途覚醒"
        .AddItem "早朝覚醒"
        .AddItem "日中傾眠"
    End With
    
    ' --- 備考 ---
    topY = topY + rowGap + 8
    
    Set lbl = pg.Controls.Add("Forms.Label.1", "lblMentalNote", True)
    With lbl
        .caption = "備考"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set txt = pg.Controls.Add("Forms.TextBox.1", "txtMentalNote", True)
    With txt
        .Left = left1 + lblW + 6
        .Top = topY - 2
         .Width = mp.Width - .Left - 12
        .Height = 50
        .multiline = True
        .EnterKeyBehavior = True
    End With
End Sub





Private Sub BuildDailyLogTab()
    Dim mp As Object
    Dim pg As MSForms.Page
    Dim i As Long
    Dim exists As Boolean
    Dim fra As MSForms.Frame

    On Error GoTo EH

    '=== MultiPage1 を取得 ===
    Set mp = Me.Controls("MultiPage1")

    '=== すでに「日々の記録」タブがあるか確認（冪等用）===
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = "日々の記録" Then
            exists = True
            Set pg = mp.Pages(i)
            Exit For
        End If
    Next i

    '=== なければ新しいページを追加 ===
    If Not exists Then
        Set pg = mp.Pages.Add
        pg.caption = "日々の記録"
    End If

    '=== フレームが無ければ1個だけ作る ===
    On Error Resume Next
    Set fra = pg.Controls("fraDailyLog")
    On Error GoTo EH

    If fra Is Nothing Then
        Set fra = pg.Controls.Add("Forms.Frame.1", "fraDailyLog")
        fra.caption = "日々の記録"
        fra.Left = 6
    fra.Top = 6
    fra.Width = mp.Width - 24      ' ← MultiPage の幅から左右12ptずつ余白
    fra.Height = mp.Height - 30    ' ← 上下のタブ＋余白ぶんを差し引いて枠いっぱい
    End If

       BuildDailyLogLayout
       BuildDailyLog_StaffAndNote

    Exit Sub

EH:


   


End Sub



Private Sub BuildDailyLogLayout()
    On Error GoTo EH

    Dim mp As Object
    Dim pg As Object
    Dim f As Object
    Dim lbl As Object
    Dim txt As Object
    Dim i As Long

        '=== 日々の記録フレーム取得（共通ヘルパー経由） ===
    Set f = GetDailyLogFrame()
    If f Is Nothing Then GoTo ExitHere

    '========================================
    ' 上段：記録入力ゾーン
    '========================================

    '=== 記録日ラベル ===
    On Error Resume Next
    Set lbl = f.Controls("lblDailyDate")
    On Error GoTo EH
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyDate")
    End If
    With lbl
        .caption = "記録日"
        .Left = 12
        .Top = 18
        .Width = 40
        .Height = 18
    End With

    '=== 記録日テキスト ===
    On Error Resume Next
    Set txt = f.Controls("txtDailyDate")
    On Error GoTo EH
    If txt Is Nothing Then
        Set txt = f.Controls.Add("Forms.TextBox.1", "txtDailyDate")
    End If
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = lbl.Top - 2
        .Width = 80
        .Height = 18
    End With

    '=== 記録者ラベル ===
   On Error Resume Next

    Set lbl = f.Controls("lblDailyStaff")
    On Error GoTo EH
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyStaff")
    End If
    With lbl
        .caption = "記録者"
        .Left = txt.Left + txt.Width + 24
        .Top = 18
        .Width = 40
        .Height = 18
    End With

    '=== 記録者テキスト ===
    On Error Resume Next
    Set txt = f.Controls("txtDailyStaff")
    On Error GoTo EH
    If txt Is Nothing Then
        Set txt = f.Controls.Add("Forms.TextBox.1", "txtDailyStaff")
    End If
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = lbl.Top - 2
        .Width = 100
        .Height = 18
    End With

    '=== 記録内容ラベル ===
    On Error Resume Next
    Set lbl = f.Controls("lblDailyNote")
    On Error GoTo EH
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyNote")
    End If
    With lbl
        .caption = "記録内容"
        .Left = 12
        .Top = 48
        .Width = 60
        .Height = 18
    End With

    '=== 記録内容テキスト（マルチライン） ===
    On Error Resume Next
    Set txt = f.Controls("txtDailyNote")
    On Error GoTo EH
    If txt Is Nothing Then
        Set txt = f.Controls.Add("Forms.TextBox.1", "txtDailyNote")
    End If
    With txt
        .Left = 12
        .Top = 66
        .Width = f.Width - 24
        .Height = f.Height - .Top - 12
        .multiline = True
        .EnterKeyBehavior = True
        .ScrollBars = 2   ' fmScrollBarsVertical
    End With

ExitHere:
    Exit Sub

EH:

    Resume ExitHere
End Sub


Private Sub BuildDailyLog_StaffAndNote()
    On Error GoTo EH
    
    Dim mp As Object
    Dim pg As Object
    Dim f As Object
    Dim lbl As Object
    Dim txt As Object
    Dim i As Long

       '=== 日々の記録フレーム取得（共通ヘルパー経由） ===
    Set f = GetDailyLogFrame()
    If f Is Nothing Then GoTo ExitHere

    
    '----------------------------------------
    ' 記録日（既存があっても位置を揃える）
    '----------------------------------------
    On Error Resume Next
    Set lbl = f.Controls("lblDailyDate")
    On Error GoTo EH
    If Not lbl Is Nothing Then
        With lbl
            .caption = "記録日"
            .Left = 12
            .Top = 18
            .Width = 40
            .Height = 18
        End With
    End If

    On Error Resume Next
    Set txt = f.Controls("txtDailyDate")
    On Error GoTo EH
    If Not txt Is Nothing Then
        With txt
            .Left = lbl.Left + lbl.Width + 6
            .Top = lbl.Top - 2
            .Width = 80
            .Height = 18
        End With
    End If
    
    '----------------------------------------
    ' 記録者ラベル
    '----------------------------------------
    Set lbl = Nothing

    On Error Resume Next
    Set lbl = f.Controls("lblDailyStaff")
    On Error GoTo EH
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyStaff")
    End If
    With lbl
        .caption = "記録者"
        .Left = txt.Left + txt.Width + 24
        .Top = 18
        .Width = 40
        .Height = 18
    End With

    '----------------------------------------
    ' 記録者テキスト
    '----------------------------------------
    Set txt = Nothing
    
    On Error Resume Next
    Set txt = f.Controls("txtDailyStaff")
    On Error GoTo EH
    If txt Is Nothing Then
        Set txt = f.Controls.Add("Forms.TextBox.1", "txtDailyStaff")
    End If
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = lbl.Top - 2
        .Width = 100
        .Height = 18
    End With

    '----------------------------------------
    ' 記録内容ラベル
    '----------------------------------------
    Set lbl = Nothing
    
    On Error Resume Next
    Set lbl = f.Controls("lblDailyNote")
    On Error GoTo EH
    If lbl Is Nothing Then
        Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyNote")
    End If
    With lbl
        .caption = "記録内容"
        .Left = 12
        .Top = 48
        .Width = 60
        .Height = 18
    End With

    '----------------------------------------
    ' 記録内容テキスト（マルチライン）
    '----------------------------------------
    Set txt = Nothing
    
    On Error Resume Next
    Set txt = f.Controls("txtDailyNote")
    On Error GoTo EH
    If txt Is Nothing Then
        Set txt = f.Controls.Add("Forms.TextBox.1", "txtDailyNote")
    End If
    With txt
        .Left = 12
        .Top = 66
        .Width = f.Width - 24
        .Height = f.Height - 280   ' ← ここは元のままに戻す
        .multiline = True
        .EnterKeyBehavior = True
        .ScrollBars = 2   ' fmScrollBarsVertical
    End With

ExitHere:
    Exit Sub

EH:
    Resume ExitHere
End Sub




Public Sub BuildDailyLog_HistoryList(owner As Object)
    Dim f As Object
    Dim txtNote As Object
    Dim lst As MSForms.ListBox
    Dim topPos As Single
    Dim margin As Single

    margin = 12

    ' fraDailyLog と 記録内容テキストを取得
    Set f = owner.Controls("fraDailyLog")
    Set txtNote = f.Controls("txtDailyNote")

    ' すでに作ってある場合はいったん削除して作り直し（冪等性確保）
    On Error Resume Next
    f.Controls.Remove "lstDailyLogList"
    On Error GoTo 0

'--- 履歴ラベル作成 ---
Dim lbl As MSForms.label

On Error Resume Next
f.Controls.Remove "lblDailyHistory"
On Error GoTo 0

Set lbl = f.Controls.Add("Forms.Label.1", "lblDailyHistory", True)

With lbl
    .caption = "この月の記録一覧"
    .Left = txtNote.Left
    .Top = txtNote.Top + txtNote.Height + 15    ' ← リストBOXより少し上
    .Width = 200
    .Height = 18
    .Font.Bold = True
End With




    ' ListBox 追加
    Set lst = f.Controls.Add("Forms.ListBox.1", "lstDailyLogList", True)

   topPos = f.Controls("lblDailyHistory").Top + f.Controls("lblDailyHistory").Height + 4



    With lst
        .Left = txtNote.Left
        .Top = topPos
        .Width = txtNote.Width
        .Height = f.Height - .Top - 8
        .ColumnCount = 3          ' 記録年月 / 名前 / 記録内容
        .ColumnHeads = False
        .IntegralHeight = False
        .ColumnWidths = "70 pt;0 pt;9999 pt"

    End With

Call owner.HookDailyLogList(lst)


End Sub







Public Sub BuildDailyLog_ExtractButton(owner As Object)
    Dim f As Object
    Dim txtStaff As Object
    Dim cmd As MSForms.CommandButton
    Dim margin As Single

    margin = 12

    ' fraDailyLog と 記録者テキストを取得
    Set f = owner.Controls("fraDailyLog")
    Set txtStaff = f.Controls("txtDailyStaff")
    BuildDailyLog_HistoryList owner   ' ★これを追加（ListBoxを必ず作る）

    ' 既にボタンがあれば削除して作り直し（冪等）
    On Error Resume Next
    f.Controls.Remove "cmdDailyExtract"
    On Error GoTo 0

    ' 抽出ボタン追加
    Set cmd = f.Controls.Add("Forms.CommandButton.1", "cmdDailyExtract", True)

    With cmd
        .caption = "この月の記録一覧"
        .Width = 120
        .Height = 24
        .Top = txtStaff.Top
        .Left = txtStaff.Left + txtStaff.Width + margin + 140

    End With
    
     Set mDailyExtract = cmd


End Sub

Public Sub BuildDailyLog_SaveButton(owner As Object)
    Dim f As Object
    Dim txtStaff As Object
    Dim cmd As MSForms.CommandButton
    Dim margin As Single

    margin = 12

    ' fraDailyLog と 記録者テキストを取得
    Set f = owner.Controls("fraDailyLog")
    Set txtStaff = f.Controls("txtDailyStaff")

    ' 既にボタンがあれば削除して作り直し（冪等）
    On Error Resume Next
    f.Controls.Remove "cmdDailySave"
    On Error GoTo 0

    ' 保存ボタン追加
    Set cmd = f.Controls.Add("Forms.CommandButton.1", "cmdDailySave", True)

    With cmd
        .caption = "日々の記録を保存"
        .Width = 110
        .Height = 24
        .Top = txtStaff.Top
        .Left = txtStaff.Left + txtStaff.Width + margin
    End With
  
    Set mDailySave = cmd


End Sub



Private Sub mDailyExtract_Click()
    ' ① 材料文を作る
    Call Me.BuildMonthlyDraft_FromDailyLog
    
    
    Dim box As Object
      Set box = Me.Controls("fraDailyLog").Controls("txtMonthlyMonitoringDraft")

        If InStr(1, box.value, "（この月の記録はありません）", vbTextCompare) > 0 Then
            
            box.value = "【月次モニタリング下書き】" & vbCrLf & _
            "対象：" & Me.Controls("frHeader").Controls("txtHdrName").value & vbCrLf & _
            "期間：" & Format$(DateSerial(Year(CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value)), _
                                      Month(CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value)), 1), "yyyy/mm/dd") & _
            " - " & _
            Format$(DateSerial(Year(CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value)), _
                                Month(CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value)) + 1, 0), "yyyy/mm/dd") & vbCrLf & vbCrLf & _
            "■ この月に記録された特記事項" & vbCrLf & _
            "この月は特記事項となる記録はありませんでした。" & vbCrLf & _
            "体調面に大きな変動はなく、日々のリハビリにも安定して取り組まれていました。" & vbCrLf & _
            "今後も現在の状態を維持できるよう、引き続き経過を観察していきます。"

                      Call ExportMonitoring_ToMonthlyWorkbook( _
                CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value), _
                Me.Controls("frHeader").Controls("txtHdrName").value, _
                box.value)

           Exit Sub

        End If

    
    

    ' ② AIで下書きに変換
    Me.Controls("fraDailyLog").Controls("txtMonthlyMonitoringDraft").value = _
        OpenAI_BuildDraft( _
            "【出力フォーマット厳守】" & vbCrLf & _
"以下の見出しを、表記・順序・記号（■）を一切変えずに必ず出力すること。" & vbCrLf & _
"見出しの追加・削除・言い換え禁止。装飾（★/【】/番号付け）禁止。" & vbCrLf & _
"必ずこの順序：" & vbCrLf & _
"■ この月に記録された特記事項" & vbCrLf & _
"■ コメント・考察" & vbCrLf & vbCrLf & _
"・本文（経過・時系列）には、事実のみを記載する。記録に書かれていない事実や推測は、本文には含めない。「コメント・考察」欄に限り、記録内容を踏まえた今後の観察視点や留意点を記載してよい。その際は、断定を避け、「○○の可能性がある」「○○に留意して経過を確認する」などの表現に限定する。医学的判断、改善・悪化の断定、因果関係の断定は行わない。文体は「です・ます調」とし、現場記録として自然で読みやすい柔らかさを持たせる。", _
            Me.Controls("fraDailyLog").Controls("txtMonthlyMonitoringDraft").value _
        )
        
        
            Call ExportMonitoring_ToMonthlyWorkbook( _
        CDate(Me.Controls("fraDailyLog").Controls("txtDailyDate").value), _
        Me.Controls("frHeader").Controls("txtHdrName").value, _
        Me.Controls("fraDailyLog").Controls("txtMonthlyMonitoringDraft").value)

        
End Sub

Private Sub mDailySave_Click()
    mDailyLogManual = True
    Call SaveDailyLog_Append(Me)
    mDailyLogManual = False
    MsgBox "日々の記録を保存しました。", vbInformation
End Sub





' 評価フォーム下部に「シートへ保存」ボタンを1つ配置する（1回実行用）
Public Sub PlaceGlobalSaveButton_Once()

    Dim btnClose As MSForms.CommandButton
    Dim btnSave As MSForms.CommandButton
    Dim c As MSForms.Control

    ' 「閉じる」ボタンをキャプションで特定
    For Each c In Me.Controls
        If TypeOf c Is MSForms.CommandButton Then
            If c.caption = "閉じる" Then
                Set btnClose = c
                Exit For
            End If
        End If
    Next c

    If btnClose Is Nothing Then
        MsgBox "閉じるボタンが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 既にグローバル保存ボタンがあるか確認
    On Error Resume Next
    Set btnSave = Me.Controls("cmdSaveGlobal")
    On Error GoTo 0

    ' なければ新規作成
    If btnSave Is Nothing Then
        Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "cmdSaveGlobal")
        btnSave.caption = "シートへ保存"
    End If

    ' 閉じるボタンと高さ・縦位置をそろえて、左隣に配置
    With btnSave
        .Height = btnClose.Height
        .Top = btnClose.Top
        .Width = btnClose.Width + 50
        .Left = btnClose.Left - .Width - 12
    End With

    ' ---- クリアボタン（cmdClearGlobal）を配置 ----
Dim btnClear As MSForms.CommandButton

' 既に存在するか確認
On Error Resume Next
Set btnClear = Me.Controls("cmdClearGlobal")
On Error GoTo 0

' なければ新規作成
If btnClear Is Nothing Then
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "cmdClearGlobal")
    btnClear.caption = "クリア"
End If

' 保存ボタンの右隣に配置
With btnClear
    .Height = btnSave.Height
    .Top = btnSave.Top
    .Width = btnSave.Width
    .Left = btnSave.Left + btnSave.Width + 12
End With

    
    If mGlobalClear Is Nothing Then
    Set mGlobalClear = New clsGlobalSaveButton
End If
Set mGlobalClear.btn = btnClear

    
    
    
    ' ---- ボタン整列（保存 → クリア → 閉じる） ----

' 閉じるボタンを基準として一番右端に固定
btnClose.Left = btnClose.Left

' 保存ボタンを閉じるの左に
btnSave.Left = btnClose.Left - btnSave.Width - 12

' クリアボタン（cmdClearGlobal）を保存の左に
btnClear.Left = btnSave.Left - btnClear.Width - 12




       If mGlobalSave Is Nothing Then
        Set mGlobalSave = New clsGlobalSaveButton
    End If
    Set mGlobalSave.btn = btnSave




End Sub



Private Sub mGlobalSave_Clicked()
    btnSaveCtl_Click
End Sub



Private Sub mGlobalClear_Clicked()
    Dim c As MSForms.Control

    For Each c In Me.Controls
        ' テキストボックスは空に
        If TypeOf c Is MSForms.TextBox Then
            ' 評価日(txtEDate)と日々記録の日付(txtDailyDate)はクリアしない
            If c.name <> "txtEDate" And c.name <> "txtDailyDate" Then
                c.value = ""
            End If
        End If

        ' コンボボックスは選択解除
        If TypeOf c Is MSForms.ComboBox Then
            c.value = ""
        End If

        ' チェックボックスはオフ
        If TypeOf c Is MSForms.CheckBox Then
            c.value = False
        End If

        ' リストボックスは選択だけ解除（項目は残す）
        If TypeOf c Is MSForms.ListBox Then
            Dim lb As MSForms.ListBox
            Dim i As Long

            Set lb = c
            lb.ListIndex = -1
            For i = 0 To lb.ListCount - 1
                lb.Selected(i) = False
            Next i
        End If
    Next c
End Sub




Private Sub mDailyList_DblClicked()
    Dim lb As MSForms.ListBox
    Dim r As Long, c As Long
    Dim buf As String
    
    ' 対象の一覧ListBoxを取得
    Set lb = Me.Controls("lstDailyLogList")
    
    ' 全行・全列をタブ区切り＋改行で連結
    For r = 0 To lb.ListCount - 1
        For c = 0 To lb.ColumnCount - 1
            If c > 0 Then buf = buf & vbTab
            buf = buf & CStr(lb.List(r, c))
        Next c
        buf = buf & vbCrLf
    Next r
    
    ' クリップボードへコピー
    Dim dobj As New MSForms.DataObject
    dobj.SetText buf
    dobj.PutInClipboard
    
    MsgBox "この月の記録一覧をクリップボードにコピーしました。" & vbCrLf & _
           "メモ帳やWordに Ctrl+V で貼り付けできます。", vbInformation
End Sub



Public Sub HookDailyLogList(lb As MSForms.ListBox)
    ' 日々の記録一覧 ListBox 用のイベントフック
    If mDailyList Is Nothing Then
        Set mDailyList = New clsDailyLogList
    End If
    Set mDailyList.lb = lb
End Sub



'=== 日々の記録フレーム取得ヘルパー（共通化用） ===
Private Function GetDailyLogFrame() As MSForms.Frame
    Dim mp As Object
    Dim pg As Object
    Dim f As Object
    Dim i As Long

    On Error Resume Next

    Set mp = Me.Controls("MultiPage1")
    If mp Is Nothing Then
        Exit Function
    End If

    ' 「日々の記録」ページを探す
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = "日々の記録" Then
            Set pg = mp.Pages(i)
            Exit For
        End If
    Next i
    If pg Is Nothing Then
        Exit Function
    End If

    ' フレーム fraDailyLog を取得
    Set f = pg.Controls("fraDailyLog")
    If f Is Nothing Then
        Exit Function
    End If

    Set GetDailyLogFrame = f
End Function



Private Function GetMainMultiPage() As MSForms.MultiPage
    On Error Resume Next
    Set GetMainMultiPage = Me.Controls("MultiPage1")
End Function



'=== 歩行評価フレーム取得ヘルパー（Frame6 固定） ===
Private Function GetWalkFrame() As MSForms.Frame
    On Error Resume Next

    ' 名前で直接取得（命名前提：Frame6）
    Set GetWalkFrame = Me.Controls("Frame6")

    ' もし見つからなければ何も返さない（Nothing）
End Function








Private Sub FixWalkRootFrameHeight()
    Dim f As MSForms.Frame
    Dim c As Control
    Dim maxBottom As Single

    Set f = GetWalkFrame()
    If f Is Nothing Then Exit Sub

    ' 子コントロールの一番下の位置を調べる
    For Each c In f.Controls
        If c.Top + c.Height > maxBottom Then
            maxBottom = c.Top + c.Height
        End If
    Next c

    ' 必要なら高さを伸ばす
    If maxBottom + 6 > f.Height Then
#If APP_DEBUG Then
        Debug.Print "[WALK-ROOT-RESIZE]", _
                    "oldH=", f.Height, _
                    "maxBottom=", maxBottom, _
                    "newH=", maxBottom + 6
#End If
        f.Height = maxBottom + 6
    End If
End Sub








Public Sub Debug_FixOverflowFrames()

Debug.Print "[CALL] Debug_FixOverflowFrames @ " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    On Error Resume Next

    ' 姿勢評価タブ
    FitFrameHeightToChildren Me.Controls("Frame2")

    ' 身体機能評価タブ（親だけ調整。子Frame12には触らない）
    FitFrameHeightToChildren Me.Controls("Frame12")
    FitFrameHeightToChildren Me.Controls("Frame3")
    ' FitFrameHeightToChildren Me.Controls("Frame14")

    ' 歩行評価タブ（大枠）
    FitFrameHeightToChildren Me.Controls("Frame6")

    ' 認知・精神タブ（親Frame7だけを調整）
    FitFrameHeightToChildren Me.Controls("Frame7")

    On Error GoTo 0
End Sub










Private Sub GetPageUsableArea( _
    ByVal pageIndex As Long, _
    ByRef x As Single, _
    ByRef Y As Single, _
    ByRef w As Single, _
    ByRef h As Single)

    Dim mp As MSForms.MultiPage

    x = 0: Y = 0: w = 0: h = 0   ' デフォルトクリア

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Sub

    If pageIndex < 0 Then Exit Sub
    If pageIndex > mp.Pages.Count - 1 Then Exit Sub

    ' 今は MultiPage 全体を「ページの利用可能領域」として返す
    ' （余白やタブ分のマイナスは、後で AlignRootFrame 側で調整する）
    x = 0
    Y = 0
    w = mp.Width
    h = mp.Height
End Sub



Private Sub AlignRootFrameToPage(ByVal pageIndex As Long, root As MSForms.Frame)
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim pageLeft As Single, pageTop As Single
    Dim pageWidth As Single, pageHeight As Single

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Sub
    If root Is Nothing Then Exit Sub
    If pageIndex < 0 Or pageIndex > mp.Pages.Count - 1 Then Exit Sub

    Set pg = mp.Pages(pageIndex)

    '=== ページのクライアント領域（タブを除いた中身部分）を算出 ===
    ' ここは PREVIEW で出ていた値と同じロジックに揃える前提で、
    ' シンプルに MultiPage の内側を使う
    pageLeft = 0
    pageTop = 0
    pageWidth = mp.Width          ' タブ左右の余白はほぼ 0 扱い
    pageHeight = mp.Height - 40   ' 下側ボタンぶん少しだけ控えめ

    '=== ルートフレームをページ一杯にフィット ===
    With root
        .Left = pageLeft
        .Top = pageTop
        .Width = pageWidth
        .Height = pageHeight
    End With
End Sub




Private Sub PreviewOnePage(ByVal idx As Long, ByVal mp As MSForms.MultiPage)
    Dim pg As MSForms.Page
    Dim root As MSForms.Frame
    Dim x As Single, Y As Single, w As Single, h As Single

    Set pg = mp.Pages(idx)
    Set root = GetPageRootFrame(idx)

    Debug.Print "--- Page", idx, "[" & pg.caption & "] ---"

    If root Is Nothing Then
        Debug.Print "  RootFrame: <NOT FOUND>"
        Exit Sub
    End If

    ' 現在値
    Debug.Print "  Current:", _
                "L=" & root.Left, _
                "T=" & root.Top, _
                "W=" & root.Width, _
                "H=" & root.Height

    ' AlignRootFrameToPage が使うページ領域
    GetPageUsableArea idx, x, Y, w, h
    Debug.Print "  PageArea:", _
                "X=" & x, "Y=" & Y, _
                "W=" & w, "H=" & h

    ' もし AlignRootFrameToPage を呼んだらこうなる（※実際には書き換えない）
    Debug.Print "  WouldAlignTo:", _
                "L=" & (x), _
                "T=" & (Y), _
                "W=" & (w), _
                "H=" & (h)
End Sub


Private Function GetPageRootFrame(ByVal pageIndex As Long) As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.Page
    Dim c As Control
    Dim f As MSForms.Frame
    Dim best As MSForms.Frame
    Dim bestArea As Single
    Dim area As Single

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Function

    If pageIndex < 0 Or pageIndex > mp.Pages.Count - 1 Then Exit Function

    Set pg = mp.Pages(pageIndex)
    If pg.caption = "認知・精神" Then
#If APP_DEBUG Then
    Debug.Print "[HIT] GetPageRootFrame COG -> Frame31"
#End If
    Set GetPageRootFrame = pg.Controls("Frame7").Controls("Frame31")
    Exit Function
End If

    ' そのページ内の「一番大きな Frame = ルートフレーム」とみなす
    For Each c In pg.Controls
        If TypeName(c) = "Frame" And TypeName(c.parent) = "Page" Then
            Set f = c
            area = f.Width * f.Height
            If best Is Nothing Or area > bestArea Then
                Set best = f
                bestArea = area
            End If
        End If
    Next

    Set GetPageRootFrame = best
End Function



Public Sub Apply_AlignRoot_All()
    Dim mp As MSForms.MultiPage
    Dim i As Long
    Dim root As MSForms.Frame

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Sub

    For i = 0 To mp.Pages.Count - 1
        Set root = GetPageRootFrame(i)
        If Not root Is Nothing Then
            AlignRootFrameToPage i, root
        End If
    Next i
End Sub


 
Public Sub TidyBaseLayout_Once()
    If mBaseLayoutDone Then Exit Sub
    mBaseLayoutDone = True

    '★ 基本レイアウト（ページ共通）はここだけでやる
    Apply_AlignRoot_All


End Sub




Public Sub SetFormHeightSafe(ByVal newH As Single)
    Me.Height = newH
    DoEvents
    Me.Controls("MultiPage1").Height = Me.InsideHeight - 12
    If mBaseLayoutDone Then
        Apply_AlignRoot_All
    End If
End Sub




Public Sub AdjustBottomButtons()

    Dim yBtn As Single

    ' ボタンがまだ無いタイミングでは何もしない
    If Not ControlExists(Me, "btnCloseCtl") Then Exit Sub
    If Not ControlExists(Me, "cmdSaveGlobal") Then Exit Sub
    If Not ControlExists(Me, "cmdClearGlobal") Then Exit Sub

    yBtn = Me.InsideHeight - Me.Controls("btnCloseCtl").Height - 12

    Me.Controls("btnCloseCtl").Top = yBtn
    Me.Controls("cmdSaveGlobal").Top = yBtn
    Me.Controls("cmdClearGlobal").Top = yBtn
    
    
     ' ★ここ（前面へ）
    Me.Controls("btnCloseCtl").ZOrder 0
    Me.Controls("cmdSaveGlobal").ZOrder 0
    Me.Controls("cmdClearGlobal").ZOrder 0
    
End Sub





Public Sub BuildEvalShell_Once()

    '
    ' Shell authority (final winner): BuildEvalShell_Once
    ' LegacyInit/FitLayout do NOT own shell layout
    ' No Controls.Add MultiPage. Use existing MultiPage1
    ' Auto-run: Initialize only. Activate is no-op
    '
    If mLayoutBuilt Then Exit Sub
    mLayoutBuilt = True
    
    
Dim minH As Single
Dim maxH As Single


minH = 620   ' ← 評価フォームとして最低限欲しい高さ（調整可）

maxH = Application.UsableHeight - (Me.Height - Me.InsideHeight) - 6

If maxH < minH Then
    ' 画面が小さすぎる場合は、最低サイズを優先
    Me.Height = minH
Else
    Me.Height = maxH
    If Me.Height < minH Then Me.Height = minH
End If



    Dim frHeader As MSForms.Frame
    Dim frViewport As MSForms.Frame
    Dim mp As MSForms.MultiPage

    '--- Header（操作バー：固定・非スクロール）---
    On Error Resume Next
    Set frHeader = Me.Controls("frHeader")
    On Error GoTo 0
    If frHeader Is Nothing Then
        Set frHeader = Me.Controls.Add("Forms.Frame.1", "frHeader", True)
        frHeader.caption = vbNullString
        frHeader.SpecialEffect = fmSpecialEffectFlat
    End If

    With frHeader
        .Left = PAD_SIDE
        .Top = PAD_SIDE
        .Width = Me.InsideWidth - PAD_SIDE * 2
        .Height = HEADER_H
        .ZOrder 0
        .Visible = True
    End With

    '--- 既存のメイン MultiPage（探すだけ。作らない）---
    Set mp = FindMainMultiPage()

    If Not mp Is Nothing Then
        mp.Left = PAD_SIDE
        mp.Top = frHeader.Top + frHeader.Height + GAP_V
        mp.Width = Me.InsideWidth - PAD_SIDE * 2
       mp.Height = Application.Max(120, (maxH - (Me.Height - Me.InsideHeight)) - mp.Top - PAD_SIDE)

        ' 高さは後段のViewportで決める
    End If

    '--- Viewport（評価コンテンツ専用：必要ならスクロール）---
    On Error Resume Next
    Set frViewport = Me.Controls("frViewport")
    On Error GoTo 0
    If frViewport Is Nothing Then
        Set frViewport = Me.Controls.Add("Forms.Frame.1", "frViewport", True)
        frViewport.caption = vbNullString
        frViewport.SpecialEffect = fmSpecialEffectFlat
    End If

    With frViewport
        .Left = PAD_SIDE
        If Not mp Is Nothing Then
            .Top = mp.Top + mp.Height + GAP_V
        Else
            .Top = frHeader.Top + frHeader.Height + GAP_V
        End If
        .Width = Me.InsideWidth - PAD_SIDE * 2
        Dim hV As Single
hV = Me.InsideHeight - .Top - PAD_SIDE


If hV < 0 Then hV = 0
.Height = hV

        .ScrollBars = fmScrollBarsNone
        .Visible = True
    End With
    
    
    Dim maxB As Single
maxB = 0




End Sub




Public Sub MoveGlobalButtonsToHeader_Once()
    Static done As Boolean
    If done Then Exit Sub
    done = True

    On Error GoTo EH
    Dim stepN As String

10  stepN = "GetHeader": Dim f As MSForms.Frame: Set f = Me.Controls("frHeader")

20  stepN = "GetButtons"
    Dim bClear As MSForms.Control, bSave As MSForms.Control, bClose As MSForms.Control
    Set bClear = Me.Controls("cmdClearGlobal")
    Set bSave = Me.Controls("cmdSaveGlobal")
    Set bClose = Me.Controls("btnCloseCtl")

30  stepN = "SetParent"
    bClear.parent = f
    bSave.parent = f
    bClose.parent = f

40  stepN = "Align"
    Const pad As Single = 8, gap As Single = 10
    bClose.Top = (f.Height - bClose.Height) / 2
    bSave.Top = bClose.Top
    bClear.Top = bClose.Top
    bClose.Left = f.Width - pad - bClose.Width
    bSave.Left = bClose.Left - gap - bSave.Width
    bClear.Left = bSave.Left - gap - bClear.Width

    Exit Sub
EH:
    Debug.Print "[MoveButtons][ERR] step=" & stepN & " Err=" & Err.Number & " " & Err.Description
End Sub






Public Sub CreateHeaderButtons_Once()
    Static done As Boolean
    If done Then Exit Sub
    done = True

    Dim f As MSForms.Frame
    Set f = Me.Controls("frHeader")

    ' 既存ボタン（処理の本体）
    Dim bClear As MSForms.Control, bSave As MSForms.Control, bClose As MSForms.Control
    Set bClear = Me.Controls("cmdClearGlobal")
    Set bSave = Me.Controls("cmdSaveGlobal")
    Set bClose = Me.Controls("btnCloseCtl")

    ' 既存は見えなくする（位置は触らない）
    bClear.Visible = False
    bSave.Visible = False
    bClose.Visible = False

    ' ヘッダー用の新ボタンを作る（名前固定）
    Dim hClear As MSForms.CommandButton
    Dim hSave  As MSForms.CommandButton
    Dim hClose As MSForms.CommandButton

    Set hClear = f.Controls.Add("Forms.CommandButton.1", "cmdClearHeader", True)
    Set hSave = f.Controls.Add("Forms.CommandButton.1", "cmdSaveHeader", True)
    Set hClose = f.Controls.Add("Forms.CommandButton.1", "cmdCloseHeader", True)

    ' 見た目は既存を踏襲
    hClear.caption = bClear.caption: hClear.Width = bClear.Width: hClear.Height = bClear.Height
    hSave.caption = bSave.caption:   hSave.Width = bSave.Width:   hSave.Height = bSave.Height
    hClose.caption = bClose.caption: hClose.Width = bClose.Width: hClose.Height = bClose.Height

    ' 右寄せ配置
    Const pad As Single = 8, gap As Single = 10
    Dim hdrBtnTop As Single
    hdrBtnTop = (44 - hClose.Height) / 2

    hClose.Top = hdrBtnTop
    hSave.Top = hClose.Top
    hClear.Top = hClose.Top

    hClose.Left = f.Width - pad - hClose.Width
    hSave.Left = hClose.Left - gap - hSave.Width
    hClear.Left = hSave.Left - gap - hClear.Width
    
    Set mHdr1 = New clsHeaderBtnEvents
Set mHdr1.btn = hSave
mHdr1.tag = "Save"

Set mHdr2 = New clsHeaderBtnEvents
Set mHdr2.btn = hClear
mHdr2.tag = "Clear"

Set mHdr3 = New clsHeaderBtnEvents
Set mHdr3.btn = hClose
mHdr3.tag = "Close"


'==============================
' LoadPrev（前回の値を読み込む）ヘッダーボタン + Hook
'==============================
Dim hLoadPrev As MSForms.CommandButton


' 既にあればそれを掴む（＝インスタンスを増やさない）
On Error Resume Next
Set hLoadPrev = f.Controls("cmdHdrLoadPrev")
On Error GoTo 0

' 無ければ作る
If hLoadPrev Is Nothing Then
    Set hLoadPrev = f.Controls.Add("Forms.CommandButton.1", "cmdHdrLoadPrev", True)
End If

hLoadPrev.caption = "前回の値を読み込む"
hLoadPrev.Width = 180
hLoadPrev.Height = 24
hLoadPrev.Top = hClose.Top

' 位置：txtHdrKana の右（txtHdrKana が無い場合は右端の左に置く）
On Error Resume Next
Dim tbKana As MSForms.Control
Set tbKana = f.Controls("txtHdrKana")
On Error GoTo 0

If Not tbKana Is Nothing Then
    hLoadPrev.Left = tbKana.Left + tbKana.Width + 12
Else
    hLoadPrev.Left = hClear.Left - 12 - hLoadPrev.Width
End If

' Hook（クリックで既存の btnLoadPrevCtl_Click へ流す）
Set mHdrLoadPrevHook = New clsHdrBtnHook
Set mHdrLoadPrevHook.btn = hLoadPrev
mHdrLoadPrevHook.tag = "LoadPrev"
Set mHdrLoadPrevHook.owner = Me

' 旧ボタンは非表示
On Error Resume Next
Me.Controls("MultiPage1").Pages(0).Controls("Frame32").Controls("btnLoadPrevCtl").Visible = False
On Error GoTo 0

    
End Sub




Public Sub DoSaveGlobal()
    Call btnSaveCtl_Click
End Sub


Public Sub DoClearGlobal()
    Call mGlobalClear_Clicked
End Sub


Public Sub DoCloseForm()
    Call btnCloseCtl_Click
End Sub


Private Sub Tidy_DailyLog_Once()
    Static done As Boolean
    If done Then Exit Sub
    done = True

    Tighten_DailyLog_Boxes
End Sub





Private Sub ApplyScroll_MP1_Page3_7_Once()
    If mScrollOnce_347 Then Exit Sub
    mScrollOnce_347 = True

   

    Dim mp As Object
    Set mp = Me.Controls("MultiPage1")

    'Page3: Frame3（ScrollHeight = 578.35 + 24 = 602.35）
    With mp.Pages(2).Controls("Frame3")
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 900
    End With

    'Page7: Frame7（必要時のみバー表示）
With mp.Pages(6).Controls("Frame7")
    .ScrollHeight = 584.35
    If .ScrollHeight > .Height Then
        .ScrollBars = fmScrollBarsVertical
    Else
        .ScrollBars = fmScrollBarsNone
    End If
End With



'Page2: Frame2（姿勢評価の下見切れ対策）
With mp.Pages(1).Controls("Frame2")
    .Height = mp.Height
    .ScrollBars = fmScrollBarsVertical
    .ScrollHeight = 488   ' 464 + 24
End With




'Page1: Frame1（小画面で下が見切れる対策）
With mp.Pages(0).Controls("Frame1")
    .Height = mp.Height
    .ScrollBars = fmScrollBarsVertical
    .ScrollHeight = 420   '←いったん安全値（後で調整可）
End With




End Sub

Private Sub RequestQuitExcelAskAndCloseForm()
    mQuitMode = qmAsk
    Unload Me
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

   Application.OnKey "^+D"

    

    '右上の小さい×：フォームだけ閉じる（Excelは閉じない）
    If CloseMode = vbFormControlMenu Then
        mQuitMode = qmNone
        Exit Sub
    End If

    '閉じるボタン経由のみ：保存確認 → Excel終了
    If mQuitMode = qmAsk Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("保存しますか？", vbYesNoCancel + vbQuestion, "終了確認")

        If ans = vbCancel Then
            Cancel = True
            mQuitMode = qmNone
            Exit Sub
        End If

        On Error Resume Next
        Application.DisplayAlerts = False

        If ans = vbYes Then
            ThisWorkbook.Save
        Else
            '保存せず終了：Excelの保存確認を出さない
            ThisWorkbook.Saved = True
        End If

        Application.Quit
        On Error GoTo 0
    End If

End Sub




Private Sub frHeader_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    On Error Resume Next
    If TypeName(Me.ActiveControl) = "CommandButton" Then
        If Me.ActiveControl.name = "cmdArchiveDelete" Then
            ArchiveAndDelete_EvalData_ByName
        End If
    End If
    On Error GoTo 0
End Sub




Public Sub AddPrintButton_TestEval()
    Dim f As MSForms.Frame
    Set f = Me.Controls("Frame23")

    Dim btn As MSForms.CommandButton
    On Error Resume Next
    Set btn = f.Controls("cmdPrintTestEval")
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = f.Controls.Add("Forms.CommandButton.1", "cmdPrintTestEval", True)
        With btn
            .caption = "グラフ印刷"
            btn.Width = 120
btn.Height = 28
btn.Left = f.InsideWidth - btn.Width - 28.35
btn.Top = f.Controls("txtTUG").Top

        End With
    End If
    
    
btn.Left = f.InsideWidth - btn.Width - 28.35


    ' ← ★ここ★（この2行だけ追加）
    Set mPrintBtnHook = New clsPrintBtnHook
    Set mPrintBtnHook.btn = btn
    
    
    
    
End Sub


Public Sub BuildMonthlyDraft_FromDailyLog()
    Dim ws As Worksheet
    Dim f As Object
    Dim nm As String
    Dim v As Variant
    Dim dFrom As Date, dTo As Date
    Dim lastRow As Long, r As Long
    Dim s As String
    Dim hit As Long
    Dim d As Date, staff As String, note As String

    Set f = Me.Controls("fraDailyLog")

    ' 対象月＝記録日（txtDailyDate）の月
    v = f.Controls("txtDailyDate").value
    If Not IsDate(v) Then
        MsgBox "記録日の欄に正しい日付を入力してください。", vbExclamation
        Exit Sub
    End If
    dFrom = DateSerial(Year(CDate(v)), Month(CDate(v)), 1)
    dTo = DateSerial(Year(CDate(v)), Month(CDate(v)) + 1, 0)

    ' 対象者＝フォーム氏名（DailyLogのB列と一致）
    nm = Trim$(Me.Controls("frHeader").Controls("txtHdrName").value)
    If nm = "" Then
        MsgBox "氏名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Dim pid As String, cntSameName As Long
    

    Set ws = ThisWorkbook.Worksheets("DailyLog")
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    
    pid = Trim$(Me.Controls("frHeader").Controls("txtHdrPID").value)
    cntSameName = Application.WorksheetFunction.CountIf(ws.Range("B:B"), nm)


    s = "【月次モニタリング下書き】" & vbCrLf _
      & "対象：" & nm & vbCrLf _
      & "期間：" & Format$(dFrom, "yyyy/mm/dd") & " - " & Format$(dTo, "yyyy/mm/dd") & vbCrLf & vbCrLf _
      & "■ この月に記録された特記事項（時系列）" & vbCrLf

    hit = 0
    For r = 2 To lastRow
        If Trim$(ws.Cells(r, 2).value) = nm And (cntSameName = 1 Or Trim$(ws.Cells(r, 3).value) = pid) Then
            If IsDate(ws.Cells(r, 1).value) Then
                d = CDate(ws.Cells(r, 1).value)
                If d >= dFrom And d <= dTo Then
                    note = CStr(ws.Cells(r, 5).value)
                    If Len(Trim$(note)) > 0 Then
                        staff = CStr(ws.Cells(r, 4).value)
                        s = s & "・" & Format$(d, "m/d") & "（" & staff & "） " & note & vbCrLf
                        hit = hit + 1
                    End If
                End If
            End If
        End If
    Next r

    If hit = 0 Then
        s = s & "・（この月の記録はありません）" & vbCrLf
    End If

    ' 出力先（起動時に確保済みだが念のため）
    Call Ensure_MonthlyDraftBox_UnderFraDailyLog
    Me.Controls("fraDailyLog").Controls("txtMonthlyMonitoringDraft").value = s
End Sub





Public Sub EnsureNameSuggestList()
    Dim host As Object
    Dim tb As MSForms.TextBox
    Dim lb As MSForms.ListBox

    Set host = Me
    Set tb = host.Controls("txtHdrName")


    On Error Resume Next
    Set lb = Me.Controls("lstNameSuggest")
    On Error GoTo 0


    If lb Is Nothing Then
        Set lb = host.Controls.Add("Forms.ListBox.1", "lstNameSuggest", True)
    End If

    With lb
        .Left = Me.Controls("frHeader").Left + tb.Left
        .Top = Me.Controls("frHeader").Top + tb.Top + tb.Height + 4
        .Width = 200     ' 横をコンパクトに
        .Height = 60     ' 3件くらい見える高さ

        .Visible = False
    End With
    
      Set mNameSuggestSink = New cNameSuggestSink
      mNameSuggestSink.Hook Me.Controls("lstNameSuggest")
    
End Sub



Private Sub txtHdrName_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' ヘッダ入力 → 裏口へ同期（既存ロジック駆動のため）
    Me.Controls("txtName").Text = Me.Controls("frHeader").Controls("txtHdrName").Text

    ' 候補BOX確保 → 更新
    EnsureNameSuggestList
    Me.UpdateNameSuggest
End Sub




Private Sub lstNameSuggest_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim lb As MSForms.ListBox
    Set lb = Me.Controls("lstNameSuggest")

    If lb.ListIndex < 0 Then Exit Sub

    ' ヘッダの氏名だけ反映
    Me.Controls("frHeader").Controls("txtHdrName").Text = lb.List(lb.ListIndex, 0)

    ' 裏口同期（既存ロジック用）
    Me.Controls("txtName").Text = Me.Controls("frHeader").Controls("txtHdrName").Text

    lb.Visible = False
End Sub



'====================================================
' 前回の値を読み込むボタン：作成＆配置（On Errorなし）
'====================================================
Private Function TryGetCtl(ByVal container As Object, ByVal ctlName As String, ByRef outCtl As Object) As Boolean
    Dim c As Object
    For Each c In container.Controls
        If StrComp(c.name, ctlName, vbTextCompare) = 0 Then
            Set outCtl = c
            TryGetCtl = True
            Exit Function
        End If
    Next
    TryGetCtl = False
End Function

Public Sub Ensure_LoadPrevButton_Once(ByVal f As Object)
    Const BTN_NAME As String = "cmdHdrLoadPrev"

    Dim hdr As Object, kana As Object, btn As Object
    Dim refBtn As Object

    ' frHeader 必須
    If Not TryGetCtl(f, "frHeader", hdr) Then Exit Sub

    ' 参照（配置の基準）: かな欄必須
    If Not TryGetCtl(hdr, "txtHdrKana", kana) Then Exit Sub

    ' 既にあればそれを使う
    If TryGetCtl(hdr, BTN_NAME, btn) Then
        ' そのまま配置だけ当て直し
    Else
        ' 無ければ作る（frHeader配下）
        Set btn = hdr.Controls.Add("Forms.CommandButton.1", BTN_NAME, True)
        btn.caption = "前回の値を読み込む"
        btn.Accelerator = "L"
        btn.Width = 180
        btn.Height = 24
    End If

    ' 見た目合わせ用の参照ボタン（あれば）
    If Not TryGetCtl(hdr, "cmdSaveHeader", refBtn) Then
        If Not TryGetCtl(hdr, "cmdClearHeader", refBtn) Then
            Call TryGetCtl(hdr, "cmdCloseHeader", refBtn)
        End If
    End If
    If Not refBtn Is Nothing Then
        btn.Font.name = refBtn.Font.name
        btn.Font.Size = refBtn.Font.Size
       
    End If

    ' 位置：txtHdrKana の右
    btn.Left = kana.Left + kana.Width + 12
    btn.Top = kana.Top + 2
End Sub

