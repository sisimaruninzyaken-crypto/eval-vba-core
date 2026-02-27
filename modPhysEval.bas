Attribute VB_Name = "modPhysEval"
' ===== modPhysEval.bas =====
Option Explicit



'=== レイアウト共通 ===
Private Const PAD_X      As Single = 12   ' ← すでに使っている全体余白
Private Const PAD_Y      As Single = 10
Private Const GAP_Y      As Single = 6
Private Const NOTE_H     As Single = 22
Private Const COL1_W     As Single = 140  ' 左の見出し列幅（他画面でも使用）
Private Const COL_EDT_W  As Single = 70   ' 他画面用の入力幅
Private Const ROW_H     As Single = 18

' === ROM 子タブ用 追加定数／フラグ ===
Private Const ROM_JOINT_GAP_Y As Single = 12   ' 関節ブロック間の縦間
Private Const ROM_MOTION_GAP_Y As Single = 4   ' 運動行の縦間
Private Const ROM_GROUP_PAD    As Single = 8   ' Frame 内パディング
Private Const ROM_HDR_RL_GAP   As Single = 24  ' 運動名とR/L列の間

Public Const USE_ROM_SUBTABS   As Boolean = True   ' 子タブ(上肢/下肢)を使う


'=== ROMレイアウト（この4つはROM専用で1回だけ定義）===
Private Const ROM_ROW_H      As Single = 14
Private Const ROM_GAP_Y      As Single = 1
Private Const ROM_HDR_GAP    As Single = 2
Private Const ROM_COL_EDT_W  As Single = 38

'=== 備考欄の共通パラメータ（新規） ===
Private Const MEMO_DESIRED_H As Single = 120   ' ほしい高さ（100?160で好みに調整可）
Private Const MEMO_MIN_H     As Single = 72    ' 最低高さ

'=== 備考/変形テキスト用 定数（modPhysEval内で共通） ===
Private Const NOTE_W_RATE    As Single = 0.6   ' 備考ボックスの横幅 = 利用可能幅の60%
Private Const DEFORM_W_RATE  As Single = 0.6   ' 変形テキストの横幅 = 利用可能幅の60%
Private Const DEFORM_H       As Single = 120   ' 変形テキストの高さ（px相当）

Private Const CAP_FUNC_REFLEX As String = "筋緊張・反射（痙縮含む）"
Private Const CAP_FUNC_PAIN   As String = "疼痛（部位／NRS）"





Public Sub PlaceMemoBelow( _
    host As MSForms.Frame, _
    ByVal w As Single, ByVal h As Single, _
    ByVal yTop As Single, _
    ByVal memoName As String, _
    Optional ByVal fr1 As MSForms.Frame, _
    Optional ByVal fr2 As MSForms.Frame, _
    Optional ByVal labelText As String = "備考欄")

    Const GAP_X As Single = 8
    Const GAP_Y As Single = 4

    ' 既存のメモとラベルを除去  ← ここ、半角の " に！
    On Error Resume Next
    If ControlExists(host, memoName) Then host.Controls.Remove memoName

    Dim i As Long
    For i = host.Controls.Count - 1 To 0 Step -1
        If host.Controls(i).name = memoName & "_lbl" Then
        host.Controls.Remove i
        Exit For
    End If
Next i

    On Error GoTo 0
' --- 下端の基準を決める（fr1/fr2 の深い方 + PAD_Y）---
Dim yBottom As Single
yBottom = yTop
If Not fr1 Is Nothing Then yBottom = Application.WorksheetFunction.Max(yBottom, fr1.Top + fr1.Height + PAD_Y)
If Not fr2 Is Nothing Then yBottom = Application.WorksheetFunction.Max(yBottom, fr2.Top + fr2.Height + PAD_Y)

' メモ領域の開始位置（ラベルのTop）を決定
Dim memoTop As Single, safeTopMax As Single
memoTop = yBottom
If memoTop < yBottom Then memoTop = yBottom

' ラベルTopの最大許容（= 残りが ROW_H + GAP_Y + MEMO_MIN_H は確保できる位置）
safeTopMax = h - PAD_Y - (ROW_H + GAP_Y + MEMO_MIN_H)
If safeTopMax < PAD_Y Then safeTopMax = PAD_Y     ' フレームが極端に低い場合の保険
If memoTop > safeTopMax Then memoTop = safeTopMax

' 見出しラベル
Dim lbl As MSForms.label
Set lbl = host.Controls.Add("Forms.Label.1", memoName & "_lbl")
With lbl
    .caption = labelText
    .Left = PAD_X
    .Top = memoTop
    .Width = w - PAD_X * 2
    .Height = ROW_H
    .Font.Bold = False
End With


' 固定下寄せはやめる：備考欄は「評価項目の直下(yBottom)」に置き、縦サイズは MEMO_DESIRED_H を上限にして下に伸びすぎないようにする（2026-01）
' テキストボックス本体
Dim txt As MSForms.TextBox, hCalc As Single
Set txt = host.Controls.Add("Forms.TextBox.1", memoName)
With txt
    .Left = PAD_X
    .Top = lbl.Top + ROW_H
    .Width = w - PAD_X * 2

    ' 残り高さを計算 → 最低高さ＆1以上に丸めてから設定
    hCalc = Application.WorksheetFunction.Min(MEMO_DESIRED_H, h - PAD_Y - .Top)
    If hCalc < MEMO_MIN_H Then hCalc = MEMO_MIN_H
    If hCalc < 1 Then hCalc = 1
    .Height = PX(hCalc)

    .multiline = True
    .WordWrap = True
    .EnterKeyBehavior = True
    .ScrollBars = fmScrollBarsVertical
End With

' ラベルの直前で左右カラムのフレームを止める
If Not fr1 Is Nothing Then fr1.Height = lbl.Top - PAD_Y
If Not fr2 Is Nothing Then fr2.Height = lbl.Top - PAD_Y

' このページではスクロール禁止
host.ScrollBars = fmScrollBarsNone
host.ScrollHeight = host.Height

End Sub





'========================================
' 公開API：身体機能評価タブ一式を作成
'========================================
Public Sub EnsurePhysicalFunctionTabs(owner As frmEval)
    Dim mp As MSForms.MultiPage: Set mp = EnsurePhysMulti(owner)
    If mp Is Nothing Then Exit Sub

    Dim pgRom As MSForms.Page, pgMMT As MSForms.Page, pgSens As MSForms.Page, _
    pgReflex As MSForms.Page, pgPain As MSForms.Page, pgNote As MSForms.Page

    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)
    Set pgSens = FindOrAddPage(mp, CAP_FUNC_SENS_REF)
    Set pgNote = FindOrAddPage(mp, CAP_FUNC_NOTE)
    Set pgReflex = FindOrAddPage(mp, CAP_FUNC_REFLEX)
    Set pgPain = FindOrAddPage(mp, CAP_FUNC_PAIN)
    

    ' 各ページにホストフレームを用意
    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame, hostSens As MSForms.Frame, _
    hostReflex As MSForms.Frame, hostPain As MSForms.Frame, hostNote As MSForms.Frame

    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)
    Set hostReflex = EnsureHostFrame(pgReflex)   ' ← 追加
    Set hostPain = EnsureHostFrame(pgPain)       ' ← 追加
    
    Set hostNote = EnsureHostFrame(pgNote)

Dim pgPar As MSForms.Page, hostPar As MSForms.Frame
Set pgPar = FindOrAddPage(mp, CAP_FUNC_PARALYSIS)
Set hostPar = EnsureHostFrame(pgPar)
BuildParalysisTabUI hostPar   ' ← 既に貼った麻痺UIビルダ




    ' ビルド（UI生成）
    If USE_ROM_SUBTABS Then
    BuildROMTabs hostRom         ' ← 新：上肢／下肢の子タブ
Else
    BuildROMSection_Compact hostRom   ' ← 既存：二列レイアウト（互換用）
End If

    BuildMMTSection owner, hostMmt
   BuildSensoryTabUI hostSens
BuildToneReflexTabUI hostReflex
BuildPainTabUI owner, hostPain

    AddNotesBox owner, hostNote, TAG_FUNC_PREFIX

    ' 初期表示はROM
    mp.value = pgRom.Index
End Sub

'========================================
' 内部：MultiPageの用意（hostBody内）
'========================================
Private Function EnsurePhysMulti(owner As frmEval) As MSForms.MultiPage
    Dim host As MSForms.Frame: Set host = FindHostByName(owner, HOST_BODY_NAME)
    If host Is Nothing Then
        MsgBox "フレーム '" & HOST_BODY_NAME & "' が見つかりません。先に Validate_App を確認してください。", vbExclamation
        Exit Function
    End If

    Dim c As Control
    For Each c In host.Controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then
                Set EnsurePhysMulti = c
                Exit Function
            End If
        End If
    Next

    Set EnsurePhysMulti = host.Controls.Add("Forms.MultiPage.1")
    With EnsurePhysMulti
        .name = MP_PHYS_NAME
        .Left = PAD_X
        .Top = PAD_Y
        .Width = host.Width - PAD_X * 2
        .Height = host.Height - PAD_Y * 2
    End With

    ' タブ切替フック（将来のIME再適用等に備え）
    On Error Resume Next
    Dim mph As New MPHook
    mph.Init owner, EnsurePhysMulti
    owner.RegisterMPHook mph
    On Error GoTo 0
End Function

Private Function FindHostByName(frm As frmEval, hostName As String) As MSForms.Frame
    Dim c As Control
    For Each c In frm.Controls
        If TypeOf c Is MSForms.Frame Then
            If c.name = hostName Then Set FindHostByName = c: Exit Function
        End If
    Next
End Function

Private Function FindOrAddPage(mp As MSForms.MultiPage, captionText As String) As MSForms.Page
    Dim i As Long
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = captionText Then
            Set FindOrAddPage = mp.Pages(i)
            Exit Function
        End If
    Next
    Set FindOrAddPage = mp.Pages.Add
    FindOrAddPage.caption = captionText
End Function

Public Function EnsureHostFrame(pg As MSForms.Page) As MSForms.Frame

    Dim c As Control
    For Each c In pg.Controls
        If TypeOf c Is MSForms.Frame Then
            Set EnsureHostFrame = c
            Exit Function
        End If
    Next

    ' ★ ここから下は「初回のみ」
    Dim f As MSForms.Frame
    Set f = pg.Controls.Add("Forms.Frame.1")
    With f
    .caption = ""
    .Left = PAD_X
    .Top = PAD_Y
    .Width = pg.parent.Width - PAD_X * 2
    f.Height = pg.parent.Height
    f.ScrollBars = fmScrollBarsVertical
    .tag = "HOST_FIXED"
End With

    Set EnsureHostFrame = f
End Function








'========================================
' MMT（主要筋群・左右）
'========================================
Private Sub BuildMMTSection(owner As frmEval, host As MSForms.Frame)
    Dim Y As Single: Y = PAD_Y

    Dim groups As Variant
    ' 代表筋群（簡潔）：肩外転／肘屈曲／手関節背屈／股屈曲／膝伸展／足背屈
    groups = Array("肩外転", "肘屈曲", "手関節背屈", "股屈曲", "膝伸展", "足背屈")

    ' 見出し
    Y = AddHeaderRow(host, "筋群", Y)

    Dim i As Long
    For i = LBound(groups) To UBound(groups)
        Y = AddMMTRow(owner, host, CStr(groups(i)), Y)
    Next i
    
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, Y + ROM_HDR_GAP, "txtMMTMemo")
    
End Sub

Private Function AddHeaderRow(host As MSForms.Frame, title As String, Y As Single) As Single
    Dim hTitle As MSForms.label, hR As MSForms.label, hL As MSForms.label
    Set hTitle = host.Controls.Add("Forms.Label.1")
    With hTitle: .caption = title: .Left = PAD_X: .Top = Y: .Width = COL1_W: .Height = ROW_H: .Font.Bold = True: End With
    Set hR = host.Controls.Add("Forms.Label.1")
    With hR: .caption = "右": .Left = PAD_X + COL1_W + 8: .Top = Y: .Width = COL_EDT_W: .Height = ROW_H: .TextAlign = fmTextAlignCenter: .Font.Bold = True: End With
    Set hL = host.Controls.Add("Forms.Label.1")
    With hL: .caption = "左": .Left = hR.Left + COL_EDT_W + 8: .Top = Y: .Width = COL_EDT_W: .Height = ROW_H: .TextAlign = fmTextAlignCenter: .Font.Bold = True: End With
    AddHeaderRow = Y + ROW_H + 2
End Function

Private Function AddMMTRow(owner As frmEval, host As MSForms.Frame, muscle As String, Y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl: .caption = muscle: .Left = PAD_X: .Top = Y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.Controls.Add("Forms.ComboBox.1")
    Set cboL = host.Controls.Add("Forms.ComboBox.1")

    SetupMMTCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = Y
    SetupMMTCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = Y

    cboR.tag = TAG_FUNC_PREFIX & "|MMT_" & muscle & "_右"
    cboL.tag = TAG_FUNC_PREFIX & "|MMT_" & muscle & "_左"

    ' （BI合計とは無関係なので CboBIHook は未適用）
    AddMMTRow = Y + ROW_H + GAP_Y
End Function

Private Sub SetupMMTCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        ' 一般的なMMT表現：0~5＋±、および「不可」
        .AddItem "0"
        .AddItem "1"
        .AddItem "2-": .AddItem "2": .AddItem "2+"
        .AddItem "3-": .AddItem "3": .AddItem "3+"
        .AddItem "4-": .AddItem "4": .AddItem "4+"
        .AddItem "5"
        .AddItem "不可"
    End With
End Sub


Private Function AddSensoryRow(host As MSForms.Frame, itemKey As String, Y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl: .caption = Replace(itemKey, "_", " / "): .Left = PAD_X: .Top = Y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.Controls.Add("Forms.ComboBox.1")
    Set cboL = host.Controls.Add("Forms.ComboBox.1")
    SetupSensoryCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = Y
    SetupSensoryCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = Y

    cboR.tag = TAG_FUNC_PREFIX & "|SENS_" & itemKey & "_右"
    cboL.tag = TAG_FUNC_PREFIX & "|SENS_" & itemKey & "_左"

    AddSensoryRow = Y + ROW_H + GAP_Y
End Function

Private Sub SetupSensoryCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        .AddItem "正常"
        .AddItem "低下"
        .AddItem "消失"
        .AddItem "未検"
        .AddItem "不可"
    End With
End Sub

Private Function AddMASRow(host As MSForms.Frame, groupName As String, Y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl: .caption = groupName: .Left = PAD_X: .Top = Y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.Controls.Add("Forms.ComboBox.1")
    Set cboL = host.Controls.Add("Forms.ComboBox.1")
    SetupMASCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = Y
    SetupMASCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = Y

    cboR.tag = TAG_FUNC_PREFIX & "|TONE_MAS_" & groupName & "_右"
    cboL.tag = TAG_FUNC_PREFIX & "|TONE_MAS_" & groupName & "_左"

    AddMASRow = Y + ROW_H + GAP_Y
End Function

Private Sub SetupMASCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        .AddItem "0"
        .AddItem "1"
        .AddItem "1+"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "未検"
    End With
End Sub

Private Function AddReflexRow(host As MSForms.Frame, reflexName As String, Y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl: .caption = reflexName: .Left = PAD_X: .Top = Y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.Controls.Add("Forms.ComboBox.1")
    Set cboL = host.Controls.Add("Forms.ComboBox.1")
    SetupReflexCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = Y
    SetupReflexCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = Y

    cboR.tag = TAG_FUNC_PREFIX & "|REFLEX_" & reflexName & "_右"
    cboL.tag = TAG_FUNC_PREFIX & "|REFLEX_" & reflexName & "_左"

    AddReflexRow = Y + ROW_H + GAP_Y
End Function

Private Sub SetupReflexCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        .AddItem "0"
        .AddItem "1+"
        .AddItem "2+"
        .AddItem "3+"
        .AddItem "4+"
        .AddItem "未検"
    End With
End Sub

Private Function AddDeformText(owner As frmEval, host As MSForms.Frame, Y As Single) As Single
    Dim lbl As MSForms.label, txt As MSForms.TextBox

    ' ラベル（←他と同じ列幅を使う）
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = "変形（所見）"
        .Left = PAD_X
        .Top = Y
        .Width = COL1_W            ' ★統一！
        .Height = ROW_H
    End With

    ' テキスト（←入力列の開始位置にスナップ）
    Set txt = host.Controls.Add("Forms.TextBox.1")
    With txt
        .Left = PAD_X + COL1_W + 8      ' ★統一！
        .Top = Y
        ' 横幅と高さは前回の設定を流用（値はあなたの好みで）
        Dim availW As Single
        availW = host.Width - .Left - PAD_X
        .Width = availW * DEFORM_W_RATE
        .Height = DEFORM_H

        .multiline = True
        .EnterKeyBehavior = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .tag = TAG_FUNC_PREFIX & "|PAIN|変形_所見"
    End With

    ' IME hook（省略可）
    On Error Resume Next
    Dim ime As New TxtImeHook
    ime.Init txt: owner.RegisterTxtHook ime
    On Error GoTo 0

    ' 次のY
    AddDeformText = txt.Top + txt.Height + GAP_Y
End Function



Private Function AddPainRow(owner As frmEval, host As MSForms.Frame, Y As Single) As Single
    ' 追加は全部ローカル定数：外部に依存しない
    Const NRS_LBL_W As Single = 28
    Const NRS_CBO_W As Single = 60
    Const GAP_X     As Single = 8

    ' ラベル「疼痛（部位）」
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = "疼痛（部位）"
        .Left = PAD_X: .Top = Y
        .Width = COL1_W: .Height = ROW_H
    End With

    ' 右側に NRS（ラベル＋コンボ）を先に配置して基準にする
    Dim lblN As MSForms.label, cbo As MSForms.ComboBox
    Set lblN = host.Controls.Add("Forms.Label.1")
    With lblN
        .caption = "NRS"
        .Top = Y: .Width = NRS_LBL_W: .Height = ROW_H
    End With

    Set cbo = host.Controls.Add("Forms.ComboBox.1")
    With cbo
        .Top = Y: .Width = NRS_CBO_W: .Height = ROW_H
        .Style = fmStyleDropDownList
        .tag = TAG_FUNC_PREFIX & "|PAIN_NRS"
        ' 必要なら 0～10 を自動で埋める（既に設定しているなら何もしない）
        If .ListCount = 0 Then
            Dim i As Integer
            For i = 0 To 10: .AddItem CStr(i): Next i
        End If
    End With

    ' 右端に揃える
    Dim rightEdge As Single: rightEdge = host.InsideWidth - PAD_X
    cbo.Left = rightEdge - NRS_CBO_W
    lblN.Left = cbo.Left - GAP_X - NRS_LBL_W

    ' 部位テキストは「入力列開始」から NRS 手前までを自動幅で
    Dim txt As MSForms.TextBox
    Set txt = host.Controls.Add("Forms.TextBox.1")
    With txt
        .Top = Y
        .Left = PAD_X + COL1_W + GAP_X
        .Width = lblN.Left - GAP_X - .Left
        If .Width < 80 Then .Width = 80        ' 安全弁
        .Height = ROW_H
        .tag = TAG_FUNC_PREFIX & "|PAIN_部位"
        .EnterKeyBehavior = False
        .multiline = False
        ' 日本語入力のままでOK。半角固定なら: .IMEMode = fmIMEModeDisable
    End With

    AddPainRow = Y + ROW_H + GAP_Y
End Function


'========================================
' 備考（自由記述）※IMEひらがなON
'========================================

Private Sub AddNotesBox(owner As frmEval, host As MSForms.Frame, keyPrefix As String)
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = "備考（自由記述）"
        .Left = PAD_X
        .Width = 120
        .Height = 18
    End With

    Dim txt As MSForms.TextBox
    Set txt = host.Controls.Add("Forms.TextBox.1")
    With txt
        .Left = lbl.Left + lbl.Width + 8
         Dim availW As Single
    availW = host.Width - .Left - PAD_X
    .Width = availW * NOTE_W_RATE
        .Height = NOTE_H
        .multiline = True
        .EnterKeyBehavior = True
        .tag = keyPrefix & "|備考"
    End With

    ' ★ここがポイント：現在の内容の直下に並べる
    Dim bottom As Single
    bottom = GetContentBottom(host)
    lbl.Top = bottom + 8
    txt.Top = lbl.Top

    ' スクロール終端を備考の下まで伸ばす
    host.ScrollBars = fmScrollBarsVertical
    host.ScrollHeight = txt.Top + txt.Height + PAD_Y

    ' IME（備考は日本語想定なので On のまま。ひらがな Hook は既存通り）
    On Error Resume Next
    Dim ime As New TxtImeHook
    ime.Init txt
    owner.RegisterTxtHook ime
    On Error GoTo 0
End Sub


Private Function FindRootMulti(frm As frmEval) As MSForms.MultiPage
    Debug.Print "[phys] search root MP"
    Set FindRootMulti = FindMultiPageRecursive(frm)
End Function

Private Function FindMultiPageRecursive(parent As Object) As MSForms.MultiPage
    Dim c As Control
    For Each c In parent.Controls
        If TypeOf c Is MSForms.MultiPage Then
            Debug.Print "  [hit] MP Name=" & c.name & " (Parent=" & TypeName(parent) & ")"
            If Not IsKnownMpADL(c) Then
                Set FindMultiPageRecursive = c
                Debug.Print "  [use] as root"
                Exit Function
            Else
                Debug.Print "  [skip] mpADL"
            End If
        End If
        If TypeOf c Is MSForms.Frame Then
            Dim M As MSForms.MultiPage
            Set M = FindMultiPageRecursive(c)
            If Not M Is Nothing Then
                Set FindMultiPageRecursive = M
                Exit Function
            End If
        End If
    Next
End Function


' mpADL 判定（名前 or ページ見出しで判定）
Private Function IsKnownMpADL(mp As MSForms.MultiPage) As Boolean
    On Error Resume Next
    If LCase$(mp.name) = "mpadl" Then
        IsKnownMpADL = True
        Exit Function
    End If
    If mp.Pages.Count >= 3 Then
        Dim c0$, c1$, c2$
        c0 = mp.Pages(0).caption
        c1 = mp.Pages(1).caption
        c2 = mp.Pages(2).caption
        If (InStr(c0, "BI") > 0 Or InStr(c0, "バーサル") > 0) _
        And (InStr(c1, "IADL") > 0) _
        And (InStr(c2, "起居") > 0) Then
            IsKnownMpADL = True
        End If
    End If
End Function


'=== 追加：身体機能評価を「親タブ」（ルート）として作成（備考は各子タブに配置） ===



Public Sub EnsurePhysicalFunctionTabs_Root(owner As frmEval)
    Debug.Print "[phys] enter EnsurePhysicalFunctionTabs_Root"

    Dim root As MSForms.MultiPage: Set root = FindRootMulti(owner)
    If root Is Nothing Then
        Debug.Print "[phys] root not found"
        ' 見つからないときの状況をダンプ
        Call DumpMultiPages(owner)
        MsgBox "最上段のMultiPageが見つかりません。イミディエイト(CTRL+G)のログを教えてください。"
        Exit Sub
    Else
        Debug.Print "[phys] root found: Name=" & root.name & ", Pages=" & root.Pages.Count
    End If
   


    ' ルートに「身体機能評価」ページを追加/取得
    Dim pgPhys As MSForms.Page
    Set pgPhys = FindOrAddPage(root, CAP_FUNC)

    ' ページ内フレーム
    Dim host As MSForms.Frame
    Set host = EnsureHostFrame(pgPhys)

    ' ページ内に“子タブ用”のMultiPage（mpPhys）を追加/取得
    Dim mp As MSForms.MultiPage
    Dim c As Control
    For Each c In host.Controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then Set mp = c: Exit For
        End If
    Next
    If mp Is Nothing Then
        Set mp = host.Controls.Add("Forms.MultiPage.1")
        With mp
            .name = MP_PHYS_NAME
            .Left = PAD_X
            .Top = PAD_Y
            .Width = host.Width - PAD_X * 2
            .Height = host.Height - PAD_Y * 2
        End With
        ' タブ切替フック
        On Error Resume Next
        Dim mph As New MPHook
        mph.Init owner, mp
        owner.RegisterMPHook mph
        On Error GoTo 0
    End If

    ' 子タブ（3枚）：ROM / MMT / 感覚・痙縮・反射・疼痛（各タブに備考あり）
    Dim pgRom As MSForms.Page, pgMMT As MSForms.Page, pgSens As MSForms.Page
    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)
    Set pgSens = FindOrAddPage(mp, CAP_FUNC_SENS_REF)

    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame, hostSens As MSForms.Frame
    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)

    ' UI生成＋各タブに備考
    BuildROMSection_Compact hostRom
    

    BuildMMTSection owner, hostMmt
   

    BuildSensoryToneReflexPain owner, hostSens
    

    ' 初期表示
    mp.value = pgRom.Index
    root.value = pgPhys.Index
End Sub


Private Sub DumpMultiPages(frm As frmEval)
    Debug.Print "------ MultiPages on frmEval ------"
    Call DumpMP_Recur(frm, 0)
    Debug.Print "-----------------------------------"
End Sub

Private Sub DumpMP_Recur(parent As Object, ByVal depth As Long)
    Dim c As Control, i As Long, pad$
    pad = String$(depth * 2, " ")
    For Each c In parent.Controls
        If TypeOf c Is MSForms.MultiPage Then
            On Error Resume Next
            Debug.Print pad & "MP Name=" & c.name & " Pages=" & c.Pages.Count
            For i = 0 To c.Pages.Count - 1
                Debug.Print pad & "  - Page(" & i & "): " & c.Pages(i).caption
            Next
            On Error GoTo 0
        End If
        If TypeOf c Is MSForms.Frame Then
            Debug.Print pad & "Frame: " & c.name
            Call DumpMP_Recur(c, depth + 1)
        End If
    Next
End Sub


'=== 追加：指定されたルート MultiPage に「身体機能評価」ページを作る ===
Public Sub EnsurePhysicalFunctionTabs_Under(owner As frmEval, root As MSForms.MultiPage)
    If root Is Nothing Then
        MsgBox "ルート MultiPage が Nothing です。呼び出し側で 'mp' を渡してください。", vbExclamation
        Exit Sub
    End If

    ' ルートに「身体機能評価」ページを追加/取得
    Dim pgPhys As MSForms.Page
    Set pgPhys = FindOrAddPage(root, CAP_FUNC)

    ' ページ内の作業フレーム
    Dim host As MSForms.Frame
    Set host = EnsureHostFrame(pgPhys)

        ' --- 子タブ MultiPage（mpPhys）を作る or 取得 ---
    Dim mp As MSForms.MultiPage
    Dim c As Control
    For Each c In host.Controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then Set mp = c: Exit For
        End If
    Next
    If mp Is Nothing Then
        Set mp = host.Controls.Add("Forms.MultiPage.1")
        With mp
            .name = MP_PHYS_NAME
            .Left = PAD_X
            .Top = PAD_Y
            .Width = host.Width - PAD_X * 2
            .Height = host.Height - PAD_Y * 2
        End With
    End If

       ' ★ Page8/9などの既定ページを掃除
    CleanDefaultPages mp

    ' --- 子タブ（5枚）を用意 ---
    Dim pgRom As MSForms.Page, pgMMT As MSForms.Page
    Dim pgSens As MSForms.Page, pgToneRef As MSForms.Page, pgPain As MSForms.Page

    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)                     ' ROM（主要関節）
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)                     ' 筋力（MMT）
    Set pgSens = FindOrAddPage(mp, "感覚（表在・深部）")              ' ←分離
    Set pgToneRef = FindOrAddPage(mp, "筋緊張・反射（痙縮含む）")     ' ←分離
    Set pgPain = FindOrAddPage(mp, "疼痛（部位／NRS）")               ' ←分離

    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame
    Dim hostSens As MSForms.Frame, hostTone As MSForms.Frame, hostPain As MSForms.Frame

    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)
    Set hostTone = EnsureHostFrame(pgToneRef)
    Set hostPain = EnsureHostFrame(pgPain)
    
    
    
' （EnsurePhysicalFunctionTabs_* の中、他のpg～と同じ並びに）
Dim pgPar As MSForms.Page, hostPar As MSForms.Frame
Set pgPar = FindOrAddPage(mp, CAP_FUNC_PARALYSIS)
Set hostPar = EnsureHostFrame(pgPar)
BuildParalysisTabUI hostPar



' --- UI構築 ---
If USE_ROM_SUBTABS Then
    BuildROMTabs hostRom            ' ← 新UI（上肢／下肢 子タブ）
Else
    BuildROMSection_Compact hostRom ' ← 旧UI（残置）
End If

BuildMMTSection owner, hostMmt
BuildSensoryTabUI hostSens
BuildToneReflexTabUI hostTone
BuildPainTabUI owner, hostPain



    ' 初期表示
    mp.value = pgRom.Index


End Sub




'=== 既定ページ "Page*" を掃除 ===
Private Sub CleanDefaultPages(mp As MSForms.MultiPage)
    On Error Resume Next
    Dim i As Long
    For i = mp.Pages.Count - 1 To 0 Step -1
        If Left$(mp.Pages(i).caption, 4) = "Page" Then
            mp.Pages.Remove i
        End If
    Next
End Sub


'=== 感覚のみ ===
Private Sub BuildSensoryTabUI(host As MSForms.Frame)
    Dim Y As Single: Y = PAD_Y

    Y = AddHeaderRow(host, "感覚（表在）", Y)
    Y = AddSensoryRow(host, "表在_触覚", Y)
    Y = AddSensoryRow(host, "表在_痛覚", Y)
    Y = AddSensoryRow(host, "表在_温度覚", Y)

    Y = Y + ROM_HDR_GAP
    Y = AddHeaderRow(host, "感覚（深部）", Y)
    Y = AddSensoryRow(host, "深部_位置覚", Y)
    Y = AddSensoryRow(host, "深部_振動覚", Y)

    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, Y + ROM_HDR_GAP, "txtSensMemo")
End Sub







'=== そのフレーム内で一番下の位置（Top+Height）を返す ===
Private Function GetContentBottom(host As MSForms.Frame) As Single
    Dim c As Control, bottom As Single
    For Each c In host.Controls
        If c.Visible Then
            If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next
    GetContentBottom = bottom
End Function




'=== ROMの1行（屈曲/外転/…）小さめ版：右・左テキストは半角 ===
Private Function AddROMRow_Compact( _
    host As MSForms.Frame, _
    jointName As String, _
    moveName As String, _
    Y As Single _
) As Single

    Dim xR As Single, xL As Single: ROM_GetCols xR, xL
Dim yPix As Single: yPix = PX(Y)

Dim txtR As MSForms.TextBox, txtL As MSForms.TextBox
Set txtR = host.Controls.Add("Forms.TextBox.1")
With txtR
    .Left = PX(xR)
    .Top = yPix
    .Width = PX(ROM_COL_EDT_W)
    .Height = ROM_ROW_H
    .TextAlign = fmTextAlignCenter
    .IMEMode = fmIMEModeDisable
End With

Set txtL = host.Controls.Add("Forms.TextBox.1")
With txtL
    .Left = PX(xL)
    .Top = yPix
    .Width = PX(ROM_COL_EDT_W)
    .Height = ROM_ROW_H
    .TextAlign = fmTextAlignCenter
    .IMEMode = fmIMEModeDisable
End With

txtR.tag = TAG_FUNC_PREFIX & "|ROM|" & jointName & "|" & moveName & "|右"
txtL.tag = TAG_FUNC_PREFIX & "|ROM|" & jointName & "|" & moveName & "|左"

AddROMRow_Compact = yPix + ROM_ROW_H + ROM_GAP_Y
End Function





' 後方互換ラッパー
Public Sub BuildROMSection_Compact(host As MSForms.Frame)
    BuildROMSection_TwoCols host
End Sub


' 見出し（【肩】など）
Private Function ROM_AddHeader(host As MSForms.Frame, title As String, y0 As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = "【" & title & "】"
        .Left = PAD_X: .Top = y0: .Width = COL1_W: .Height = ROM_ROW_H
        .Font.Bold = True
    End With
    ROM_AddHeader = y0 + ROM_ROW_H + ROM_GAP_Y
End Function


'=== ROM compact 用の呼び出し名統一ラッパー =========================

' 見出し（「運動」「右(°)」「左(°)」の行）
Private Function ROM_AddDirHeader(host As MSForms.Frame, y0 As Single) As Single
    ROM_AddDirHeader = AddROMDirHeader_Compact(host, y0)
End Function

'=== ROM compact：R / L 見出し（ROM専用幅で） ===
Private Function AddROMDirHeader_Compact(host As MSForms.Frame, y0 As Single) As Single
    Dim xR As Single, xL As Single: ROM_GetCols xR, xL
    Dim Y As Single: Y = PX(y0)              ' ← y も丸めておく（任意だけど安定します）

    Dim lblR As MSForms.label, lblL As MSForms.label

    Set lblR = host.Controls.Add("Forms.Label.1")
    With lblR
        .caption = "R"
        .Left = PX(xR)                        ' ← ここに PX
        .Top = Y
        .Width = PX(ROM_COL_EDT_W)            ' ← ここに PX
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Set lblL = host.Controls.Add("Forms.Label.1")
    With lblL
        .caption = "L"
        .Left = PX(xL)                        ' ← ここに PX
        .Top = Y
        .Width = PX(ROM_COL_EDT_W)            ' ← ここに PX
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    AddROMDirHeader_Compact = Y + ROM_ROW_H + ROM_GAP_Y
End Function




'========================================
' ROM（二列：左=上肢 / 右=下肢）本体
'========================================
Public Sub BuildROMSection_TwoCols(host As MSForms.Frame)
    Const COL_GAP_X As Single = 12

    ' 既存配置があれば除去（重複描画防止）
    On Error Resume Next
    host.Controls.Remove "fraROM_Upper"
    host.Controls.Remove "fraROM_Lower"
    host.Controls.Remove "txtROMMemo"
    On Error GoTo 0

    ' ←ここは残す：On Error GoTo 0 の直後から差し替え

Dim w As Single, h As Single, colW As Single
w = host.InsideWidth: h = host.InsideHeight
' 列幅を整数に丸める
colW = PX((w - (PAD_X * 2) - COL_GAP_X) / 2)

Dim frUL As MSForms.Frame, frLL As MSForms.Frame

' 左列フレーム（上肢）
Set frUL = host.Controls.Add("Forms.Frame.1", "fraROM_Upper")
With frUL
    .caption = ""
    .Left = PX(PAD_X)                 ' ★整数丸め
    .Top = PAD_Y
    .Width = colW                     ' ★丸め済み列幅
    .Height = h - PAD_Y * 2
    .ScrollBars = fmScrollBarsNone
    .ScrollHeight = .InsideHeight
End With

' 右列フレーム（下肢）
Set frLL = host.Controls.Add("Forms.Frame.1", "fraROM_Lower")
With frLL
    .caption = ""
    .Left = PX(PAD_X + colW + COL_GAP_X)  ' ★整数丸め
    .Top = PAD_Y
    .Width = colW                          ' ★丸め済み列幅
    .Height = h - PAD_Y * 2
    .ScrollBars = fmScrollBarsNone
    .ScrollHeight = .InsideHeight
End With


    Dim yL As Single, yR As Single
    yL = PAD_Y: yR = PAD_Y

    ' ---------- 左列：上肢 ----------
    yL = ROM_AddHeader(frUL, "上肢", yL)
    yL = ROM_AddHeader(frUL, "【肩関節】", yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "肩", "屈曲", yL)
    yL = AddROMRow_Compact(frUL, "肩", "外転", yL)
    yL = AddROMRow_Compact(frUL, "肩", "外旋", yL)

    yL = ROM_AddHeader(frUL, "【肘関節】", yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "肘", "屈曲", yL)
    yL = AddROMRow_Compact(frUL, "肘", "伸展", yL)

    yL = ROM_AddHeader(frUL, "【前腕】", yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "前腕", "回内", yL)
    yL = AddROMRow_Compact(frUL, "前腕", "回外", yL)

    yL = ROM_AddHeader(frUL, "【手関節】", yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "手関節", "掌屈", yL)
    yL = AddROMRow_Compact(frUL, "手関節", "背屈", yL)

    ' ---------- 右列：下肢 ----------
    yR = ROM_AddHeader(frLL, "下肢", yR)
    yR = ROM_AddHeader(frLL, "【股関節】", yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "股", "屈曲", yR)
    yR = AddROMRow_Compact(frLL, "股", "外転", yR)

    yR = ROM_AddHeader(frLL, "【膝関節】", yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "膝", "屈曲", yR)
    yR = AddROMRow_Compact(frLL, "膝", "伸展", yR)

    yR = ROM_AddHeader(frLL, "【足関節】", yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "足関節", "背屈", yR)
    yR = AddROMRow_Compact(frLL, "足関節", "底屈", yR)

    ' 備考は共通ヘルパーに統一（自動でhostスクロールもOFF）
    Call PlaceMemoBelow(host, w, h, Application.WorksheetFunction.Max(yL, yR) + ROM_HDR_GAP, _
                        "txtROMMemo", frUL, frLL)
    ' …上肢/下肢の行をすべて作り終えた直後に…
    NormalizeRomColumns frUL
    NormalizeRomColumns frLL

End Sub


'========================================
' 筋緊張（MAS）＋ 反射（ビルダー）
' 呼び出し例: BuildToneReflexTabUI hostReflex
'========================================
Private Sub BuildToneReflexTabUI(host As MSForms.Frame)
    Dim Y As Single: Y = PAD_Y

    ' --- 筋緊張（MAS） ---
    Y = AddHeaderRow(host, "筋緊張（MAS）", Y)
    Y = AddMASRow(host, "上肢屈筋群", Y)
    Y = AddMASRow(host, "上肢伸筋群", Y)
    Y = AddMASRow(host, "下肢屈筋群", Y)
    Y = AddMASRow(host, "下肢伸筋群", Y)

    ' --- 反射 ---
    Y = Y + ROM_HDR_GAP
    Y = AddHeaderRow(host, "腱反射", Y)
    Y = AddReflexRow(host, "上腕二頭筋（C5-6）", Y)
    Y = AddReflexRow(host, "上腕三頭筋（C7）", Y)
    Y = AddReflexRow(host, "膝蓋腱（L2-4）", Y)
    Y = AddReflexRow(host, "アキレス腱（S1）", Y)

    ' 備考（下端に確保）
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, Y + ROM_HDR_GAP, "txtReflexMemo")
End Sub

'========================================
' 疼痛（変形テキスト＋部位＋NRS）ビルダー
' 呼び出し例: BuildPainTabUI owner, hostPain
'========================================
Private Sub BuildPainTabUI(owner As frmEval, host As MSForms.Frame)
    Dim Y As Single: Y = PAD_Y

    Y = AddDeformText(owner, host, Y) ' 変形 所見（自由テキスト）
    Y = Y + ROM_HDR_GAP
    Y = AddPainRow(owner, host, Y)    ' 部位＋NRS

    ' 備考（下端に確保）
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, Y + ROM_HDR_GAP, "txtPainMemo")
End Sub


'========================================
' 互換: 感覚＋MAS＋反射＋疼痛 を1ページに描く（Root用）
' 呼び出し元: EnsurePhysicalFunctionTabs_Root
'========================================
Private Sub BuildSensoryToneReflexPain(owner As frmEval, host As MSForms.Frame)
    Dim Y As Single: Y = PAD_Y

    ' --- 感覚（表在） ---
    Y = AddHeaderRow(host, "感覚（表在）", Y)
    Y = AddSensoryRow(host, "表在_触覚", Y)
    Y = AddSensoryRow(host, "表在_痛覚", Y)
    Y = AddSensoryRow(host, "表在_温度覚", Y)

    ' --- 感覚（深部） ---
    Y = Y + ROM_HDR_GAP
    Y = AddHeaderRow(host, "感覚（深部）", Y)
    Y = AddSensoryRow(host, "深部_関節位置覚", Y)
    Y = AddSensoryRow(host, "深部_振動覚", Y)

    ' --- 筋緊張（MAS） ---
    Y = Y + ROM_HDR_GAP
    Y = AddHeaderRow(host, "筋緊張（MAS）", Y)
    Y = AddMASRow(host, "上肢屈筋群", Y)
    Y = AddMASRow(host, "上肢伸筋群", Y)
    Y = AddMASRow(host, "下肢屈筋群", Y)
    Y = AddMASRow(host, "下肢伸筋群", Y)

    ' --- 反射 ---
    Y = Y + ROM_HDR_GAP
    Y = AddHeaderRow(host, "腱反射", Y)
    Y = AddReflexRow(host, "上腕二頭筋（C5-6）", Y)
    Y = AddReflexRow(host, "上腕三頭筋（C7）", Y)
    Y = AddReflexRow(host, "膝蓋腱（L2-4）", Y)
    Y = AddReflexRow(host, "アキレス腱（S1）", Y)

    ' --- 変形（自由テキスト） ---
    Y = Y + ROM_HDR_GAP
    Y = AddDeformText(owner, host, Y)

    ' --- 疼痛（部位＋NRS） ---
    Y = Y + ROM_HDR_GAP
    Y = AddPainRow(owner, host, Y)

    ' 備考：ページ下端に確保
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, Y + ROM_HDR_GAP, "txtSensReflexPainMemo")
End Sub

' --- ROM 1行分の右(R)/左(L)のX座標を返すだけ ---
Private Sub ROM_GetCols(ByRef xR As Single, ByRef xL As Single)
    Const GAP_X As Single = 8  ' 既に定数があればこの行は不要
    xR = PX(PAD_X + COL1_W + GAP_X)            ' 右入力欄の左端
    xL = PX(xR + ROM_COL_EDT_W + GAP_X)         ' 左入力欄の左端
End Sub



' 小数 → 最寄りピクセルにスナップ
Public Function PX(v As Single) As Single
    PX = Int(v + 0.5)
End Function


' ROMコラムのX位置/幅をフレーム内で一括整列
Private Sub NormalizeRomColumns(host As MSForms.Frame)
    Dim xR As Single, xL As Single
    ROM_GetCols xR, xL
    xR = PX(xR): xL = PX(xL)
    Dim w As Single: w = PX(ROM_COL_EDT_W)

    Dim c As Control
    For Each c In host.Controls
        Select Case TypeName(c)
            Case "TextBox"
                ' タグに「右」「左」が入るようにしておく（③参照）
                If InStr(c.tag, "右") > 0 Then c.Left = xR: c.Width = w
                If InStr(c.tag, "左") > 0 Then c.Left = xL: c.Width = w
            Case "Label"
                If c.caption = "R" Then c.Left = xR: c.Width = w
                If c.caption = "L" Then c.Left = xL: c.Width = w
        End Select
    Next
End Sub

'========================================================
' 麻痺タブ UI
'========================================================
Public Sub BuildParalysisTabUI(host As MSForms.Frame)
    Dim w As Single, h As Single, Y As Single
    w = host.Width: h = host.Height
    Y = PAD_Y

    ' ---- 基本情報 ----
    Y = AddSectionTitle(host, "基本情報", Y)
    Y = AddComboRow(host, "麻痺側", "cboParalysisSide", Array("右", "左", "両側"), Y)
    Y = AddComboRow(host, "麻痺の種類", "cboParalysisType", Array("片麻痺", "四肢麻痺", "単麻痺"), Y)

    ' ---- BRS ----
    Y = Y + ROM_HDR_GAP
    Y = AddSectionTitle(host, "Brunnstrom Recovery Stage（BRS）", Y)
    Dim brsValues As Variant
    brsValues = Array("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ")
    Y = AddComboRow(host, "上肢", "cboBRS_Upper", brsValues, Y)
    Y = AddComboRow(host, "手指", "cboBRS_Hand", brsValues, Y)
    Y = AddComboRow(host, "下肢", "cboBRS_Lower", brsValues, Y)

    ' ---- 随伴現象 ----
    Y = Y + ROM_HDR_GAP
    Y = AddSectionTitle(host, "随伴現象", Y)
    Y = AddCheckRow(host, "共同運動", "chkSynergy", Y)
    Y = AddCheckRow(host, "連合反応", "chkAssociatedRxn", Y)

    ' ---- 備考 ----
    PlaceMemoBelow host, w, h, Y, "txtParalysisMemo"
End Sub

'---- 小物：行ビルダ（見出し／コンボ行／チェック行）----
Private Function AddSectionTitle(host As MSForms.Frame, ttl As String, Y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = ttl
        .Left = PX(PAD_X)
        .Top = PX(Y)
        .Width = PX(host.Width - PAD_X * 2)
        .Height = ROW_H
        .Font.Bold = True
    End With
    AddSectionTitle = PX(Y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddComboRow(host As MSForms.Frame, cap As String, nameCombo As String, _
                             items As Variant, Y As Single) As Single
    Dim wCaption As Single, wCombo As Single, xCaption As Single, xCombo As Single
    wCaption = PX(COL1_W)
    wCombo = PX(ROM_COL_EDT_W)
    xCaption = PX(PAD_X)
    xCombo = PX(PAD_X + wCaption + 8)

    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCaption
        .Top = PX(Y)
        .Width = wCaption
        .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim cbo As MSForms.ComboBox
    Set cbo = host.Controls.Add("Forms.ComboBox.1", nameCombo, True)
    With cbo
        .Left = xCombo
        .Top = PX(Y)
        .Width = wCombo
        .Height = ROW_H
        .Style = fmStyleDropDownList
        .List = items
    End With

    AddComboRow = PX(Y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddCheckRow(host As MSForms.Frame, cap As String, nameChk As String, Y As Single) As Single
    Dim wCaption As Single, xCaption As Single, xChk As Single
    wCaption = PX(COL1_W)
    xCaption = PX(PAD_X)
    xChk = PX(PAD_X + wCaption + 8)

    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCaption
        .Top = PX(Y)
        .Width = wCaption
        .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim chk As MSForms.CheckBox
    Set chk = host.Controls.Add("Forms.CheckBox.1", nameChk, True)
    With chk
        .caption = "有"
        .Left = xChk
        .Top = PX(Y)
        .Width = PX(60)
        .Height = ROW_H
    End With

    AddCheckRow = PX(Y + ROW_H + ROM_GAP_Y)
End Function


'============================================================
' ROM ページ：子タブ（上肢／下肢）ビルダー
'============================================================
Public Sub BuildROMTabs(host As MSForms.Frame)
    Debug.Print "[ROM] BuildROMTabs called"

    Dim mp As MSForms.MultiPage
    Set mp = host.Controls.Add("Forms.MultiPage.1", "mpROM")
    With mp
        .Left = PX(PAD_X)
        .Top = PX(PAD_Y)
        .Width = PX(host.InsideWidth - PAD_X * 2)
        .Height = PX(host.InsideHeight - PAD_Y * 2)
        .Style = fmTabStyleTabs
    End With

    Dim pUpper As MSForms.Page, pLower As MSForms.Page
    Set pUpper = mp.Pages.Add: pUpper.caption = "上肢": pUpper.name = "pgROM_Upper"
    Set pLower = mp.Pages.Add: pLower.caption = "下肢": pLower.name = "pgROM_Lower"

    Dim hostUpper As MSForms.Frame, hostLower As MSForms.Frame
    Set hostUpper = EnsureHostFrame(pUpper)
    Set hostLower = EnsureHostFrame(pLower)

    BuildROM_Upper hostUpper
    NormalizeRomColumns hostUpper

    BuildROM_Lower hostLower
    NormalizeRomColumns hostLower

  
End Sub



'------------------------------------------------------------
' 上肢 子タブ
'------------------------------------------------------------
Public Sub BuildROM_Upper(host As MSForms.Frame)
    Dim w As Single, h As Single
    w = host.Width: h = host.Height

    Dim Y As Single: Y = ROM_GROUP_PAD
    
    

    ' 肩
    Y = BuildRomJointBlock(host, "Upper", "Shoulder", "【 肩関節 】", _
            Split("Flex,Ext,Abd,Add,ER,IR", ","), Y)
    Y = Y + ROM_JOINT_GAP_Y

    ' 肘
    Y = BuildRomJointBlock(host, "Upper", "Elbow", "【 肘関節 】", _
            Split("Flex,Ext", ","), Y)
    Y = Y + ROM_JOINT_GAP_Y

    ' 前腕
    Y = BuildRomJointBlock(host, "Upper", "Forearm", "【 前腕 】", _
            Split("Sup,Pro", ","), Y)
    Y = Y + ROM_JOINT_GAP_Y

    ' 手関節
    Y = BuildRomJointBlock(host, "Upper", "Wrist", "【 手関節 】", _
            Split("Dorsi,Palmar,Radial,Ulnar", ","), Y)
            
            
 PlaceMemoBelow host, w, h, Y, "txtROM_Upper_Memo"
 

End Sub

'------------------------------------------------------------
' 下肢 子タブ
'------------------------------------------------------------
Public Sub BuildROM_Lower(host As MSForms.Frame)
    Dim w As Single, h As Single
    w = host.Width: h = host.Height

    Dim Y As Single: Y = ROM_GROUP_PAD
    
    
    ' 股
    Y = BuildRomJointBlock(host, "Lower", "Hip", "【 股関節 】", _
            Split("Flex,Ext,Abd,Add,ER,IR", ","), Y)
    Y = Y + ROM_JOINT_GAP_Y

    ' 膝
    Y = BuildRomJointBlock(host, "Lower", "Knee", "【 膝関節 】", _
            Split("Flex,Ext", ","), Y)
    Y = Y + ROM_JOINT_GAP_Y

    ' 足関節
    Y = BuildRomJointBlock(host, "Lower", "Ankle", "【 足関節 】", _
            Split("Dorsi,Plantar,Inv,Ev", ","), Y)
            
            
PlaceMemoBelow host, w, h, Y, "txtROM_Lower_Memo"
End Sub

'------------------------------------------------------------
' 関節ブロック（枠＋運動行）
'   motions: "Flex","Ext"...（英略キー）
'   戻り: 次ブロックの開始Top
'------------------------------------------------------------
Private Function BuildRomJointBlock(host As MSForms.Frame, _
            region As String, jointKey As String, jointTitle As String, _
            motions As Variant, y0 As Single) As Single

    Dim fr As MSForms.Frame
    Set fr = host.Controls.Add("Forms.Frame.1")
    With fr
        .caption = jointTitle
        .Left = PX(PAD_X)
        .Top = PX(y0)
        .Width = PX(host.Width - PAD_X * 2)
        ' 高さは行数とヘッダ・パディングから概算
        .Height = PX(ROM_ROW_H * (UBound(motions) - LBound(motions) + 1) + _
                     ROM_GROUP_PAD * 2 + ROM_HDR_GAP + (UBound(motions) - LBound(motions)) * ROM_MOTION_GAP_Y)
    End With

    ' R/L ヘッダ
    Dim xName As Single, xR As Single, xL As Single
    xName = PX(ROM_GROUP_PAD)
    xR = PX(fr.Width - ROM_GROUP_PAD - ROM_COL_EDT_W * 2 - 6)
    xL = PX(fr.Width - ROM_GROUP_PAD - ROM_COL_EDT_W)

    Dim lblR As MSForms.label, lblL As MSForms.label
    Set lblR = fr.Controls.Add("Forms.Label.1")
    With lblR
        .caption = "R": .Left = xR: .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter: .Font.Bold = True
    End With
    Set lblL = fr.Controls.Add("Forms.Label.1")
    With lblL
        .caption = "L": .Left = xL: .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter: .Font.Bold = True
    End With

    ' 運動行
    Dim i As Long, topY As Single
    topY = PX(ROM_GROUP_PAD + ROM_HDR_GAP)

    For i = LBound(motions) To UBound(motions)
        topY = BuildRomMotionRow(fr, region, jointKey, CStr(motions(i)), _
                                 topY, xName, xR, xL)
        topY = topY + ROM_MOTION_GAP_Y
    Next i

    ' IME Off（TxtImeHook）をフレーム内のTextBoxへ一括アタッチ
    AttachTxtImeHookInFrame fr

    BuildRomJointBlock = fr.Top + fr.Height + ROM_GAP_Y
End Function

'------------------------------------------------------------
' 1運動分の行（ラベル＋R/Lテキスト）
'------------------------------------------------------------
Private Function BuildRomMotionRow(host As MSForms.Frame, _
        region As String, jointKey As String, motionKey As String, _
        y0 As Single, xName As Single, xR As Single, xL As Single) As Single

    Dim lbl As MSForms.label
    Set lbl = host.Controls.Add("Forms.Label.1")
    With lbl
        .caption = GetMotionCaption(jointKey, motionKey)
        .Left = xName: .Top = PX(y0)
        .Width = PX(host.Width - xName - (host.Width - xR) + ROM_HDR_RL_GAP)
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim tR As MSForms.TextBox, tL As MSForms.TextBox
    ' 右（R）
Set tR = host.Controls.Add("Forms.TextBox.1", _
    "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_R")
With tR
    .Left = xR: .Top = PX(y0)
    .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
    .IMEMode = fmIMEModeDisable
    .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|E"   ' ★ここを追加
End With

' 左（L）
Set tL = host.Controls.Add("Forms.TextBox.1", _
    "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_L")
With tL
    .Left = xL: .Top = PX(y0)
    .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
    .IMEMode = fmIMEModeDisable
    .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|"    ' ★ここを追加（末尾は空＝L）
End With


    BuildRomMotionRow = tR.Top + ROM_ROW_H
End Function

'------------------------------------------------------------
' 運動ラベル（和名）
'------------------------------------------------------------
Private Function GetMotionCaption(jointKey As String, motionKey As String) As String
    Select Case jointKey
        Case "Shoulder", "Hip"
            Select Case motionKey
                Case "Flex":   GetMotionCaption = "屈曲"
                Case "Ext":    GetMotionCaption = "伸展"
                Case "Abd":    GetMotionCaption = "外転"
                Case "Add":    GetMotionCaption = "内転"
                Case "ER":     GetMotionCaption = "外旋"
                Case "IR":     GetMotionCaption = "内旋"
            End Select
        Case "Elbow", "Knee"
            Select Case motionKey
                Case "Flex":   GetMotionCaption = "屈曲"
                Case "Ext":    GetMotionCaption = "伸展"
            End Select
        Case "Forearm"
            Select Case motionKey
                Case "Sup":    GetMotionCaption = "回外"
                Case "Pro":    GetMotionCaption = "回内"
            End Select
        Case "Wrist"
            Select Case motionKey
                Case "Dorsi":  GetMotionCaption = "背屈"
                Case "Palmar": GetMotionCaption = "掌屈"
                Case "Radial": GetMotionCaption = "橈屈"
                Case "Ulnar":  GetMotionCaption = "尺屈"
            End Select
        Case "Ankle"
            Select Case motionKey
                Case "Dorsi":   GetMotionCaption = "背屈"
                Case "Plantar": GetMotionCaption = "底屈"
                Case "Inv":     GetMotionCaption = "内がえし"
                Case "Ev":      GetMotionCaption = "外がえし"
            End Select
    End Select
End Function

'------------------------------------------------------------
' フレーム内の TextBox に IME Off フックを一括アタッチ
'   ※ TxtImeHook.cls の公開メソッド名 Attach を想定
'------------------------------------------------------------
Private Sub AttachTxtImeHookInFrame(fr As MSForms.Frame)
    On Error Resume Next
    Dim ctl As MSForms.Control
    For Each ctl In fr.Controls
        If TypeOf ctl Is MSForms.TextBox Then
            Dim Hook As TxtImeHook
            Set Hook = New TxtImeHook
            Hook.Init ctl   ' ★ここを Attach → Init に
        End If
    Next
    On Error GoTo 0
End Sub


'=== LoadLatestROMNow（2025-10-22統合版）===
Public Sub LoadLatestROMNow(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("ROM_Upper_Shoulder_Flex_R", ws)
    If r <= 0 Then
        Debug.Print "[LoadROM] header not found"
        Exit Sub
    End If

    '--- 旧読込は後方互換のためコメントアウト ---
    'Call ParseROMData(ws.Cells(r, HeaderCol("IO_ROM", ws)).Value)
    '-----------------------------------------------------

    Dim raw As String
    raw = LoadLatestROMNow_Raw(ws)
    Debug.Print "[LoadROM] R=" & r & " Len=" & Len(raw) & " | " & Left$(raw, 60)
End Sub



