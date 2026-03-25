Attribute VB_Name = "modParalysisUI"

Option Explicit

' ===== UI用のローカル定数（このモジュールだけで有効）=====
Private Const PAD_X As Single = 12
Private Const PAD_Y As Single = 8
Private Const COL1_W As Single = 110
Private Const ROM_COL_EDT_W As Single = 70
Private Const ROW_H As Single = 22
Private Const ROM_HDR_GAP As Single = 10
Private Const ROM_GAP_Y As Single = 6


Public Sub BuildParalysisTabUI(host As MSForms.Frame)
    Dim w As Single, h As Single, y As Single
    w = host.Width: h = host.Height
    y = PAD_Y

    ' 見出し:基本情報
    y = AddSectionTitle_(host, "基本情報", y)
    y = AddComboRow_(host, "麻痺側", "cboParalysisSide", Array("右", "左", "両側"), y)
    y = AddComboRow_(host, "麻痺の種類", "cboParalysisType", Array("片麻痺", "四肢麻痺", "単麻痺"), y)

    ' 見出し:BRS
    y = y + ROM_HDR_GAP
y = AddSectionTitle_(host, "Brunnstrom Recovery Stage(BRS)", y)

Dim brsValues As Variant          ' ← ここを修正(Dim の後にスペース)
brsValues = Array("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ")

y = AddComboRow_(host, "上肢", "cboBRS_Upper", brsValues, y)
y = AddComboRow_(host, "手指", "cboBRS_Hand", brsValues, y)
y = AddComboRow_(host, "下肢", "cboBRS_Lower", brsValues, y)
    ' 見出し:随伴現象
    y = y + ROM_HDR_GAP
    y = AddSectionTitle_(host, "随伴現象", y)
    y = AddCheckRow_(host, "共同運動", "chkSynergy", y)
    y = AddCheckRow_(host, "連合反応", "chkAssociatedRxn", y)

    ' 備考欄
    PlaceMemoBelow host, w, h, y, "txtParalysisMemo"
End Sub

' ---- 行ビルダ(Private) ----
Private Function AddSectionTitle_(host As MSForms.Frame, ttl As String, y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = ttl
        .Left = PX(PAD_X)
        .Top = PX(y)
        .Width = PX(host.Width - PAD_X * 2)
        .Height = ROW_H
        .Font.Bold = True
    End With
    AddSectionTitle_ = PX(y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddComboRow_(host As MSForms.Frame, cap As String, ctlName As String, _
                              items As Variant, y As Single) As Single
    Dim wCap As Single, wCombo As Single, xCap As Single, xCombo As Single
    wCap = PX(COL1_W): wCombo = PX(ROM_COL_EDT_W)
    xCap = PX(PAD_X):  xCombo = PX(PAD_X + wCap + 8)

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCap: .Top = PX(y)
        .Width = wCap: .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim cbo As MSForms.ComboBox
    Set cbo = host.controls.Add("Forms.ComboBox.1", ctlName, True)
    With cbo
        .Left = xCombo: .Top = PX(y)
        .Width = wCombo: .Height = ROW_H
        .Style = fmStyleDropDownList
        .List = items
    End With

    AddComboRow_ = PX(y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddCheckRow_(host As MSForms.Frame, cap As String, ctlName As String, y As Single) As Single
    Dim wCap As Single, xCap As Single, xChk As Single
    wCap = PX(COL1_W): xCap = PX(PAD_X): xChk = PX(PAD_X + wCap + 8)

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCap: .Top = PX(y)
        .Width = wCap: .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim chk As MSForms.CheckBox
    Set chk = host.controls.Add("Forms.CheckBox.1", ctlName, True)
    With chk
        .caption = "有"
        .Left = xChk: .Top = PX(y)
        .Width = PX(60): .Height = ROW_H
    End With

    AddCheckRow_ = PX(y + ROW_H + ROM_GAP_Y)
End Function


