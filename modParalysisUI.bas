Attribute VB_Name = "modParalysisUI"

Option Explicit

' ===== UI逕ｨ縺ｮ繝ｭ繝ｼ繧ｫ繝ｫ螳壽焚・医％縺ｮ繝｢繧ｸ繝･繝ｼ繝ｫ縺縺代〒譛牙柑・・====
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

    ' 隕句・縺・蝓ｺ譛ｬ諠・ｱ
    y = AddSectionTitle_(host, "蝓ｺ譛ｬ諠・ｱ", y)
    y = AddComboRow_(host, "鮗ｻ逞ｺ蛛ｴ", "cboParalysisSide", Array("蜿ｳ", "蟾ｦ", "荳｡蛛ｴ"), y)
    y = AddComboRow_(host, "鮗ｻ逞ｺ縺ｮ遞ｮ鬘・, "cboParalysisType", Array("迚・ｺｻ逞ｺ", "蝗幄い鮗ｻ逞ｺ", "蜊倬ｺｻ逞ｺ"), y)

    ' 隕句・縺・BRS
    y = y + ROM_HDR_GAP
y = AddSectionTitle_(host, "Brunnstrom Recovery Stage(BRS)", y)

Dim brsValues As Variant          ' 竊・縺薙％繧剃ｿｮ豁｣(Dim 縺ｮ蠕後↓繧ｹ繝壹・繧ｹ)
brsValues = Array("竇", "竇｡", "竇｢", "竇｣", "竇､", "竇･")

y = AddComboRow_(host, "荳願い", "cboBRS_Upper", brsValues, y)
y = AddComboRow_(host, "謇区欠", "cboBRS_Hand", brsValues, y)
y = AddComboRow_(host, "荳玖い", "cboBRS_Lower", brsValues, y)
    ' 隕句・縺・髫丈ｼｴ迴ｾ雎｡
    y = y + ROM_HDR_GAP
    y = AddSectionTitle_(host, "髫丈ｼｴ迴ｾ雎｡", y)
    y = AddCheckRow_(host, "蜈ｱ蜷碁°蜍・, "chkSynergy", y)
    y = AddCheckRow_(host, "騾｣蜷亥渚蠢・, "chkAssociatedRxn", y)

    ' 蛯呵・ｬ・
    PlaceMemoBelow host, w, h, y, "txtParalysisMemo"
End Sub

' ---- 陦後ン繝ｫ繝(Private) ----
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
        .caption = "譛・
        .Left = xChk: .Top = PX(y)
        .Width = PX(60): .Height = ROW_H
    End With

    AddCheckRow_ = PX(y + ROW_H + ROM_GAP_Y)
End Function


