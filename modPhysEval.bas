Attribute VB_Name = "modPhysEval"
' ===== modPhysEval.bas =====
Option Explicit



'=== 繝ｬ繧､繧｢繧ｦ繝亥・騾・===
Private Const PAD_X      As Single = 12   ' 竊・縺吶〒縺ｫ菴ｿ縺｣縺ｦ縺・ｋ蜈ｨ菴謎ｽ咏區
Private Const PAD_Y      As Single = 10
Private Const GAP_Y      As Single = 6
Private Const NOTE_H     As Single = 22
Private Const COL1_W     As Single = 140  ' 蟾ｦ縺ｮ隕句・縺怜・蟷・ｼ井ｻ也判髱｢縺ｧ繧ゆｽｿ逕ｨ・・
Private Const COL_EDT_W  As Single = 70   ' 莉也判髱｢逕ｨ縺ｮ蜈･蜉帛ｹ・
Private Const ROW_H     As Single = 18
Private Const ROM_LABEL_W As Single = 60

' === ROM 蟄舌ち繝也畑 霑ｽ蜉螳壽焚・上ヵ繝ｩ繧ｰ ===
Private Const ROM_JOINT_GAP_Y As Single = 2   ' 髢｢遽繝悶Ο繝・け髢薙・邵ｦ髢・
Private Const ROM_MOTION_GAP_Y As Single = 4   ' 驕句虚陦後・邵ｦ髢・
Private Const ROM_GROUP_PAD    As Single = 8   ' Frame 蜀・ヱ繝・ぅ繝ｳ繧ｰ
Private Const ROM_HDR_RL_GAP   As Single = 24  ' 驕句虚蜷阪→R/L蛻励・髢・
Private Const ROM_COL_SHIFT_L  As Single = 34
Private Const ROM_FRAME_TRIM_R As Single = 24
Private Const ROM_MULTI_COL_GAP As Single = 16
Private Const ROM_MULTI_EDIT_GAP As Single = 6


Public Const USE_ROM_SUBTABS   As Boolean = True   ' 蟄舌ち繝・荳願い/荳玖い)繧剃ｽｿ縺・


'=== ROM繝ｬ繧､繧｢繧ｦ繝茨ｼ医％縺ｮ4縺､縺ｯROM蟆ら畑縺ｧ1蝗槭□縺大ｮ夂ｾｩ・・==
Private Const ROM_ROW_H      As Single = 22
Private Const ROM_GAP_Y      As Single = 1
Private Const ROM_HDR_GAP    As Single = 2
Private Const ROM_COL_EDT_W  As Single = 50
Private Const ROM_TXT_H      As Single = 18

'=== 蛯呵・ｬ・・蜈ｱ騾壹ヱ繝ｩ繝｡繝ｼ繧ｿ・域眠隕擾ｼ・===
Private Const MEMO_DESIRED_H As Single = 120   ' 縺ｻ縺励＞鬮倥＆・・00?160縺ｧ螂ｽ縺ｿ縺ｫ隱ｿ謨ｴ蜿ｯ・・
Private Const MEMO_MIN_H     As Single = 72    ' 譛菴朱ｫ倥＆

'=== 蛯呵・螟牙ｽ｢繝・く繧ｹ繝育畑 螳壽焚・・odPhysEval蜀・〒蜈ｱ騾夲ｼ・===
Private Const NOTE_W_RATE    As Single = 0.6   ' 蛯呵・・繝・け繧ｹ縺ｮ讓ｪ蟷・= 蛻ｩ逕ｨ蜿ｯ閭ｽ蟷・・60%
Private Const DEFORM_W_RATE  As Single = 0.6   ' 螟牙ｽ｢繝・く繧ｹ繝医・讓ｪ蟷・= 蛻ｩ逕ｨ蜿ｯ閭ｽ蟷・・60%
Private Const DEFORM_H       As Single = 120   ' 螟牙ｽ｢繝・く繧ｹ繝医・鬮倥＆・・x逶ｸ蠖難ｼ・

Private Const CAP_FUNC_REFLEX As String = "遲狗ｷ雁ｼｵ繝ｻ蜿榊ｰ・ｼ育吏邵ｮ蜷ｫ繧・・
Private Const CAP_FUNC_PAIN   As String = "逍ｼ逞幢ｼ磯Κ菴搾ｼ蒐RS・・





Public Sub PlaceMemoBelow( _
    host As MSForms.Frame, _
    ByVal w As Single, ByVal h As Single, _
    ByVal yTop As Single, _
    ByVal memoName As String, _
    Optional ByVal fr1 As MSForms.Frame, _
    Optional ByVal fr2 As MSForms.Frame, _
    Optional ByVal labelText As String = "蛯呵・ｬ・)

    Const GAP_X As Single = 8
    Const GAP_Y As Single = 4

    ' 譌｢蟄倥・繝｡繝｢縺ｨ繝ｩ繝吶Ν繧帝勁蜴ｻ  竊・縺薙％縲∝濠隗偵・ " 縺ｫ・・
    On Error Resume Next
    If ControlExists(host, memoName) Then host.controls.Remove memoName

    Dim i As Long
    For i = host.controls.count - 1 To 0 Step -1
        If host.controls(i).name = memoName & "_lbl" Then
        host.controls.Remove i
        Exit For
    End If
Next i

    On Error GoTo 0
' --- 荳狗ｫｯ縺ｮ蝓ｺ貅悶ｒ豎ｺ繧√ｋ・・r1/fr2 縺ｮ豺ｱ縺・婿 + PAD_Y・・--
Dim yBottom As Single
yBottom = yTop
If Not fr1 Is Nothing Then yBottom = Application.WorksheetFunction.Max(yBottom, fr1.Top + fr1.Height + PAD_Y)
If Not fr2 Is Nothing Then yBottom = Application.WorksheetFunction.Max(yBottom, fr2.Top + fr2.Height + PAD_Y)

' 繝｡繝｢鬆伜沺縺ｮ髢句ｧ倶ｽ咲ｽｮ・医Λ繝吶Ν縺ｮTop・峨ｒ豎ｺ螳・
Dim memoTop As Single, safeTopMax As Single
memoTop = yBottom
If memoTop < yBottom Then memoTop = yBottom

' 繝ｩ繝吶ΝTop縺ｮ譛螟ｧ險ｱ螳ｹ・・ 谿九ｊ縺・ROW_H + GAP_Y + MEMO_MIN_H 縺ｯ遒ｺ菫昴〒縺阪ｋ菴咲ｽｮ・・
safeTopMax = h - PAD_Y - (ROW_H + GAP_Y + MEMO_MIN_H)
If safeTopMax < PAD_Y Then safeTopMax = PAD_Y     ' 繝輔Ξ繝ｼ繝縺梧･ｵ遶ｯ縺ｫ菴弱＞蝣ｴ蜷医・菫晞匱
If memoTop > safeTopMax Then memoTop = safeTopMax

' 隕句・縺励Λ繝吶Ν
Dim lbl As MSForms.label
Set lbl = host.controls.Add("Forms.Label.1", memoName & "_lbl")
With lbl
    .caption = labelText
    .Left = PAD_X
    .Top = memoTop
    .Width = w - PAD_X * 2
    .Height = ROW_H
    .Font.Bold = False
End With


' 蝗ｺ螳壻ｸ句ｯ・○縺ｯ繧・ａ繧具ｼ壼ｙ閠・ｬ・・縲瑚ｩ穂ｾ｡鬆・岼縺ｮ逶ｴ荳・yBottom)縲阪↓鄂ｮ縺阪∫ｸｦ繧ｵ繧､繧ｺ縺ｯ MEMO_DESIRED_H 繧剃ｸ企剞縺ｫ縺励※荳九↓莨ｸ縺ｳ縺吶℃縺ｪ縺・ｈ縺・↓縺吶ｋ・・026-01・・
' 繝・く繧ｹ繝医・繝・け繧ｹ譛ｬ菴・
Dim txt As MSForms.TextBox, hCalc As Single
Set txt = host.controls.Add("Forms.TextBox.1", memoName)
With txt
    .Left = PAD_X
    .Top = lbl.Top + ROW_H
    .Width = w - PAD_X * 2

    ' 谿九ｊ鬮倥＆繧定ｨ育ｮ・竊・譛菴朱ｫ倥＆・・莉･荳翫↓荳ｸ繧√※縺九ｉ險ｭ螳・
    hCalc = Application.WorksheetFunction.Min(MEMO_DESIRED_H, h - PAD_Y - .Top)
    If hCalc < MEMO_MIN_H Then hCalc = MEMO_MIN_H
    If hCalc < 1 Then hCalc = 1
    .Height = PX(hCalc)

    .multiline = True
    .WordWrap = True
    .EnterKeyBehavior = True
    .ScrollBars = fmScrollBarsVertical
End With

' 繝ｩ繝吶Ν縺ｮ逶ｴ蜑阪〒蟾ｦ蜿ｳ繧ｫ繝ｩ繝縺ｮ繝輔Ξ繝ｼ繝繧呈ｭ｢繧√ｋ
If Not fr1 Is Nothing Then fr1.Height = lbl.Top - PAD_Y
If Not fr2 Is Nothing Then fr2.Height = lbl.Top - PAD_Y

' 縺薙・繝壹・繧ｸ縺ｧ縺ｯ繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ遖∵ｭ｢
host.ScrollBars = fmScrollBarsNone
host.ScrollHeight = host.Height

End Sub





'========================================
' 蜈ｬ髢帰PI・夊ｺｫ菴捺ｩ溯・隧穂ｾ｡繧ｿ繝紋ｸ蠑上ｒ菴懈・
'========================================
Public Sub EnsurePhysicalFunctionTabs(owner As frmEval)
    Dim mp As MSForms.MultiPage: Set mp = EnsurePhysMulti(owner)
    If mp Is Nothing Then Exit Sub

    Dim pgRom As MSForms.page, pgMMT As MSForms.page, pgSens As MSForms.page, _
    pgReflex As MSForms.page, pgPain As MSForms.page, pgNote As MSForms.page

    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)
    Set pgSens = FindOrAddPage(mp, CAP_FUNC_SENS_REF)
    Set pgNote = FindOrAddPage(mp, CAP_FUNC_NOTE)
    Set pgReflex = FindOrAddPage(mp, CAP_FUNC_REFLEX)
    Set pgPain = FindOrAddPage(mp, CAP_FUNC_PAIN)
    

    ' 蜷・・繝ｼ繧ｸ縺ｫ繝帙せ繝医ヵ繝ｬ繝ｼ繝繧堤畑諢・
    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame, hostSens As MSForms.Frame, _
    hostReflex As MSForms.Frame, hostPain As MSForms.Frame, hostNote As MSForms.Frame

    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)
    Set hostReflex = EnsureHostFrame(pgReflex)   ' 竊・霑ｽ蜉
    Set hostPain = EnsureHostFrame(pgPain)       ' 竊・霑ｽ蜉
    
    Set hostNote = EnsureHostFrame(pgNote)

Dim pgPar As MSForms.page, hostPar As MSForms.Frame
Set pgPar = FindOrAddPage(mp, CAP_FUNC_PARALYSIS)
Set hostPar = EnsureHostFrame(pgPar)
BuildParalysisTabUI hostPar   ' 竊・譌｢縺ｫ雋ｼ縺｣縺滄ｺｻ逞ｺUI繝薙Ν繝




    ' 繝薙Ν繝会ｼ・I逕滓・・・
    If USE_ROM_SUBTABS Then
    BuildROMTabs hostRom         ' 竊・譁ｰ・壻ｸ願い・丈ｸ玖い縺ｮ蟄舌ち繝・
Else
    BuildROMSection_Compact hostRom   ' 竊・譌｢蟄假ｼ壻ｺ悟・繝ｬ繧､繧｢繧ｦ繝茨ｼ井ｺ呈鋤逕ｨ・・
End If

    If Not UseMMTChildTabs() Then BuildMMTSection owner, hostMmt
   BuildSensoryTabUI hostSens
BuildToneReflexTabUI hostReflex
BuildPainTabUI owner, hostPain

    AddNotesBox owner, hostNote, TAG_FUNC_PREFIX

    ' 蛻晄悄陦ｨ遉ｺ縺ｯROM
    mp.value = pgRom.Index
End Sub

'========================================
' 蜀・Κ・哺ultiPage縺ｮ逕ｨ諢擾ｼ・ostBody蜀・ｼ・
'========================================
Private Function EnsurePhysMulti(owner As frmEval) As MSForms.MultiPage
    Dim host As MSForms.Frame: Set host = FindHostByName(owner, HOST_BODY_NAME)
    If host Is Nothing Then
        MsgBox "繝輔Ξ繝ｼ繝 '" & HOST_BODY_NAME & "' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲ょ・縺ｫ Validate_App 繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲・, vbExclamation
        Exit Function
    End If

    Dim c As Control
    For Each c In host.controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then
                Set EnsurePhysMulti = c
                Exit Function
            End If
        End If
    Next

    Set EnsurePhysMulti = host.controls.Add("Forms.MultiPage.1")
    With EnsurePhysMulti
        .name = MP_PHYS_NAME
        .Left = PAD_X
        .Top = PAD_Y
        .Width = host.Width - PAD_X * 2
        .Height = host.Height - PAD_Y * 2
    End With

    ' 繧ｿ繝門・譖ｿ繝輔ャ繧ｯ・亥ｰ・擂縺ｮIME蜀埼←逕ｨ遲峨↓蛯吶∴・・
    On Error Resume Next
    Dim mph As New MPHook
    mph.Init owner, EnsurePhysMulti
    owner.RegisterMPHook mph
    On Error GoTo 0
End Function

Private Function FindHostByName(frm As frmEval, hostName As String) As MSForms.Frame
    Dim c As Control
    For Each c In frm.controls
        If TypeOf c Is MSForms.Frame Then
            If c.name = hostName Then Set FindHostByName = c: Exit Function
        End If
    Next
End Function

Private Function FindOrAddPage(mp As MSForms.MultiPage, captionText As String) As MSForms.page
    Dim i As Long
    For i = 0 To mp.Pages.count - 1
        If mp.Pages(i).caption = captionText Then
            Set FindOrAddPage = mp.Pages(i)
            Exit Function
        End If
    Next
    Set FindOrAddPage = mp.Pages.Add
    FindOrAddPage.caption = captionText
End Function

Public Function EnsureHostFrame(pg As MSForms.page) As MSForms.Frame

    Dim c As Control
    For Each c In pg.controls
        If TypeOf c Is MSForms.Frame Then
            Set EnsureHostFrame = c
            Exit Function
        End If
    Next

    ' 笘・縺薙％縺九ｉ荳九・縲悟・蝗槭・縺ｿ縲・
    Dim f As MSForms.Frame
    Set f = pg.controls.Add("Forms.Frame.1")
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
' MMT・井ｸｻ隕∫ｭ狗ｾ､繝ｻ蟾ｦ蜿ｳ・・
'========================================
Private Sub BuildMMTSection(owner As frmEval, host As MSForms.Frame)
    Dim y As Single: y = PAD_Y

    Dim groups As Variant
    ' 莉｣陦ｨ遲狗ｾ､・育ｰ｡貎費ｼ会ｼ夊か螟冶ｻ｢・剰ｘ螻域峇・乗焔髢｢遽閭悟ｱ茨ｼ剰ぃ螻域峇・剰・莨ｸ螻包ｼ剰ｶｳ閭悟ｱ・
    groups = Array("閧ｩ螟冶ｻ｢", "閧伜ｱ域峇", "謇矩未遽閭悟ｱ・, "閧｡螻域峇", "閹昜ｼｸ螻・, "雜ｳ閭悟ｱ・)

    ' 隕句・縺・
    y = AddHeaderRow(host, "遲狗ｾ､", y)

    Dim i As Long
    For i = LBound(groups) To UBound(groups)
        y = AddMMTRow(owner, host, CStr(groups(i)), y)
    Next i
    
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, y + ROM_HDR_GAP, "txtMMTMemo")
    
End Sub

Private Function AddHeaderRow(host As MSForms.Frame, title As String, y As Single) As Single
    Dim hTitle As MSForms.label, hR As MSForms.label, hL As MSForms.label
    Set hTitle = host.controls.Add("Forms.Label.1")
    With hTitle: .caption = title: .Left = PAD_X: .Top = y: .Width = COL1_W: .Height = ROW_H: .Font.Bold = True: End With
    Set hR = host.controls.Add("Forms.Label.1")
    With hR: .caption = "蜿ｳ": .Left = PAD_X + COL1_W + 8: .Top = y: .Width = COL_EDT_W: .Height = ROW_H: .TextAlign = fmTextAlignCenter: .Font.Bold = True: End With
    Set hL = host.controls.Add("Forms.Label.1")
    With hL: .caption = "蟾ｦ": .Left = hR.Left + COL_EDT_W + 8: .Top = y: .Width = COL_EDT_W: .Height = ROW_H: .TextAlign = fmTextAlignCenter: .Font.Bold = True: End With
    AddHeaderRow = y + ROW_H + 2
End Function

Private Function AddMMTRow(owner As frmEval, host As MSForms.Frame, muscle As String, y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl: .caption = muscle: .Left = PAD_X: .Top = y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.controls.Add("Forms.ComboBox.1")
    Set cboL = host.controls.Add("Forms.ComboBox.1")

    SetupMMTCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = y
    SetupMMTCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = y

    cboR.tag = TAG_FUNC_PREFIX & "|MMT_" & muscle & "_蜿ｳ"
    cboL.tag = TAG_FUNC_PREFIX & "|MMT_" & muscle & "_蟾ｦ"

    ' ・・I蜷郁ｨ医→縺ｯ辟｡髢｢菫ゅ↑縺ｮ縺ｧ CboBIHook 縺ｯ譛ｪ驕ｩ逕ｨ・・
    AddMMTRow = y + ROW_H + GAP_Y
End Function

Private Sub SetupMMTCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        ' 荳闊ｬ逧・↑MMT陦ｨ迴ｾ・・~5・仰ｱ縲√♀繧医・縲御ｸ榊庄縲・
        .AddItem "0"
        .AddItem "1"
        .AddItem "2-": .AddItem "2": .AddItem "2+"
        .AddItem "3-": .AddItem "3": .AddItem "3+"
        .AddItem "4-": .AddItem "4": .AddItem "4+"
        .AddItem "5"
        .AddItem "荳榊庄"
    End With
End Sub


Private Function AddSensoryRow(host As MSForms.Frame, itemKey As String, y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl: .caption = Replace(itemKey, "_", " / "): .Left = PAD_X: .Top = y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.controls.Add("Forms.ComboBox.1")
    Set cboL = host.controls.Add("Forms.ComboBox.1")
    SetupSensoryCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = y
    SetupSensoryCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = y

    cboR.tag = TAG_FUNC_PREFIX & "|SENS_" & itemKey & "_蜿ｳ"
    cboL.tag = TAG_FUNC_PREFIX & "|SENS_" & itemKey & "_蟾ｦ"

    AddSensoryRow = y + ROW_H + GAP_Y
End Function

Private Sub SetupSensoryCombo(cbo As MSForms.ComboBox)
    With cbo
        .Width = COL_EDT_W: .Height = ROW_H: .Style = fmStyleDropDownList
        .Clear
        .AddItem "豁｣蟶ｸ"
        .AddItem "菴惹ｸ・
        .AddItem "豸亥､ｱ"
        .AddItem "譛ｪ讀・
        .AddItem "荳榊庄"
    End With
End Sub

Private Function AddMASRow(host As MSForms.Frame, groupName As String, y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl: .caption = groupName: .Left = PAD_X: .Top = y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.controls.Add("Forms.ComboBox.1")
    Set cboL = host.controls.Add("Forms.ComboBox.1")
    SetupMASCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = y
    SetupMASCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = y

    cboR.tag = TAG_FUNC_PREFIX & "|TONE_MAS_" & groupName & "_蜿ｳ"
    cboL.tag = TAG_FUNC_PREFIX & "|TONE_MAS_" & groupName & "_蟾ｦ"

    AddMASRow = y + ROW_H + GAP_Y
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
        .AddItem "譛ｪ讀・
    End With
End Sub

Private Function AddReflexRow(host As MSForms.Frame, reflexName As String, y As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl: .caption = reflexName: .Left = PAD_X: .Top = y: .Width = COL1_W: .Height = ROW_H: End With

    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox
    Set cboR = host.controls.Add("Forms.ComboBox.1")
    Set cboL = host.controls.Add("Forms.ComboBox.1")
    SetupReflexCombo cboR: cboR.Left = PAD_X + COL1_W + 8: cboR.Top = y
    SetupReflexCombo cboL: cboL.Left = cboR.Left + COL_EDT_W + 8: cboL.Top = y

    cboR.tag = TAG_FUNC_PREFIX & "|REFLEX_" & reflexName & "_蜿ｳ"
    cboL.tag = TAG_FUNC_PREFIX & "|REFLEX_" & reflexName & "_蟾ｦ"

    AddReflexRow = y + ROW_H + GAP_Y
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
        .AddItem "譛ｪ讀・
    End With
End Sub

Private Function AddDeformText(owner As frmEval, host As MSForms.Frame, y As Single) As Single
    Dim lbl As MSForms.label, txt As MSForms.TextBox

    ' 繝ｩ繝吶Ν・遺・莉悶→蜷後§蛻怜ｹ・ｒ菴ｿ縺・ｼ・
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = "螟牙ｽ｢・域園隕具ｼ・
        .Left = PAD_X
        .Top = y
        .Width = COL1_W            ' 笘・ｵｱ荳・・
        .Height = ROW_H
    End With

    ' 繝・く繧ｹ繝茨ｼ遺・蜈･蜉帛・縺ｮ髢句ｧ倶ｽ咲ｽｮ縺ｫ繧ｹ繝翫ャ繝暦ｼ・
    Set txt = host.controls.Add("Forms.TextBox.1")
    With txt
        .Left = PAD_X + COL1_W + 8      ' 笘・ｵｱ荳・・
        .Top = y
        ' 讓ｪ蟷・→鬮倥＆縺ｯ蜑榊屓縺ｮ險ｭ螳壹ｒ豬∫畑・亥､縺ｯ縺ゅ↑縺溘・螂ｽ縺ｿ縺ｧ・・
        Dim availW As Single
        availW = host.Width - .Left - PAD_X
        .Width = availW * DEFORM_W_RATE
        .Height = DEFORM_H

        .multiline = True
        .EnterKeyBehavior = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .tag = TAG_FUNC_PREFIX & "|PAIN|螟牙ｽ｢_謇隕・
    End With

    ' IME hook・育怐逡･蜿ｯ・・
    On Error Resume Next
    Dim ime As New TxtImeHook
    ime.Init txt: owner.RegisterTxtHook ime
    On Error GoTo 0

    ' 谺｡縺ｮY
    AddDeformText = txt.Top + txt.Height + GAP_Y
End Function



Private Function AddPainRow(owner As frmEval, host As MSForms.Frame, y As Single) As Single
    ' 霑ｽ蜉縺ｯ蜈ｨ驛ｨ繝ｭ繝ｼ繧ｫ繝ｫ螳壽焚・壼､夜Κ縺ｫ萓晏ｭ倥＠縺ｪ縺・
    Const NRS_LBL_W As Single = 28
    Const NRS_CBO_W As Single = 60
    Const GAP_X     As Single = 8

    ' 繝ｩ繝吶Ν縲檎名逞幢ｼ磯Κ菴搾ｼ峨・
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = "逍ｼ逞幢ｼ磯Κ菴搾ｼ・
        .Left = PAD_X: .Top = y
        .Width = COL1_W: .Height = ROW_H
    End With

    ' 蜿ｳ蛛ｴ縺ｫ NRS・医Λ繝吶Ν・九さ繝ｳ繝懶ｼ峨ｒ蜈医↓驟咲ｽｮ縺励※蝓ｺ貅悶↓縺吶ｋ
    Dim lblN As MSForms.label, cbo As MSForms.ComboBox
    Set lblN = host.controls.Add("Forms.Label.1")
    With lblN
        .caption = "NRS"
        .Top = y: .Width = NRS_LBL_W: .Height = ROW_H
    End With

    Set cbo = host.controls.Add("Forms.ComboBox.1")
    With cbo
        .Top = y: .Width = NRS_CBO_W: .Height = ROW_H
        .Style = fmStyleDropDownList
        .tag = TAG_FUNC_PREFIX & "|PAIN_NRS"
        ' 蠢・ｦ√↑繧・0・・0 繧定・蜍輔〒蝓九ａ繧具ｼ域里縺ｫ險ｭ螳壹＠縺ｦ縺・ｋ縺ｪ繧我ｽ輔ｂ縺励↑縺・ｼ・
        If .ListCount = 0 Then
            Dim i As Integer
            For i = 0 To 10: .AddItem CStr(i): Next i
        End If
    End With

    ' 蜿ｳ遶ｯ縺ｫ謠・∴繧・
    Dim rightEdge As Single: rightEdge = host.InsideWidth - PAD_X
    cbo.Left = rightEdge - NRS_CBO_W
    lblN.Left = cbo.Left - GAP_X - NRS_LBL_W

    ' 驛ｨ菴阪ユ繧ｭ繧ｹ繝医・縲悟・蜉帛・髢句ｧ九阪°繧・NRS 謇句燕縺ｾ縺ｧ繧定・蜍募ｹ・〒
    Dim txt As MSForms.TextBox
    Set txt = host.controls.Add("Forms.TextBox.1")
    With txt
        .Top = y
        .Left = PAD_X + COL1_W + GAP_X
        .Width = lblN.Left - GAP_X - .Left
        If .Width < 80 Then .Width = 80        ' 螳牙・蠑・
        .Height = ROW_H
        .tag = TAG_FUNC_PREFIX & "|PAIN_驛ｨ菴・
        .EnterKeyBehavior = False
        .multiline = False
        ' 譌･譛ｬ隱槫・蜉帙・縺ｾ縺ｾ縺ｧOK縲ょ濠隗貞崋螳壹↑繧・ .IMEMode = fmIMEModeDisable
    End With

    AddPainRow = y + ROW_H + GAP_Y
End Function


'========================================
' 蛯呵・ｼ郁・逕ｱ險倩ｿｰ・俄ｻIME縺ｲ繧峨′縺ｪON
'========================================

Private Sub AddNotesBox(owner As frmEval, host As MSForms.Frame, keyPrefix As String)
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = "蛯呵・ｼ郁・逕ｱ險倩ｿｰ・・
        .Left = PAD_X
        .Width = 120
        .Height = 18
    End With

    Dim txt As MSForms.TextBox
    Set txt = host.controls.Add("Forms.TextBox.1")
    With txt
        .Left = lbl.Left + lbl.Width + 8
         Dim availW As Single
    availW = host.Width - .Left - PAD_X
    .Width = availW * NOTE_W_RATE
        .Height = NOTE_H
        .multiline = True
        .EnterKeyBehavior = True
        .tag = keyPrefix & "|蛯呵・
    End With

    ' 笘・％縺薙′繝昴う繝ｳ繝茨ｼ夂樟蝨ｨ縺ｮ蜀・ｮｹ縺ｮ逶ｴ荳九↓荳ｦ縺ｹ繧・
    Dim bottom As Single
    bottom = GetContentBottom(host)
    lbl.Top = bottom + 8
    txt.Top = lbl.Top

    ' 繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ邨らｫｯ繧貞ｙ閠・・荳九∪縺ｧ莨ｸ縺ｰ縺・
    host.ScrollBars = fmScrollBarsVertical
    host.ScrollHeight = txt.Top + txt.Height + PAD_Y

    ' IME・亥ｙ閠・・譌･譛ｬ隱樊Φ螳壹↑縺ｮ縺ｧ On 縺ｮ縺ｾ縺ｾ縲ゅ・繧峨′縺ｪ Hook 縺ｯ譌｢蟄倬壹ｊ・・
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
    For Each c In parent.controls
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
            Dim m As MSForms.MultiPage
            Set m = FindMultiPageRecursive(c)
            If Not m Is Nothing Then
                Set FindMultiPageRecursive = m
                Exit Function
            End If
        End If
    Next
End Function


' mpADL 蛻､螳夲ｼ亥錐蜑・or 繝壹・繧ｸ隕句・縺励〒蛻､螳夲ｼ・
Private Function IsKnownMpADL(mp As MSForms.MultiPage) As Boolean
    On Error Resume Next
    If LCase$(mp.name) = "mpadl" Then
        IsKnownMpADL = True
        Exit Function
    End If
    If mp.Pages.count >= 3 Then
        Dim c0$, c1$, c2$
        c0 = mp.Pages(0).caption
        c1 = mp.Pages(1).caption
        c2 = mp.Pages(2).caption
        If (InStr(c0, "BI") > 0 Or InStr(c0, "繝舌・繧ｵ繝ｫ") > 0) _
        And (InStr(c1, "IADL") > 0) _
        And (InStr(c2, "襍ｷ螻・) > 0) Then
            IsKnownMpADL = True
        End If
    End If
End Function


'=== 霑ｽ蜉・夊ｺｫ菴捺ｩ溯・隧穂ｾ｡繧偵瑚ｦｪ繧ｿ繝悶搾ｼ医Ν繝ｼ繝茨ｼ峨→縺励※菴懈・・亥ｙ閠・・蜷・ｭ舌ち繝悶↓驟咲ｽｮ・・===



Public Sub EnsurePhysicalFunctionTabs_Root(owner As frmEval)
    Debug.Print "[phys] enter EnsurePhysicalFunctionTabs_Root"

    Dim root As MSForms.MultiPage: Set root = FindRootMulti(owner)
    If root Is Nothing Then
        Debug.Print "[phys] root not found"
        ' 隕九▽縺九ｉ縺ｪ縺・→縺阪・迥ｶ豕√ｒ繝繝ｳ繝・
        Call DumpMultiPages(owner)
        MsgBox "譛荳頑ｮｵ縺ｮMultiPage縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲ゅう繝溘ョ繧｣繧ｨ繧､繝・CTRL+G)縺ｮ繝ｭ繧ｰ繧呈蕗縺医※縺上□縺輔＞縲・
        Exit Sub
    Else
        Debug.Print "[phys] root found: Name=" & root.name & ", Pages=" & root.Pages.count
    End If
   


    ' 繝ｫ繝ｼ繝医↓縲瑚ｺｫ菴捺ｩ溯・隧穂ｾ｡縲阪・繝ｼ繧ｸ繧定ｿｽ蜉/蜿門ｾ・
    Dim pgPhys As MSForms.page
    Set pgPhys = FindOrAddPage(root, CAP_FUNC)

    ' 繝壹・繧ｸ蜀・ヵ繝ｬ繝ｼ繝
    Dim host As MSForms.Frame
    Set host = EnsureHostFrame(pgPhys)

    ' 繝壹・繧ｸ蜀・↓窶懷ｭ舌ち繝也畑窶昴・MultiPage・・pPhys・峨ｒ霑ｽ蜉/蜿門ｾ・
    Dim mp As MSForms.MultiPage
    Dim c As Control
    For Each c In host.controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then Set mp = c: Exit For
        End If
    Next
    If mp Is Nothing Then
        Set mp = host.controls.Add("Forms.MultiPage.1")
        With mp
            .name = MP_PHYS_NAME
            .Left = PAD_X
            .Top = PAD_Y
            .Width = host.Width - PAD_X * 2
            .Height = host.Height - PAD_Y * 2
        End With
        ' 繧ｿ繝門・譖ｿ繝輔ャ繧ｯ
        On Error Resume Next
        Dim mph As New MPHook
        mph.Init owner, mp
        owner.RegisterMPHook mph
        On Error GoTo 0
    End If

    ' 蟄舌ち繝厄ｼ・譫夲ｼ会ｼ啌OM / MMT / 諢溯ｦ壹・逞咏ｸｮ繝ｻ蜿榊ｰ・・逍ｼ逞幢ｼ亥推繧ｿ繝悶↓蛯呵・≠繧奇ｼ・
    Dim pgRom As MSForms.page, pgMMT As MSForms.page, pgSens As MSForms.page
    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)
    Set pgSens = FindOrAddPage(mp, CAP_FUNC_SENS_REF)

    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame, hostSens As MSForms.Frame
    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)

    ' UI逕滓・・句推繧ｿ繝悶↓蛯呵・
    BuildROMSection_Compact hostRom
    
    If Not UseMMTChildTabs() Then BuildMMTSection owner, hostMmt

    BuildSensoryToneReflexPain owner, hostSens
    

    ' 蛻晄悄陦ｨ遉ｺ
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
    For Each c In parent.controls
        If TypeOf c Is MSForms.MultiPage Then
            On Error Resume Next
            Debug.Print pad & "MP Name=" & c.name & " Pages=" & c.Pages.count
            For i = 0 To c.Pages.count - 1
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


'=== 霑ｽ蜉・壽欠螳壹＆繧後◆繝ｫ繝ｼ繝・MultiPage 縺ｫ縲瑚ｺｫ菴捺ｩ溯・隧穂ｾ｡縲阪・繝ｼ繧ｸ繧剃ｽ懊ｋ ===
Public Sub EnsurePhysicalFunctionTabs_Under(owner As frmEval, root As MSForms.MultiPage)
    If root Is Nothing Then
        MsgBox "繝ｫ繝ｼ繝・MultiPage 縺・Nothing 縺ｧ縺吶ょ他縺ｳ蜃ｺ縺怜・縺ｧ 'mp' 繧呈ｸ｡縺励※縺上□縺輔＞縲・, vbExclamation
        Exit Sub
    End If

    ' 繝ｫ繝ｼ繝医↓縲瑚ｺｫ菴捺ｩ溯・隧穂ｾ｡縲阪・繝ｼ繧ｸ繧定ｿｽ蜉/蜿門ｾ・
    Dim pgPhys As MSForms.page
    Set pgPhys = FindOrAddPage(root, CAP_FUNC)

    ' 繝壹・繧ｸ蜀・・菴懈･ｭ繝輔Ξ繝ｼ繝
    Dim host As MSForms.Frame
    Set host = EnsureHostFrame(pgPhys)

        ' --- 蟄舌ち繝・MultiPage・・pPhys・峨ｒ菴懊ｋ or 蜿門ｾ・---
    Dim mp As MSForms.MultiPage
    Dim c As Control
    For Each c In host.controls
        If TypeOf c Is MSForms.MultiPage Then
            If c.name = MP_PHYS_NAME Then Set mp = c: Exit For
        End If
    Next
    If mp Is Nothing Then
        Set mp = host.controls.Add("Forms.MultiPage.1")
        With mp
            .name = MP_PHYS_NAME
            .Left = PAD_X
            .Top = PAD_Y
            .Width = host.Width - PAD_X * 2
            .Height = host.Height - PAD_Y * 2
        End With
    End If

       ' 笘・Page8/9縺ｪ縺ｩ縺ｮ譌｢螳壹・繝ｼ繧ｸ繧呈祉髯､
    CleanDefaultPages mp

    ' --- 蟄舌ち繝厄ｼ・譫夲ｼ峨ｒ逕ｨ諢・---
    Dim pgRom As MSForms.page, pgMMT As MSForms.page
    Dim pgSens As MSForms.page, pgToneRef As MSForms.page, pgPain As MSForms.page

    Set pgRom = FindOrAddPage(mp, CAP_FUNC_ROM)                     ' ROM・井ｸｻ隕・未遽・・
    Set pgMMT = FindOrAddPage(mp, CAP_FUNC_MMT)                     ' 遲句鴨・・MT・・
    Set pgSens = FindOrAddPage(mp, "諢溯ｦ夲ｼ郁｡ｨ蝨ｨ繝ｻ豺ｱ驛ｨ・・)              ' 竊仙・髮｢
    Set pgToneRef = FindOrAddPage(mp, "遲狗ｷ雁ｼｵ繝ｻ蜿榊ｰ・ｼ育吏邵ｮ蜷ｫ繧・・)     ' 竊仙・髮｢
    Set pgPain = FindOrAddPage(mp, "逍ｼ逞幢ｼ磯Κ菴搾ｼ蒐RS・・)               ' 竊仙・髮｢

    Dim hostRom As MSForms.Frame, hostMmt As MSForms.Frame
    Dim hostSens As MSForms.Frame, hostTone As MSForms.Frame, hostPain As MSForms.Frame

    Set hostRom = EnsureHostFrame(pgRom)
    Set hostMmt = EnsureHostFrame(pgMMT)
    Set hostSens = EnsureHostFrame(pgSens)
    Set hostTone = EnsureHostFrame(pgToneRef)
    Set hostPain = EnsureHostFrame(pgPain)
    
  ' --- Re-layout all mpPhys pages to actual size ---
    Dim iPg As Long
    Dim pgFit As MSForms.page
    Dim ctlFit As Control
    Dim fitW As Single, fitH As Single

    fitW = mp.Width - PAD_X * 2
    fitH = mp.Height - PAD_Y * 2
    If fitW < 120 Then fitW = 120
    If fitH < 80 Then fitH = 80

    For iPg = 0 To mp.Pages.count - 1
        Set pgFit = mp.Pages(iPg)

        ' Activate each page once to force page metrics refresh.
        mp.value = iPg

        ' Normalize direct host frames under each page.
        For Each ctlFit In pgFit.controls
            If TypeName(ctlFit) = "Frame" Then
                With ctlFit
                    .Left = PAD_X
                    .Top = PAD_Y
                    .Width = fitW
                    .Height = fitH
                End With
            End If
        Next ctlFit
    Next iPg
    
    
    
' ・・nsurePhysicalFunctionTabs_* 縺ｮ荳ｭ縲∽ｻ悶・pg・槭→蜷後§荳ｦ縺ｳ縺ｫ・・
Dim pgPar As MSForms.page, hostPar As MSForms.Frame
Set pgPar = FindOrAddPage(mp, CAP_FUNC_PARALYSIS)
Set hostPar = EnsureHostFrame(pgPar)
BuildParalysisTabUI hostPar



' --- UI讒狗ｯ・---
If USE_ROM_SUBTABS Then
    BuildROMTabs hostRom            ' 竊・譁ｰUI・井ｸ願い・丈ｸ玖い 蟄舌ち繝厄ｼ・
Else
    BuildROMSection_Compact hostRom ' 竊・譌ｧUI・域ｮ狗ｽｮ・・
End If

If Not UseMMTChildTabs() Then BuildMMTSection owner, hostMmt
BuildSensoryTabUI hostSens
BuildToneReflexTabUI hostTone
BuildPainTabUI owner, hostPain



    ' 蛻晄悄陦ｨ遉ｺ
    mp.value = pgRom.Index


End Sub




'=== 譌｢螳壹・繝ｼ繧ｸ "Page*" 繧呈祉髯､ ===
Private Sub CleanDefaultPages(mp As MSForms.MultiPage)
    On Error Resume Next
    Dim i As Long
    For i = mp.Pages.count - 1 To 0 Step -1
        If Left$(mp.Pages(i).caption, 4) = "Page" Then
            mp.Pages.Remove i
        End If
    Next
End Sub


'=== 諢溯ｦ壹・縺ｿ ===
Private Sub BuildSensoryTabUI(host As MSForms.Frame)
    Dim y As Single: y = PAD_Y

    y = AddHeaderRow(host, "諢溯ｦ夲ｼ郁｡ｨ蝨ｨ・・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_隗ｦ隕・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_逞幄ｦ・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_貂ｩ蠎ｦ隕・, y)

    y = y + ROM_HDR_GAP
    y = AddHeaderRow(host, "諢溯ｦ夲ｼ域ｷｱ驛ｨ・・, y)
    y = AddSensoryRow(host, "豺ｱ驛ｨ_菴咲ｽｮ隕・, y)
    y = AddSensoryRow(host, "豺ｱ驛ｨ_謖ｯ蜍戊ｦ・, y)

    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, y + ROM_HDR_GAP, "txtSensMemo")
End Sub







'=== 縺昴・繝輔Ξ繝ｼ繝蜀・〒荳逡ｪ荳九・菴咲ｽｮ・・op+Height・峨ｒ霑斐☆ ===
Private Function GetContentBottom(host As MSForms.Frame) As Single
    Dim c As Control, bottom As Single
    For Each c In host.controls
        If c.Visible Then
            If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next
    GetContentBottom = bottom
End Function




'=== ROM縺ｮ1陦鯉ｼ亥ｱ域峇/螟冶ｻ｢/窶ｦ・牙ｰ上＆繧∫沿・壼承繝ｻ蟾ｦ繝・く繧ｹ繝医・蜊願ｧ・===
Private Function AddROMRow_Compact( _
    host As MSForms.Frame, _
    jointName As String, _
    moveName As String, _
    y As Single _
) As Single

    Dim xR As Single, xL As Single: ROM_GetCols xR, xL
Dim yPix As Single: yPix = PX(y)

Dim txtR As MSForms.TextBox, txtL As MSForms.TextBox
Set txtR = host.controls.Add("Forms.TextBox.1")
With txtR
    .Left = PX(xR)
    .Top = yPix
    .Width = PX(ROM_COL_EDT_W)
    .Height = ROM_ROW_H
    .TextAlign = fmTextAlignCenter
    .IMEMode = fmIMEModeDisable
End With

Set txtL = host.controls.Add("Forms.TextBox.1")
With txtL
    .Left = PX(xL)
    .Top = yPix
    .Width = PX(ROM_COL_EDT_W)
    .Height = ROM_ROW_H
    .TextAlign = fmTextAlignCenter
    .IMEMode = fmIMEModeDisable
End With

txtR.tag = TAG_FUNC_PREFIX & "|ROM|" & jointName & "|" & moveName & "|蜿ｳ"
txtL.tag = TAG_FUNC_PREFIX & "|ROM|" & jointName & "|" & moveName & "|蟾ｦ"

AddROMRow_Compact = yPix + ROM_ROW_H + ROM_GAP_Y
End Function





' 蠕梧婿莠呈鋤繝ｩ繝・ヱ繝ｼ
Public Sub BuildROMSection_Compact(host As MSForms.Frame)
    BuildROMSection_TwoCols host
End Sub


' 隕句・縺暦ｼ医占か縲代↑縺ｩ・・
Private Function ROM_AddHeader(host As MSForms.Frame, title As String, y0 As Single) As Single
    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = "縲・ & title & "縲・
        .Left = PAD_X: .Top = y0: .Width = COL1_W: .Height = ROM_ROW_H
        .Font.Bold = True
    End With
    ROM_AddHeader = y0 + ROM_ROW_H + ROM_GAP_Y
End Function


'=== ROM compact 逕ｨ縺ｮ蜻ｼ縺ｳ蜃ｺ縺怜錐邨ｱ荳繝ｩ繝・ヱ繝ｼ =========================

' 隕句・縺暦ｼ医碁°蜍輔阪悟承(ﾂｰ)縲阪悟ｷｦ(ﾂｰ)縲阪・陦鯉ｼ・
Private Function ROM_AddDirHeader(host As MSForms.Frame, y0 As Single) As Single
    ROM_AddDirHeader = AddROMDirHeader_Compact(host, y0)
End Function

'=== ROM compact・啌 / L 隕句・縺暦ｼ・OM蟆ら畑蟷・〒・・===
Private Function AddROMDirHeader_Compact(host As MSForms.Frame, y0 As Single) As Single
    Dim xR As Single, xL As Single: ROM_GetCols xR, xL
    Dim y As Single: y = PX(y0)              ' 竊・y 繧ゆｸｸ繧√※縺翫￥・井ｻｻ諢上□縺代←螳牙ｮ壹＠縺ｾ縺呻ｼ・

    Dim lblR As MSForms.label, lblL As MSForms.label

    Set lblR = host.controls.Add("Forms.Label.1")
    With lblR
        .caption = "R"
        .Left = PX(xR)                        ' 竊・縺薙％縺ｫ PX
        .Top = y
        .Width = PX(ROM_COL_EDT_W)            ' 竊・縺薙％縺ｫ PX
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Set lblL = host.controls.Add("Forms.Label.1")
    With lblL
        .caption = "L"
        .Left = PX(xL)                        ' 竊・縺薙％縺ｫ PX
        .Top = y
        .Width = PX(ROM_COL_EDT_W)            ' 竊・縺薙％縺ｫ PX
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    AddROMDirHeader_Compact = y + ROM_ROW_H + ROM_GAP_Y
End Function




'========================================
' ROM・井ｺ悟・・壼ｷｦ=荳願い / 蜿ｳ=荳玖い・画悽菴・
'========================================
Public Sub BuildROMSection_TwoCols(host As MSForms.Frame)
    Const COL_GAP_X As Single = 12

    ' 譌｢蟄倬・鄂ｮ縺後≠繧後・髯､蜴ｻ・磯㍾隍・緒逕ｻ髦ｲ豁｢・・
    On Error Resume Next
    host.controls.Remove "fraROM_Upper"
    host.controls.Remove "fraROM_Lower"
    host.controls.Remove "txtROMMemo"
    On Error GoTo 0

    ' 竊舌％縺薙・谿九☆・唹n Error GoTo 0 縺ｮ逶ｴ蠕後°繧牙ｷｮ縺玲崛縺・

Dim w As Single, h As Single, colW As Single
w = host.InsideWidth: h = host.InsideHeight
' 蛻怜ｹ・ｒ謨ｴ謨ｰ縺ｫ荳ｸ繧√ｋ
colW = PX((w - (PAD_X * 2) - COL_GAP_X) / 2)

Dim frUL As MSForms.Frame, frLL As MSForms.Frame

' 蟾ｦ蛻励ヵ繝ｬ繝ｼ繝・井ｸ願い・・
Set frUL = host.controls.Add("Forms.Frame.1", "fraROM_Upper")
With frUL
    .caption = ""
    .Left = PX(PAD_X)                 ' 笘・紛謨ｰ荳ｸ繧・
    .Top = PAD_Y
    .Width = colW                     ' 笘・ｸｸ繧∵ｸ医∩蛻怜ｹ・
    .Height = h - PAD_Y * 2
    .ScrollBars = fmScrollBarsNone
    .ScrollHeight = .InsideHeight
End With

' 蜿ｳ蛻励ヵ繝ｬ繝ｼ繝・井ｸ玖い・・
Set frLL = host.controls.Add("Forms.Frame.1", "fraROM_Lower")
With frLL
    .caption = ""
    .Left = PX(PAD_X + colW + COL_GAP_X)  ' 笘・紛謨ｰ荳ｸ繧・
    .Top = PAD_Y
    .Width = colW                          ' 笘・ｸｸ繧∵ｸ医∩蛻怜ｹ・
    .Height = h - PAD_Y * 2
    .ScrollBars = fmScrollBarsNone
    .ScrollHeight = .InsideHeight
End With


    Dim yL As Single, yR As Single
    yL = PAD_Y: yR = PAD_Y

    ' ---------- 蟾ｦ蛻暦ｼ壻ｸ願い ----------
    yL = ROM_AddHeader(frUL, "荳願い", yL)
    yL = ROM_AddHeader(frUL, "縲占か髢｢遽縲・, yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "閧ｩ", "螻域峇", yL)
    yL = AddROMRow_Compact(frUL, "閧ｩ", "螟冶ｻ｢", yL)
    yL = AddROMRow_Compact(frUL, "閧ｩ", "螟匁雷", yL)

    yL = ROM_AddHeader(frUL, "縲占ｘ髢｢遽縲・, yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "閧・, "螻域峇", yL)
    yL = AddROMRow_Compact(frUL, "閧・, "莨ｸ螻・, yL)

    yL = ROM_AddHeader(frUL, "縲仙燕閻輔・, yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "蜑崎・", "蝗槫・", yL)
    yL = AddROMRow_Compact(frUL, "蜑崎・", "蝗槫､・, yL)

    yL = ROM_AddHeader(frUL, "縲先焔髢｢遽縲・, yL): yL = ROM_AddDirHeader(frUL, yL)
    yL = AddROMRow_Compact(frUL, "謇矩未遽", "謗悟ｱ・, yL)
    yL = AddROMRow_Compact(frUL, "謇矩未遽", "閭悟ｱ・, yL)

    ' ---------- 蜿ｳ蛻暦ｼ壻ｸ玖い ----------
    yR = ROM_AddHeader(frLL, "荳玖い", yR)
    yR = ROM_AddHeader(frLL, "縲占ぃ髢｢遽縲・, yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "閧｡", "螻域峇", yR)
    yR = AddROMRow_Compact(frLL, "閧｡", "螟冶ｻ｢", yR)

    yR = ROM_AddHeader(frLL, "縲占・髢｢遽縲・, yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "閹・, "螻域峇", yR)
    yR = AddROMRow_Compact(frLL, "閹・, "莨ｸ螻・, yR)

    yR = ROM_AddHeader(frLL, "縲占ｶｳ髢｢遽縲・, yR): yR = ROM_AddDirHeader(frLL, yR)
    yR = AddROMRow_Compact(frLL, "雜ｳ髢｢遽", "閭悟ｱ・, yR)
    yR = AddROMRow_Compact(frLL, "雜ｳ髢｢遽", "蠎募ｱ・, yR)

    ' 蛯呵・・蜈ｱ騾壹・繝ｫ繝代・縺ｫ邨ｱ荳・郁・蜍輔〒host繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ繧０FF・・
    Call PlaceMemoBelow(host, w, h, Application.WorksheetFunction.Max(yL, yR) + ROM_HDR_GAP, _
                        "txtROMMemo", frUL, frLL)
    ' 窶ｦ荳願い/荳玖い縺ｮ陦後ｒ縺吶∋縺ｦ菴懊ｊ邨ゅ∴縺溽峩蠕後↓窶ｦ
    NormalizeRomColumns frUL
    NormalizeRomColumns frLL

End Sub


'========================================
' 遲狗ｷ雁ｼｵ・・AS・会ｼ・蜿榊ｰ・ｼ医ン繝ｫ繝繝ｼ・・
' 蜻ｼ縺ｳ蜃ｺ縺嶺ｾ・ BuildToneReflexTabUI hostReflex
'========================================
Private Sub BuildToneReflexTabUI(host As MSForms.Frame)
    Dim y As Single: y = PAD_Y

    ' --- 遲狗ｷ雁ｼｵ・・AS・・---
    y = AddHeaderRow(host, "遲狗ｷ雁ｼｵ・・AS・・, y)
    y = AddMASRow(host, "荳願い螻育ｭ狗ｾ､", y)
    y = AddMASRow(host, "荳願い莨ｸ遲狗ｾ､", y)
    y = AddMASRow(host, "荳玖い螻育ｭ狗ｾ､", y)
    y = AddMASRow(host, "荳玖い莨ｸ遲狗ｾ､", y)

    ' --- 蜿榊ｰ・---
    y = y + ROM_HDR_GAP
    y = AddHeaderRow(host, "閻ｱ蜿榊ｰ・, y)
    y = AddReflexRow(host, "荳願・莠碁ｭ遲具ｼ・5-6・・, y)
    y = AddReflexRow(host, "荳願・荳蛾ｭ遲具ｼ・7・・, y)
    y = AddReflexRow(host, "閹晁搭閻ｱ・・2-4・・, y)
    y = AddReflexRow(host, "繧｢繧ｭ繝ｬ繧ｹ閻ｱ・・1・・, y)

    ' 蛯呵・ｼ井ｸ狗ｫｯ縺ｫ遒ｺ菫晢ｼ・
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, y + ROM_HDR_GAP, "txtReflexMemo")
End Sub

'========================================
' 逍ｼ逞幢ｼ亥､牙ｽ｢繝・く繧ｹ繝茨ｼ矩Κ菴搾ｼ起RS・峨ン繝ｫ繝繝ｼ
' 蜻ｼ縺ｳ蜃ｺ縺嶺ｾ・ BuildPainTabUI owner, hostPain
'========================================
Private Sub BuildPainTabUI(owner As frmEval, host As MSForms.Frame)

    If host Is Nothing Then Exit Sub

    Dim y As Single: y = PAD_Y

    'y = AddDeformText(owner, host, y)
    'y = AddPainRow(owner, host, y)

    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, y + ROM_HDR_GAP, "txtPainMemo")

End Sub

'========================================
' 莠呈鋤: 諢溯ｦ夲ｼ貴AS・句渚蟆・ｼ狗名逞・繧・繝壹・繧ｸ縺ｫ謠上￥・・oot逕ｨ・・
' 蜻ｼ縺ｳ蜃ｺ縺怜・: EnsurePhysicalFunctionTabs_Root
'========================================
Private Sub BuildSensoryToneReflexPain(owner As frmEval, host As MSForms.Frame)
    Dim y As Single: y = PAD_Y

    ' --- 諢溯ｦ夲ｼ郁｡ｨ蝨ｨ・・---
    y = AddHeaderRow(host, "諢溯ｦ夲ｼ郁｡ｨ蝨ｨ・・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_隗ｦ隕・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_逞幄ｦ・, y)
    y = AddSensoryRow(host, "陦ｨ蝨ｨ_貂ｩ蠎ｦ隕・, y)

    ' --- 諢溯ｦ夲ｼ域ｷｱ驛ｨ・・---
    y = y + ROM_HDR_GAP
    y = AddHeaderRow(host, "諢溯ｦ夲ｼ域ｷｱ驛ｨ・・, y)
    y = AddSensoryRow(host, "豺ｱ驛ｨ_髢｢遽菴咲ｽｮ隕・, y)
    y = AddSensoryRow(host, "豺ｱ驛ｨ_謖ｯ蜍戊ｦ・, y)

    ' --- 遲狗ｷ雁ｼｵ・・AS・・---
    y = y + ROM_HDR_GAP
    y = AddHeaderRow(host, "遲狗ｷ雁ｼｵ・・AS・・, y)
    y = AddMASRow(host, "荳願い螻育ｭ狗ｾ､", y)
    y = AddMASRow(host, "荳願い莨ｸ遲狗ｾ､", y)
    y = AddMASRow(host, "荳玖い螻育ｭ狗ｾ､", y)
    y = AddMASRow(host, "荳玖い莨ｸ遲狗ｾ､", y)

    ' --- 蜿榊ｰ・---
    y = y + ROM_HDR_GAP
    y = AddHeaderRow(host, "閻ｱ蜿榊ｰ・, y)
    y = AddReflexRow(host, "荳願・莠碁ｭ遲具ｼ・5-6・・, y)
    y = AddReflexRow(host, "荳願・荳蛾ｭ遲具ｼ・7・・, y)
    y = AddReflexRow(host, "閹晁搭閻ｱ・・2-4・・, y)
    y = AddReflexRow(host, "繧｢繧ｭ繝ｬ繧ｹ閻ｱ・・1・・, y)

    ' --- 螟牙ｽ｢・郁・逕ｱ繝・く繧ｹ繝茨ｼ・---
    y = y + ROM_HDR_GAP
    y = AddDeformText(owner, host, y)

    ' --- 逍ｼ逞幢ｼ磯Κ菴搾ｼ起RS・・---
    y = y + ROM_HDR_GAP
    y = AddPainRow(owner, host, y)

    ' 蛯呵・ｼ壹・繝ｼ繧ｸ荳狗ｫｯ縺ｫ遒ｺ菫・
    Call PlaceMemoBelow(host, host.InsideWidth, host.InsideHeight, y + ROM_HDR_GAP, "txtSensReflexPainMemo")
End Sub

' --- ROM 1陦悟・縺ｮ蜿ｳ(R)/蟾ｦ(L)縺ｮX蠎ｧ讓吶ｒ霑斐☆縺縺・---
Private Sub ROM_GetCols(ByRef xR As Single, ByRef xL As Single)
    Const GAP_X As Single = 8  ' 譌｢縺ｫ螳壽焚縺後≠繧後・縺薙・陦後・荳崎ｦ・
    xR = PX(PAD_X + COL1_W + GAP_X)            ' 蜿ｳ蜈･蜉帶ｬ・・蟾ｦ遶ｯ
    xL = PX(xR + ROM_COL_EDT_W + GAP_X)         ' 蟾ｦ蜈･蜉帶ｬ・・蟾ｦ遶ｯ
End Sub



' 蟆乗焚 竊・譛蟇・ｊ繝斐け繧ｻ繝ｫ縺ｫ繧ｹ繝翫ャ繝・
Public Function PX(v As Single) As Single
    PX = Int(v + 0.5)
End Function


' ROM繧ｳ繝ｩ繝縺ｮX菴咲ｽｮ/蟷・ｒ繝輔Ξ繝ｼ繝蜀・〒荳諡ｬ謨ｴ蛻・
Private Sub NormalizeRomColumns(host As MSForms.Frame)
    Dim xR As Single, xL As Single
    ROM_GetCols xR, xL
    xR = PX(xR): xL = PX(xL)
    Dim w As Single: w = PX(ROM_COL_EDT_W)

    Dim c As Control
    For Each c In host.controls
        Select Case TypeName(c)
            Case "TextBox"
                ' 繧ｿ繧ｰ縺ｫ縲悟承縲阪悟ｷｦ縲阪′蜈･繧九ｈ縺・↓縺励※縺翫￥・遺造蜿ら・・・
                If InStr(c.tag, "蜿ｳ") > 0 Then c.Left = xR: c.Width = w
                If InStr(c.tag, "蟾ｦ") > 0 Then c.Left = xL: c.Width = w
            Case "Label"
                If c.caption = "R" Then c.Left = xR: c.Width = w
                If c.caption = "L" Then c.Left = xL: c.Width = w
        End Select
    Next
End Sub

'========================================================
' 鮗ｻ逞ｺ繧ｿ繝・UI
'========================================================
Public Sub BuildParalysisTabUI(host As MSForms.Frame)
    Dim w As Single, h As Single, y As Single
    w = host.Width: h = host.Height
    y = PAD_Y

    ' ---- 蝓ｺ譛ｬ諠・ｱ ----
    y = AddSectionTitle(host, "蝓ｺ譛ｬ諠・ｱ", y)
    y = AddComboRow(host, "鮗ｻ逞ｺ蛛ｴ", "cboParalysisSide", Array("蜿ｳ", "蟾ｦ", "荳｡蛛ｴ"), y)
    y = AddComboRow(host, "鮗ｻ逞ｺ縺ｮ遞ｮ鬘・, "cboParalysisType", Array("迚・ｺｻ逞ｺ", "蝗幄い鮗ｻ逞ｺ", "蜊倬ｺｻ逞ｺ"), y)

    ' ---- BRS ----
    y = y + ROM_HDR_GAP
    y = AddSectionTitle(host, "Brunnstrom Recovery Stage・・RS・・, y)
    Dim brsValues As Variant
    brsValues = Array("竇", "竇｡", "竇｢", "竇｣", "竇､", "竇･")
    y = AddComboRow(host, "荳願い", "cboBRS_Upper", brsValues, y)
    y = AddComboRow(host, "謇区欠", "cboBRS_Hand", brsValues, y)
    y = AddComboRow(host, "荳玖い", "cboBRS_Lower", brsValues, y)

    ' ---- 髫丈ｼｴ迴ｾ雎｡ ----
    y = y + ROM_HDR_GAP
    y = AddSectionTitle(host, "髫丈ｼｴ迴ｾ雎｡", y)
    y = AddCheckRow(host, "蜈ｱ蜷碁°蜍・, "chkSynergy", y)
    y = AddCheckRow(host, "騾｣蜷亥渚蠢・, "chkAssociatedRxn", y)

    ' ---- 蛯呵・----
    PlaceMemoBelow host, w, h, y, "txtParalysisMemo"
End Sub

'---- 蟆冗黄・夊｡後ン繝ｫ繝・郁ｦ句・縺暦ｼ上さ繝ｳ繝懆｡鯉ｼ上メ繧ｧ繝・け陦鯉ｼ・---
Private Function AddSectionTitle(host As MSForms.Frame, ttl As String, y As Single) As Single
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
    AddSectionTitle = PX(y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddComboRow(host As MSForms.Frame, cap As String, nameCombo As String, _
                             items As Variant, y As Single) As Single
    Dim wCaption As Single, wCombo As Single, xCaption As Single, xCombo As Single
    wCaption = PX(COL1_W)
    wCombo = PX(ROM_COL_EDT_W)
    xCaption = PX(PAD_X)
    xCombo = PX(PAD_X + wCaption + 8)

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCaption
        .Top = PX(y)
        .Width = wCaption
        .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim cbo As MSForms.ComboBox
    Set cbo = host.controls.Add("Forms.ComboBox.1", nameCombo, True)
    With cbo
        .Left = xCombo
        .Top = PX(y)
        .Width = wCombo
        .Height = ROW_H
        .Style = fmStyleDropDownList
        .List = items
    End With

    AddComboRow = PX(y + ROW_H + ROM_GAP_Y)
End Function

Private Function AddCheckRow(host As MSForms.Frame, cap As String, nameChk As String, y As Single) As Single
    Dim wCaption As Single, xCaption As Single, xChk As Single
    wCaption = PX(COL1_W)
    xCaption = PX(PAD_X)
    xChk = PX(PAD_X + wCaption + 8)

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = cap
        .Left = xCaption
        .Top = PX(y)
        .Width = wCaption
        .Height = ROW_H
        .TextAlign = fmTextAlignLeft
    End With

    Dim chk As MSForms.CheckBox
    Set chk = host.controls.Add("Forms.CheckBox.1", nameChk, True)
    With chk
        .caption = "譛・
        .Left = xChk
        .Top = PX(y)
        .Width = PX(60)
        .Height = ROW_H
    End With

    AddCheckRow = PX(y + ROW_H + ROM_GAP_Y)
End Function


'============================================================
' ROM 繝壹・繧ｸ・壼ｭ舌ち繝厄ｼ井ｸ願い・丈ｸ玖い・峨ン繝ｫ繝繝ｼ
'============================================================
Public Sub BuildROMTabs(host As MSForms.Frame)
    Debug.Print "[ROM] BuildROMTabs called"

    Dim mp As MSForms.MultiPage
    Set mp = host.controls.Add("Forms.MultiPage.1", "mpROM")
    With mp
        .Left = PX(PAD_X)
        .Top = PX(PAD_Y)
        .Width = PX(host.InsideWidth - PAD_X * 2)
        .Height = PX(host.InsideHeight - PAD_Y * 2)
        .Style = fmTabStyleTabs
    End With

    Dim pUpper As MSForms.page, pLower As MSForms.page, pTrunk As MSForms.page
    Set pUpper = mp.Pages.Add: pUpper.caption = "荳願い": pUpper.name = "pgROM_Upper"
    Set pLower = mp.Pages.Add: pLower.caption = "荳玖い": pLower.name = "pgROM_Lower"
    Set pTrunk = mp.Pages.Add: pTrunk.caption = "菴灘ｹｹ": pTrunk.name = "pgROM_Trunk"
    
    Dim hostUpper As MSForms.Frame, hostLower As MSForms.Frame, hostTrunk As MSForms.Frame
    Set hostUpper = EnsureHostFrame(pUpper)
    Set hostLower = EnsureHostFrame(pLower)
    Set hostTrunk = EnsureHostFrame(pTrunk)
    
    BuildROM_Upper hostUpper
    NormalizeRomColumns hostUpper

    BuildROM_Lower hostLower
    NormalizeRomColumns hostLower
    
    BuildROM_Trunk hostTrunk
    NormalizeRomColumns hostTrunk

  
End Sub

Public Sub BuildROM_Trunk(host As MSForms.Frame)

    Dim w As Single, h As Single
    w = host.Width
    h = host.Height

    Dim y As Single
    y = ROM_GROUP_PAD

    ' 鬆ｸ驛ｨ
    y = BuildRomTrunkJointTable(host, "Trunk", "Neck", "鬆ｸ驛ｨ", y)
    y = y + ROM_JOINT_GAP_Y

    ' 菴灘ｹｹ
    y = BuildRomTrunkJointTable(host, "Trunk", "Trunk", "菴灘ｹｹ", y)
    y = y + ROM_JOINT_GAP_Y

    ' 閭ｸ驛ｭ蜿ｯ蜍・
    y = BuildThoraxMobilityBlock(host, y)

    ' 繝｡繝｢谺・
    PlaceMemoBelow host, w, h, y, "txtROM_Trunk_Memo"

End Sub

Private Function BuildRomTrunkJointTable(host As MSForms.Frame, _
            region As String, jointKey As String, jointTitle As String, _
            y0 As Single) As Single

    Const TRUNK_COL_GAP As Single = 8
    Const TRUNK_LABEL_W As Single = 72
    Const TRUNK_START_GAP As Single = 10

    Dim motions As Variant
    motions = Split("Flex,Ext,Rot,LatFlex", ",")

    Dim fr As MSForms.Frame
    Set fr = host.controls.Add("Forms.Frame.1")

    With fr
        .caption = NormalizeRomFrameTitle(jointTitle)
        .Left = PX(PAD_X)
        .Top = PX(y0)
        .Width = PX(host.Width - PAD_X * 2 - ROM_FRAME_TRIM_R)
        .Height = PX(ROM_GROUP_PAD * 2 + ROM_ROW_H + _
                     (UBound(motions) - LBound(motions) + 1) * ROM_ROW_H + _
                     (UBound(motions) - LBound(motions)) * ROM_MOTION_GAP_Y)
    End With

    Dim xName As Single, xSingle As Single, xR As Single, xL As Single
    xName = PX(ROM_GROUP_PAD)

    ' 蛻励ｒ蜿ｳ遶ｯ蝓ｺ貅悶〒縺ｯ縺ｪ縺上√Λ繝吶Ν蝓ｺ貅悶〒蝗ｺ螳・
    xSingle = PX(xName + TRUNK_LABEL_W + TRUNK_START_GAP)
    xR = PX(xSingle + ROM_COL_EDT_W + TRUNK_COL_GAP)
    xL = PX(xR + ROM_COL_EDT_W + TRUNK_COL_GAP)

    Dim hdrSingle As MSForms.label, hdrR As MSForms.label, hdrL As MSForms.label

    Set hdrSingle = fr.controls.Add("Forms.Label.1")
    With hdrSingle
        .caption = "蜊倡峡"
        .Left = xSingle
        .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Set hdrR = fr.controls.Add("Forms.Label.1")
    With hdrR
        .caption = "蜿ｳ"
        .Left = xR
        .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Set hdrL = fr.controls.Add("Forms.Label.1")
    With hdrL
        .caption = "蟾ｦ"
        .Left = xL
        .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Dim i As Long, rowY As Single, motionKey As String
    rowY = PX(ROM_GROUP_PAD + ROM_ROW_H)

    For i = LBound(motions) To UBound(motions)
        motionKey = CStr(motions(i))
        rowY = BuildRomTrunkMotionRow(fr, region, jointKey, motionKey, rowY, xName, xSingle, xR, xL)
        If i < UBound(motions) Then rowY = rowY + ROM_MOTION_GAP_Y
    Next i

    BuildRomTrunkJointTable = fr.Top + fr.Height + ROM_GAP_Y
End Function


Private Function BuildRomTrunkMotionRow(host As MSForms.Frame, _
            region As String, jointKey As String, motionKey As String, _
            y0 As Single, xName As Single, xSingle As Single, xR As Single, xL As Single) As Single

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")

    With lbl
        .caption = GetTrunkMotionCaption(motionKey)
        .Left = xName
        .Top = PX(y0)
        .Width = PX(host.Width - xName - (host.Width - xSingle) + ROM_HDR_RL_GAP)
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignLeft
        .Font.Bold = True
    End With

    If motionKey = "Flex" Or motionKey = "Ext" Then

        Dim txtSingle As MSForms.TextBox
        Set txtSingle = host.controls.Add("Forms.TextBox.1", _
            "txtROM_" & region & "_" & jointKey & "_" & motionKey)

        With txtSingle
            .Left = xSingle
            .Top = PX(y0 + (ROM_ROW_H - ROM_TXT_H) / 2)
            .Width = ROM_COL_EDT_W
            .Height = ROM_TXT_H
            .IMEMode = fmIMEModeDisable
        End With

    Else

        Dim tR As MSForms.TextBox, tL As MSForms.TextBox

        Set tR = host.controls.Add("Forms.TextBox.1", _
            "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_R")

        With tR
            .Left = xR
            .Top = PX(y0 + (ROM_ROW_H - ROM_TXT_H) / 2)
            .Width = ROM_COL_EDT_W
            .Height = ROM_TXT_H
            .IMEMode = fmIMEModeDisable
            .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|R"
        End With

        Set tL = host.controls.Add("Forms.TextBox.1", _
            "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_L")

        With tL
            .Left = xL
            .Top = PX(y0 + (ROM_ROW_H - ROM_TXT_H) / 2)
            .Width = ROM_COL_EDT_W
            .Height = ROM_TXT_H
            .IMEMode = fmIMEModeDisable
            .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|L"
        End With

    End If

    BuildRomTrunkMotionRow = y0 + ROM_ROW_H
End Function


Private Function GetTrunkMotionCaption(motionKey As String) As String

    Select Case motionKey

        Case "Flex"
            GetTrunkMotionCaption = "螻域峇"

        Case "Ext"
            GetTrunkMotionCaption = "莨ｸ螻・

        Case "Rot"
            GetTrunkMotionCaption = "蝗樊雷"

        Case "LatFlex"
            GetTrunkMotionCaption = "蛛ｴ螻・

        Case Else
            GetTrunkMotionCaption = motionKey

    End Select

End Function

'------------------------------------------------------------
' 荳願い 蟄舌ち繝・
'------------------------------------------------------------
Public Sub BuildROM_Upper(host As MSForms.Frame)
    Dim w As Single, h As Single
    w = host.Width: h = host.Height

    Dim y As Single: y = ROM_GROUP_PAD
    
    

    ' 閧ｩ
    y = BuildRomJointBlock(host, "Upper", "Shoulder", "縲・閧ｩ髢｢遽 縲・, _
            Split("Flex,Ext,Abd,Add,ER,IR", ","), y)
    y = y + ROM_JOINT_GAP_Y

    ' 閧・
    y = BuildRomJointBlock(host, "Upper", "Elbow", "縲・閧倬未遽 縲・, _
            Split("Flex,Ext", ","), y)
    y = y + ROM_JOINT_GAP_Y

    ' 蜑崎・
    y = BuildRomJointBlock(host, "Upper", "Forearm", "縲・蜑崎・ 縲・, _
            Split("Sup,Pro", ","), y)
    y = y + ROM_JOINT_GAP_Y

    ' 謇矩未遽
    y = BuildRomJointBlock(host, "Upper", "Wrist", "縲・謇矩未遽 縲・, _
            Split("Dorsi,Palmar,Radial,Ulnar", ","), y)
            
            
 PlaceMemoBelow host, w, h, y, "txtROM_Upper_Memo"
 

End Sub

'------------------------------------------------------------
' 荳玖い 蟄舌ち繝・
'------------------------------------------------------------
Public Sub BuildROM_Lower(host As MSForms.Frame)
    Dim w As Single, h As Single
    w = host.Width: h = host.Height

    Dim y As Single: y = ROM_GROUP_PAD
    
    
    ' 閧｡
    y = BuildRomJointBlock(host, "Lower", "Hip", "縲・閧｡髢｢遽 縲・, _
            Split("Flex,Ext,Abd,Add,ER,IR", ","), y)
    y = y + ROM_JOINT_GAP_Y

    ' 閹・
    y = BuildRomJointBlock(host, "Lower", "Knee", "縲・閹晞未遽 縲・, _
            Split("Flex,Ext", ","), y)
    y = y + ROM_JOINT_GAP_Y

    ' 雜ｳ髢｢遽
    y = BuildRomJointBlock(host, "Lower", "Ankle", "縲・雜ｳ髢｢遽 縲・, _
            Split("Dorsi,Plantar,Inv,Ev", ","), y)
            
            
PlaceMemoBelow host, w, h, y, "txtROM_Lower_Memo"
End Sub

'------------------------------------------------------------
' 髢｢遽繝悶Ο繝・け・域棧・矩°蜍戊｡鯉ｼ・
'   motions: "Flex","Ext"...・郁恭逡･繧ｭ繝ｼ・・
'   謌ｻ繧・ 谺｡繝悶Ο繝・け縺ｮ髢句ｧ亀op
'------------------------------------------------------------
Private Function BuildRomJointBlock(host As MSForms.Frame, _
            region As String, jointKey As String, jointTitle As String, _
            motions As Variant, y0 As Single) As Single

    Dim fr As MSForms.Frame
    Set fr = host.controls.Add("Forms.Frame.1")
    With fr
        .caption = jointTitle
        .Left = PX(PAD_X)
        .Top = PX(y0)
        .Width = PX(host.Width - PAD_X * 2)
    End With

Dim motionCount As Long
motionCount = UBound(motions) - LBound(motions) + 1

Dim useTwoCols As Boolean
useTwoCols = IsRomTwoColumnJoint(region, jointKey)

Dim rowCount As Long
If useTwoCols Then
    rowCount = (motionCount + 1) \ 2
Else
    rowCount = motionCount
End If

Dim topY As Single
topY = PX(ROM_GROUP_PAD + ROM_ROW_H + ROM_HDR_GAP)

If useTwoCols Then
    fr.Height = PX(topY + rowCount * ROM_ROW_H + _
                   (rowCount - 1) * ROM_MOTION_GAP_Y + ROM_GROUP_PAD)
Else
    fr.Height = PX(topY + rowCount * ROM_ROW_H + _
                   (rowCount - 1) * ROM_MOTION_GAP_Y + ROM_GROUP_PAD)
End If

If useTwoCols Then

 
    Dim groupW As Single
    groupW = (fr.Width - ROM_GROUP_PAD * 2 - ROM_MULTI_COL_GAP) / 2

    Dim xGroupL As Single, xGroupR As Single
    xGroupL = PX(ROM_GROUP_PAD)
    xGroupR = PX(xGroupL + groupW + ROM_MULTI_COL_GAP)

    Dim labelW As Single
    labelW = PX(groupW - (ROM_COL_EDT_W * 2 + ROM_MULTI_EDIT_GAP))
    If labelW < 36 Then labelW = 36
    
    Dim xNameL As Single, xRL As Single, xLL As Single

    xNameL = xGroupL
    xRL = PX(xNameL + ROM_LABEL_W)
    xLL = PX(xRL + ROM_COL_EDT_W + ROM_MULTI_EDIT_GAP)
    Dim xNameR As Single, xRR As Single, xLR As Single

    xNameR = xGroupR
    xRR = PX(xNameR + ROM_LABEL_W)
    xLR = PX(xRR + ROM_COL_EDT_W + ROM_MULTI_EDIT_GAP)

    AddRomRLHeader fr, xRL, xLL
    AddRomRLHeader fr, xRR, xLR

    Dim splitIndex As Long
    splitIndex = rowCount

    Dim i As Long, rowIndex As Long, rowY As Single

    For i = LBound(motions) To UBound(motions)

        If i - LBound(motions) < splitIndex Then
            rowIndex = i - LBound(motions)
            rowY = topY + rowIndex * (ROM_ROW_H + ROM_MOTION_GAP_Y)

         
            BuildRomMotionRowAt fr, region, jointKey, CStr(motions(i)), _
                                rowY, xNameL, xRL, xLL

        Else
            rowIndex = i - LBound(motions) - splitIndex
            rowY = topY + rowIndex * (ROM_ROW_H + ROM_MOTION_GAP_Y)

  
            BuildRomMotionRowAt fr, region, jointKey, CStr(motions(i)), _
                                rowY, xNameR, xRR, xLR
        End If

    Next i

Else

    Dim xName As Single, xR As Single, xL As Single

    xName = PX(ROM_GROUP_PAD)
    xR = PX(xName + ROM_LABEL_W)
    xL = PX(xR + ROM_COL_EDT_W + ROM_MULTI_EDIT_GAP)
    
    AddRomRLHeader fr, xR, xL

    Dim rowYSingle As Single
    Dim j As Long

    rowYSingle = topY

    For j = LBound(motions) To UBound(motions)

        BuildRomMotionRowAt fr, region, jointKey, CStr(motions(j)), _
                            rowYSingle, xName, xR, xL

        rowYSingle = rowYSingle + ROM_ROW_H + ROM_MOTION_GAP_Y

    Next j

End If

AttachTxtImeHookInFrame fr

BuildRomJointBlock = fr.Top + fr.Height + ROM_GAP_Y

End Function


Private Function IsRomTwoColumnJoint(ByVal region As String, ByVal jointKey As String) As Boolean

    If region = "Upper" Then
        ' 荳願い
        IsRomTwoColumnJoint = (jointKey = "Shoulder" Or jointKey = "Wrist")

    ElseIf region = "Lower" Then
        ' 荳玖い
        IsRomTwoColumnJoint = (jointKey = "Hip" Or jointKey = "Ankle")

    Else
        IsRomTwoColumnJoint = False
    End If

End Function


Private Sub AddRomRLHeader(host As MSForms.Frame, _
                           ByVal xR As Single, _
                           ByVal xL As Single)

    Dim lblR As MSForms.label
    Dim lblL As MSForms.label
    

    Set lblR = host.controls.Add("Forms.Label.1")
    With lblR
        .caption = "E"
        .Left = xR
        .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    Set lblL = host.controls.Add("Forms.Label.1")
    With lblL
        .caption = ""
        .Left = xL
        .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With

    
   
    Set lblR = host.controls.Add("Forms.Label.1")
    With lblR
        .caption = "蜿ｳ": .Left = xR: .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter: .Font.Bold = True
    End With
    
    Set lblL = host.controls.Add("Forms.Label.1")
    With lblL
        .caption = "蟾ｦ": .Left = xL: .Top = PX(ROM_GROUP_PAD)
        .Width = ROM_COL_EDT_W: .Height = ROM_ROW_H
        .TextAlign = fmTextAlignCenter: .Font.Bold = True
    End With
    
End Sub

'------------------------------------------------------------
' 1驕句虚蛻・・陦鯉ｼ医Λ繝吶Ν・騎/L繝・く繧ｹ繝茨ｼ・
'------------------------------------------------------------
Private Sub BuildRomMotionRowAt(host As MSForms.Frame, _
        region As String, jointKey As String, motionKey As String, _
        y0 As Single, xName As Single, xR As Single, xL As Single)

    Dim lbl As MSForms.label
    Set lbl = host.controls.Add("Forms.Label.1")
    With lbl
        .caption = GetMotionCaption(jointKey, motionKey)
        .Left = xName: .Top = PX(y0)
        .Width = PX(host.Width - xName - (host.Width - xR) + ROM_HDR_RL_GAP)
        .Height = ROM_ROW_H
        .TextAlign = fmTextAlignLeft
        .Font.Bold = True
    End With

   Dim tR As MSForms.TextBox, tL As MSForms.TextBox
    ' EiRj
Set tR = host.controls.Add("Forms.TextBox.1", _
    "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_R")
With tR
    .Left = xR: .Top = PX(y0 + (ROM_ROW_H - ROM_TXT_H) / 2)
    .Width = ROM_COL_EDT_W: .Height = ROM_TXT_H
    .IMEMode = fmIMEModeDisable
    .TextAlign = fmTextAlignRight
    .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|E"   ' ?
End With


Set tL = host.controls.Add("Forms.TextBox.1", _
    "txtROM_" & region & "_" & jointKey & "_" & motionKey & "_L")
With tL
    .Left = xL: .Top = PX(y0 + (ROM_ROW_H - ROM_TXT_H) / 2)
    .Width = ROM_COL_EDT_W: .Height = ROM_TXT_H
    .IMEMode = fmIMEModeDisable
    .TextAlign = fmTextAlignRight
    .tag = TAG_FUNC_PREFIX & "|ROM|" & jointKey & "|" & motionKey & "|"    ' ?i?Lj
End With

End Sub

Private Function NormalizeRomFrameTitle(ByVal title As String) As String
    Dim normalized As String
    normalized = Replace(title, " ", "")
    normalized = Replace(normalized, ChrW$(12288), "")
    NormalizeRomFrameTitle = normalized
End Function


Private Function BuildRomSingleJointBlock(host As MSForms.Frame, _
            region As String, jointKey As String, jointTitle As String, _
            motionKey As String, motionCaption As String, y0 As Single) As Single

    Dim fr As MSForms.Frame
  

    Set fr = host.controls.Add("Forms.Frame.1")

    With fr

        .caption = jointTitle
        .Left = PX(PAD_X)
        .Top = PX(y0)
        .Width = PX(host.Width - PAD_X * 2)
        .Height = PX(ROM_ROW_H + ROM_GROUP_PAD * 2)
    End With

    Dim lbl As MSForms.label
    Set lbl = fr.controls.Add("Forms.Label.1")

    With lbl
        .caption = motionCaption
        .Left = PX(ROM_GROUP_PAD)
        .Top = PX(ROM_GROUP_PAD)
        .Width = PX(fr.Width - ROM_GROUP_PAD * 3 - ROM_COL_EDT_W)
        .Height = ROM_ROW_H
        .Font.Bold = True
    End With

    Dim txt As MSForms.TextBox
    Set txt = fr.controls.Add("Forms.TextBox.1", "txtROM_" & region & "_" & jointKey & "_" & motionKey)

    With txt
        .Left = PX(fr.Width - ROM_GROUP_PAD - ROM_COL_EDT_W)
        .Top = PX(ROM_GROUP_PAD + (ROM_ROW_H - ROM_TXT_H) / 2)
        .Width = ROM_COL_EDT_W
        .Height = ROM_TXT_H
        .IMEMode = fmIMEModeDisable
    End With

    BuildRomSingleJointBlock = fr.Top + fr.Height + ROM_MOTION_GAP_Y

End Function

Private Function BuildThoraxMobilityBlock(host As MSForms.Frame, y0 As Single) As Single

    Dim fr As MSForms.Frame
    Set fr = host.controls.Add("Forms.Frame.1")

    With fr
        .caption = "閭ｸ驛ｭ蜿ｯ蜍・
        .Left = PX(PAD_X)
        .Top = PX(y0)
        .Width = PX(host.Width - PAD_X * 2)
        .Height = PX(ROM_ROW_H + ROM_GROUP_PAD * 2)
    End With

    Dim lbl As MSForms.label
    Set lbl = fr.controls.Add("Forms.Label.1")

    With lbl
        .caption = "閭ｸ蝗ｲ蟾ｮ・亥精豌暦ｼ榊他豌暦ｼ・
        .Left = PX(ROM_GROUP_PAD)
        .Top = PX(ROM_GROUP_PAD)
        .Width = PX(110)
        .Height = ROM_ROW_H
        .Font.Bold = True
    End With

    Dim txt As MSForms.TextBox
    Set txt = fr.controls.Add("Forms.TextBox.1", "txtROM_Trunk_Thorax_ChestDiff")

    With txt
        .Left = PX(lbl.Left + lbl.Width + 8)
        .Top = PX(ROM_GROUP_PAD + (ROM_ROW_H - ROM_TXT_H) / 2)
        .Width = ROM_COL_EDT_W
        .Height = ROM_TXT_H
        .IMEMode = fmIMEModeDisable
    End With

    Dim lblCm As MSForms.label
    Set lblCm = fr.controls.Add("Forms.Label.1")

    With lblCm
        .caption = "cm"
        .Left = PX(txt.Left + txt.Width + 4)
        .Top = PX(ROM_GROUP_PAD)
        .Width = 24
        .Height = ROM_ROW_H
    End With

    BuildThoraxMobilityBlock = fr.Top + fr.Height + ROM_MOTION_GAP_Y

End Function

'------------------------------------------------------------
' 驕句虚繝ｩ繝吶Ν・亥柱蜷搾ｼ・
'------------------------------------------------------------
Private Function GetMotionCaption(jointKey As String, motionKey As String) As String
    Select Case jointKey
        Case "Shoulder", "Hip"
            Select Case motionKey
                Case "Flex":   GetMotionCaption = "螻域峇"
                Case "Ext":    GetMotionCaption = "莨ｸ螻・
                Case "Abd":    GetMotionCaption = "螟冶ｻ｢"
                Case "Add":    GetMotionCaption = "蜀・ｻ｢"
                Case "ER":     GetMotionCaption = "螟匁雷"
                Case "IR":     GetMotionCaption = "蜀・雷"
            End Select
        Case "Elbow", "Knee"
            Select Case motionKey
                Case "Flex":   GetMotionCaption = "螻域峇"
                Case "Ext":    GetMotionCaption = "莨ｸ螻・
            End Select
        Case "Forearm"
            Select Case motionKey
                Case "Sup":    GetMotionCaption = "蝗槫､・
                Case "Pro":    GetMotionCaption = "蝗槫・"
            End Select
        Case "Wrist"
            Select Case motionKey
                Case "Dorsi":  GetMotionCaption = "閭悟ｱ・
                Case "Palmar": GetMotionCaption = "謗悟ｱ・
                Case "Radial": GetMotionCaption = "讖亥ｱ・
                Case "Ulnar":  GetMotionCaption = "蟆ｺ螻・
            End Select
        Case "Ankle"
            Select Case motionKey
                Case "Dorsi":   GetMotionCaption = "閭悟ｱ・
                Case "Plantar": GetMotionCaption = "蠎募ｱ・
                Case "Inv":     GetMotionCaption = "蜀・′縺医＠"
                Case "Ev":      GetMotionCaption = "螟悶′縺医＠"
            End Select
            
        Case "Neck", "Trunk"
            Select Case motionKey
                Case "Rot":      GetMotionCaption = ""
                Case "LatFlex":  GetMotionCaption = ""
            End Select
    End Select
End Function

'------------------------------------------------------------
' 繝輔Ξ繝ｼ繝蜀・・ TextBox 縺ｫ IME Off 繝輔ャ繧ｯ繧剃ｸ諡ｬ繧｢繧ｿ繝・メ
'   窶ｻ TxtImeHook.cls 縺ｮ蜈ｬ髢九Γ繧ｽ繝・ラ蜷・Attach 繧呈Φ螳・
'------------------------------------------------------------
Private Sub AttachTxtImeHookInFrame(fr As MSForms.Frame)
    On Error Resume Next
    Dim ctl As MSForms.Control
    For Each ctl In fr.controls
        If TypeOf ctl Is MSForms.TextBox Then
            Dim Hook As TxtImeHook
            Set Hook = New TxtImeHook
            Hook.Init ctl   ' 笘・％縺薙ｒ Attach 竊・Init 縺ｫ
        End If
    Next
    On Error GoTo 0
End Sub


'=== LoadLatestROMNow・・025-10-22邨ｱ蜷育沿・・==
Public Sub LoadLatestROMNow(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("ROM_Upper_Shoulder_Flex_R", ws)
    If r <= 0 Then
        Debug.Print "[LoadROM] header not found"
        Exit Sub
    End If

    '--- 譌ｧ隱ｭ霎ｼ縺ｯ蠕梧婿莠呈鋤縺ｮ縺溘ａ繧ｳ繝｡繝ｳ繝医い繧ｦ繝・---
    'Call ParseROMData(ws.Cells(r, HeaderCol("IO_ROM", ws)).Value)
    '-----------------------------------------------------

    Dim raw As String
    raw = LoadLatestROMNow_Raw(ws)
    Debug.Print "[LoadROM] R=" & r & " Len=" & Len(raw) & " | " & Left$(raw, 60)
End Sub



