VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEval 
   Caption         =   "隧穂ｾ｡繝輔か繝ｼ繝"
   ClientHeight    =   8580.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17124
   OleObjectBlob   =   "frmEval.frx":0000
   StartUpPosition =   1  '繧ｪ繝ｼ繝翫・ 繝輔か繝ｼ繝縺ｮ荳ｭ螟ｮ
End
Attribute VB_Name = "frmEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'=== frmEval 繝倥ャ繝・壼・騾壹〒菴ｿ縺・､画焚 ===
Private mp As MSForms.MultiPage            ' 繝ｫ繝ｼ繝・MultiPage
Private mpWalk As MSForms.MultiPage        ' 豁ｩ陦瑚ｩ穂ｾ｡繧ｿ繝門・縺ｮ繧ｵ繝・MultiPage  竊・縺薙ｌ縺御ｻ雁屓縺ｮ霑ｽ蜉
Private hostWalkGait As MSForms.Frame
Private fGait As MSForms.Frame
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
Private Const CAP_POSTURE_PAGE As String = "霄ｫ菴捺ｩ溯・・玖ｵｷ螻・虚菴・
Private Const POSTURE_TAG_PREFIX As String = "POSTURE|"
Private Const POSTURE_COLS As Long = 2
Private mStyleDone As Boolean
Private mMPHooks  As Collection
Private mTxtHooks As Collection
Private mHooked As Boolean
Private mLayoutDone As Boolean
Private mPainBuilt As Boolean
Private mPainLayoutDone As Boolean
' 譌｢蟄倥・縲後す繝ｼ繝医∈菫晏ｭ倥阪悟燕蝗槭・蛟､繧定ｪｭ縺ｿ霎ｼ繧縲阪↓繝輔ャ繧ｯ縺吶ｋ
Private WithEvents btnHdrSave     As MSForms.CommandButton
Attribute btnHdrSave.VB_VarHelpID = -1
Private WithEvents btnHdrLoadPrev As MSForms.CommandButton
Attribute btnHdrLoadPrev.VB_VarHelpID = -1
Private fBasicRef As MSForms.Frame
Private nextTop As Single
' BI繧ｳ繝ｳ繝懊・繧､繝吶Φ繝医ヵ繝・け繧剃ｿ晄戟
Private BIHooks As Collection
' 繝ｬ繧､繧｢繧ｦ繝育畑
Private FWIDTH As Single, COL_LX As Single, COL_RX As Single, lblW As Single, pad As Single, rowH As Single
' === 繧ｦ繧｣繝ｳ繝峨え蛻ｶ邏・ｼ医せ繧ｯ繝ｭ繝ｼ繝ｫ辟｡縺礼畑縺ｮ譛蟆上し繧､繧ｺ・・==
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
Private mMPPhysHookAttached As Boolean
Private mMMTFirstBuildDone As Boolean
Private mPendingMMTLoad As Boolean
Private mPendingMMTRow As Long
Private mApplyingPendingMMTLoad As Boolean
Private mPendingMMTWs As Worksheet
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
    qmAsk   '髢峨§繧九・繧ｿ繝ｳ・壻ｿ晏ｭ倡｢ｺ隱阪＠縺ｦExcel邨ゆｺ・
End Enum

Private mQuitMode As QuitMode
Private mHdrArchiveHook As clsHdrBtnHook
Private mHdrLoadPrevHook As clsHdrBtnHook
Private mPrintBtnHook As clsPrintBtnHook
Private mRomMirrorHooks As Collection
Private mHdrNameSink As cHdrNameSink
Private mNameSuggestSink As cNameSuggestSink
Private mDupNameWarned As Boolean
Private mBasicInfoTidyDone As Boolean
Private mAgeBusy As Boolean
Private WithEvents mBIEnter_txtLiving As MSForms.TextBox
Attribute mBIEnter_txtLiving.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtEvaluator As MSForms.TextBox
Attribute mBIEnter_txtEvaluator.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtEvaluatorJob As MSForms.TextBox
Attribute mBIEnter_txtEvaluatorJob.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtOnset As MSForms.TextBox
Attribute mBIEnter_txtOnset.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtDx As MSForms.TextBox
Attribute mBIEnter_txtDx.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtAdmDate As MSForms.TextBox
Attribute mBIEnter_txtAdmDate.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtDisDate As MSForms.TextBox
Attribute mBIEnter_txtDisDate.VB_VarHelpID = -1
Private WithEvents mBIEnter_txtTxCourse As MSForms.TextBox
Attribute mBIEnter_txtTxCourse.VB_VarHelpID = -1



Public Sub SyncAgeFromBirth()
    On Error GoTo EH

    UpdateAgeFromBirth

    Exit Sub
EH:
#If APP_DEBUG Then
    Debug.Print "[SyncAgeFromBirth][ERR]", Err.Number, Err.Description
#End If
End Sub

Public Function TryGetBirthDateForStorage(ByVal raw As String, ByRef outDate As Date) As Boolean
    TryGetBirthDateForStorage = TryParseBirthDate_ShowaOrAD(raw, outDate)
End Function


'=== 蜈ｱ騾壹・繝ｫ繝代・・壹ヵ繝ｬ繝ｼ繝縺ｮ鬮倥＆繧貞ｭ舌さ繝ｳ繝医Ο繝ｼ繝ｫ縺ｮ荳逡ｪ荳具ｼ倶ｽ咏區縺ｾ縺ｧ莨ｸ縺ｰ縺・===
Private Sub FitFrameHeightToChildren(f As MSForms.Frame, Optional margin As Single = 6)
    
    'FitFrameHeightToChildren Me.Controls("Frame7")
    
    Dim c As Control
    Dim maxBottom As Single

    If f Is Nothing Then Exit Sub

    For Each c In f.controls
        If c.Top + c.Height > maxBottom Then
            maxBottom = c.Top + c.Height
        End If
    Next c

    If maxBottom + margin > f.Height Then
        f.Height = maxBottom + margin
    End If
End Sub







'--- 蛻苓ｧ｣豎ｺ縺ｮ螳牙・繝ｩ繝・ヱ・医≠繧後・ ResolveColumnLocal縲√↑縺代ｌ縺ｰ隕句・縺怜錐縺ｧ謗｢縺呻ｼ・--
Private Function RCol(ws As Worksheet, look As Object, ParamArray headers()) As Long
    Dim c As Long, i As Long
    On Error Resume Next
    c = ResolveColumnLocal(look, CStr(headers(LBound(headers))))  ' 蟄伜惠縺励↑縺・腸蠅・〒繧０K
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



' 譌｢蟄倥さ繝ｼ繝峨〒蜿ら・縺輔ｌ繧九′譛ｪ螳夂ｾｩ縺縺｣縺溘ｂ縺ｮ繧呈怙蟆城剞縺縺醍畑諢・
Private Sub SetupLayout()
    ' 窶ｻ謨ｰ蛟､縺ｯ荳闊ｬ逧・↑譌｢螳壹し繧､繧ｺ縲６I隕九◆逶ｮ縺ｯ螟峨∴縺ｪ縺・ｯ・峇縺ｧ螳牙・縺ｪ蛟､縲・
    FWIDTH = 800
    lblW = 70
    COL_LX = 12
    COL_RX = 380
    pad = 6
    rowH = 24
End Sub

' 譌｢蟄倥さ繝ｼ繝峨〒 Call 縺輔ｌ繧九′荳ｭ霄ｫ縺御ｸ崎ｦ√↑莠呈鋤繝繝溘・・・o-op・・
Private Sub BuildBliadlControls(ByVal mp As MSForms.MultiPage)
    ' no-op・井ｺ呈鋤逕ｨ・・
End Sub
Private Sub BuildBIPage(ByVal mpADL As MSForms.MultiPage)
    ' no-op・井ｺ呈鋤逕ｨ・・
End Sub
Private Sub BuildIADLPage(ByVal mpADL As MSForms.MultiPage)
    ' no-op・井ｺ呈鋤逕ｨ・・
End Sub

' 譌｢蟄倥・菫晏ｭ伜・逅・〒蜿ら・縺輔ｌ繧九Λ繝・ヱ繝ｼ・郁ｦ九◆逶ｮ縺ｨ謖吝虚縺ｯ螟峨∴縺ｪ縺・ｼ・
Private Function CtrlText(ByVal ctrlName As String) As String
    On Error Resume Next
    CtrlText = Trim$(Me.controls(ctrlName).text & "")
End Function
Private Sub SetCtrlText(ByVal ctrlName As String, ByVal v As String)
    On Error Resume Next
    Me.controls(ctrlName).text = v
End Sub

Private Function SafeGetControl(ByVal parent As Object, ByVal nm As String) As Object
    Set SafeGetControl = modCommonUtil.SafeGetControl(parent, nm)
End Function

Public Function EvalCtl(ByVal ctrlName As String, Optional ByVal pageKey As Variant) As Object
    Dim root As Object
    Dim mpRoot As Object

    If IsMissing(pageKey) Then
        Set root = Me
    Else
        Set mpRoot = SafeGetControl(Me, "MultiPage1")
        If mpRoot Is Nothing Then Exit Function
        Set root = SafeGetPage(mpRoot, pageKey)
    End If

    If root Is Nothing Then Exit Function
    Set EvalCtl = SafeGetControl(root, ctrlName)
End Function

Private Function GetPainHost() As Object
    Dim c As Object

    Set c = EvalCtl("fraPainCourse")
    If Not c Is Nothing Then Set GetPainHost = c.parent: Exit Function

    Set c = EvalCtl("fraPainSite")
    If Not c Is Nothing Then Set GetPainHost = c.parent: Exit Function

    Set c = EvalCtl("fraVAS")
    If Not c Is Nothing Then Set GetPainHost = c.parent: Exit Function

    Set c = EvalCtl("txtPainMemo")
    If Not c Is Nothing Then Set GetPainHost = c.parent: Exit Function

    Set c = EvalCtl("cmbNRS_Move")
    If Not c Is Nothing Then Set GetPainHost = c.parent
End Function

Private Function DailyLogCtl(ByVal ctrlName As String) As Object
    Dim f As Object
    Set f = GetDailyLogFrame()
    If f Is Nothing Then Exit Function
    Set DailyLogCtl = SafeGetControl(f, ctrlName)
End Function


'=== 縺薙％縺九ｉ 逕ｻ髱｢菴懈・繝倥Ν繝代・縺ｮ譛蟆丞ｮ溯｣・=========================
Private Function CreateFrameP(parent As MSForms.Frame, title As String, _
                              Optional minHeight As Single = 120) As MSForms.Frame
    Dim f As MSForms.Frame
    Set f = parent.controls.Add("Forms.Frame.1")
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

Private Sub nL(ByRef y As Single, Optional ByVal rows As Long = 1)
    y = y + rows * 26
End Sub



' 譌｢蟄伜他縺ｳ蜃ｺ縺嶺ｺ呈鋤・喞aption 竊・x 竊・y 竊・[w] 竊・[name]
Function CreateLabel( _
    parent As MSForms.Frame, _
    ByVal caption As String, _
    ByVal x As Single, _
    ByVal y As Single, _
    Optional ByVal w As Single = 160, _
    Optional ByVal nm As String = "" _
) As MSForms.label

    Dim lb As MSForms.label
    If nm <> "" Then
        Set lb = parent.controls.Add("Forms.Label.1", nm)
    Else
        Set lb = parent.controls.Add("Forms.Label.1")
    End If

    With lb
        .caption = caption
        .Left = x
        .Top = y
        .AutoSize = False
        .Width = w
    End With

    Set CreateLabel = lb
End Function



Function CreateLabelXY( _
    parent As MSForms.Frame, _
    ByVal x As Single, _
    ByVal y As Single, _
    Optional ByVal caption As String = "", _
    Optional ByVal nm As String = "", _
    Optional ByVal w As Single = 160 _
) As MSForms.label
    Set CreateLabelXY = CreateLabel(parent, caption, x, y, w, nm)
End Function




Private Function CreateTextBox(parent As MSForms.Frame, x As Single, y As Single, _
                               w As Single, h As Single, multiline As Boolean, _
                               Optional name As String = "", Optional tag As String = "") As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = parent.controls.Add("Forms.TextBox.1", IIf(name = "", vbNullString, name))
    With tb
        .Left = x
        .Top = y
        .Width = w
        .Height = IIf(h > 0, h, 20)
        .multiline = multiline
        .EnterKeyBehavior = multiline
        .tag = tag
    End With
    Set CreateTextBox = tb
End Function

Private Function CreateCombo(parent As MSForms.Frame, x As Single, y As Single, _
                             w As Single, Optional name As String = "", Optional tag As String = "") As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    Set cb = parent.controls.Add("Forms.ComboBox.1", IIf(name = "", vbNullString, name))
    With cb
        .Left = x
        .Top = y
        .Width = w
        .Style = fmStyleDropDownList
        .tag = tag
    End With
    Set CreateCombo = cb
End Function

Private Function CreateCheck(parent As MSForms.Frame, caption As String, _
                             x As Single, y As Single, Optional name As String = "", _
                             Optional tag As String = "") As MSForms.CheckBox
    Dim ck As MSForms.CheckBox
    Set ck = parent.controls.Add("Forms.CheckBox.1", IIf(name = "", vbNullString, name))
    With ck
        .caption = caption
        .Left = x
        .Top = y
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

'=== 繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ繧剃ｸｦ縺ｹ繧区ｱ守畑繝輔Ξ繝ｼ繝 ===
Private Function BuildCheckFrame(parent As MSForms.Frame, _
    title As String, x As Single, y As Single, w As Single, _
    items As Variant, Optional groupTag As String = "") As MSForms.Frame

    Dim f As MSForms.Frame
    Set f = parent.controls.Add("Forms.Frame.1")
    With f
        .caption = title
        .Left = x
        .Top = y
        .Width = w
        .Height = 60                ' 莉ｮ鬮倥＆縲ゆｸ九〒荳ｭ霄ｫ縺ｫ蜷医ｏ縺帙※莨ｸ縺ｰ縺・
        .ScrollBars = fmScrollBarsNone
    End With

    ' 繝√ぉ繝・け繧・蛻励〒驟咲ｽｮ・磯聞縺上↑繧翫☆縺弱↑縺・ｨ句ｺｦ・・
    Dim i As Long, col As Long, row As Long
    Dim colW As Single: colW = (w - 24) / 2
    Dim rowH As Single: rowH = 20
    Dim maxRow As Long: maxRow = 0

    For i = LBound(items) To UBound(items)
        col = (i - LBound(items)) Mod 2
        row = (i - LBound(items)) \ 2
        Dim ck As MSForms.CheckBox
        Set ck = f.controls.Add("Forms.CheckBox.1", "ck_" & CStr(i))
        With ck
            .caption = CStr(items(i))
            .Left = 12 + col * colW
            .Top = 18 + row * rowH
            .tag = IIf(Len(groupTag) = 0, "", groupTag)
        End With
        If row > maxRow Then maxRow = row
    Next i

    ' 荳ｭ霄ｫ縺ｮ鬮倥＆縺ｫ蜷医ｏ縺帙※繝輔Ξ繝ｼ繝繧剃ｼｸ縺ｰ縺・
    f.Height = 18 + (maxRow + 1) * rowH + 12

    Set BuildCheckFrame = f
End Function

Private Sub BuildAssistiveChecksInWalkEval(ByVal assistiveCsv As String)
    Dim frTarget As MSForms.Frame
    Set frTarget = GetWalkEvalAssistiveTargetFrame()
    If frTarget Is Nothing Then Exit Sub

    Dim i As Long
    For i = frTarget.controls.count - 1 To 0 Step -1
        If TypeName(frTarget.controls(i)) = "CheckBox" Then
            If frTarget.controls(i).tag = "AssistiveGroup" Then
                frTarget.controls.Remove frTarget.controls(i).name
            End If
        End If
    Next

    Dim maxBottom As Single
    maxBottom = 0
    For i = 0 To frTarget.controls.count - 1
        With frTarget.controls(i)
            If .Top + .Height > maxBottom Then maxBottom = .Top + .Height
        End With
    Next

    Dim addTop As Single
    addTop = IIf(maxBottom <= 0, 120, maxBottom + 120)

    Dim frAssist As MSForms.Frame
    Set frAssist = BuildCheckFrame(frTarget, "陬懷勧蜈ｷ", 8, addTop, frTarget.InsideWidth - 16, MakeList(assistiveCsv), "AssistiveGroup")

    If frAssist.Top + frAssist.Height + 8 > frTarget.Height Then
        frTarget.Height = frAssist.Top + frAssist.Height + 8
    End If
End Sub

Private Function GetWalkEvalAssistiveTargetFrame() As MSForms.Frame
    Set GetWalkEvalAssistiveTargetFrame = GetWalkAssistiveTargetFrame()
End Function

Public Function GetWalkAssistiveTargetFrame() As MSForms.Frame
    Dim root As MSForms.Frame
    Dim c As Control
    
    On Error Resume Next
    If Not fGait Is Nothing Then
        Set GetWalkAssistiveTargetFrame = fGait
        Exit Function
    End If
    On Error GoTo 0
    
    Set root = GetWalkRootFrame()
    If root Is Nothing Then Exit Function
    
    For Each c In root.controls
        If TypeName(c) = "Frame" Then
            If InStr(1, CStr(c.name), "gait", vbTextCompare) > 0 _
               Or InStr(1, CStr(c.caption), "", vbTextCompare) > 0 Then
                Set GetWalkAssistiveTargetFrame = c
                Exit Function
            End If
        End If
    Next

    Set GetWalkAssistiveTargetFrame = root
    
End Function

'=== 髢｢遽諡倡ｸｮ・壼ｷｦ蜿ｳ繝√ぉ繝・け繧・陦檎函謌舌☆繧九・繝ｫ繝代・ ======================
Private Sub CreateContractureRLRow(parent As MSForms.Frame, _
                                   ByRef y As Single, _
                                   ByVal partCaption As String, _
                                   ByVal baseTag As String)
    ' 繧ｬ繧､繝会ｼ壼・鬆ｭ縺ｧ隕句・縺励ｒ菴懊▲縺溘Ξ繧､繧｢繧ｦ繝医↓蜷医ｏ縺帙ｋ
    '   驛ｨ菴・ COL_LX
    '   蜿ｳ  : COL_LX + 90 + 20
    '   蟾ｦ  : ・亥承縺ｮ蛻暦ｼ・ 60

    ' 隕句・縺暦ｼ磯Κ菴榊錐・・
    Call CreateLabel(parent, partCaption, COL_LX, y)

    ' 蜿ｳ繝√ぉ繝・け
    Call CreateCheck(parent, "蜿ｳ", COL_LX + 90 + 20, y, , baseTag & ".蜿ｳ")

    ' 蟾ｦ繝√ぉ繝・け
    Call CreateCheck(parent, "蟾ｦ", COL_LX + 90 + 20 + 60, y, , baseTag & ".蟾ｦ")

    ' 谺｡縺ｮ陦後∈
    nL y
End Sub




'=== RLA繝√ぉ繝・け鄒､繧剃ｽ懊ｋ譛蟆丞ｮ溯｣・======================================
Private Sub Build_RLA_ChecksPart(f As MSForms.Frame, ByVal kind As String)
    Dim phases As Variant
    If LCase$(kind) = "stance" Then
        phases = Array("IC", "LR", "MSt", "TSt")        ' 遶玖・譛・
    Else
        phases = Array("PSw", "ISw", "MSw", "TSw")      ' 驕願・譛・
    End If

    Dim i As Long, y As Single
    y = 22

    For i = LBound(phases) To UBound(phases)
        Dim key As String: key = CStr(phases(i))

        ' 蟾ｦ遶ｯ・壹ヵ繧ｧ繝ｼ繧ｺ蜷阪Λ繝吶Ν
        Call CreateLabel(f, RLAPhaseCaption(key), 12, y, 90)

        ' 邁｡譏薙メ繧ｧ繝・け・・縺､・俄ｻ蜷榊燕縺ｯ "RLA_<key>_<逡ｪ蜿ｷ>" 縺ｨ縺吶ｋ
        Call AddRLAChk(f, key, "蜿ｯ蜍募沺荳崎ｶｳ", 120, y)
        Call AddRLAChk(f, key, "遲句鴨菴惹ｸ・, 300, y)
        y = y + 22
        Call AddRLAChk(f, key, "逍ｼ逞・荳榊ｮ牙ｮ・, 120, y)
        Call AddRLAChk(f, key, "蜊碑ｪｿ荳崎憶", 300, y)

        ' 蜿ｳ遶ｯ・壹Ξ繝吶Ν驕ｸ謚橸ｼ・ptionButton, GroupName=key・・
        Call AddRLAOpt(f, key, "霆ｽ蠎ｦ", f.Width - 180, y - 22)
        Call AddRLAOpt(f, key, "荳ｭ遲牙ｺｦ", f.Width - 120, y - 22)
        Call AddRLAOpt(f, key, "鬮伜ｺｦ", f.Width - 60, y - 22)

        y = y + 30
    Next

    ' 繝輔Ξ繝ｼ繝縺ｮ鬮倥＆縺ｯ蜻ｼ縺ｳ蜃ｺ縺怜・縺ｧ ResizeFrameToContent 縺励※縺・ｋ縺溘ａ縺薙％縺ｧ縺ｯ隗ｦ繧峨↑縺・
End Sub

' 繝輔ぉ繝ｼ繧ｺ蜷搾ｼ郁ｦ句・縺礼畑・・
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

' 繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ霑ｽ蜉・亥錐蜑阪・ RLA_<key>_n・・
Private Sub AddRLAChk(f As MSForms.Frame, ByVal key As String, ByVal caption As String, _
                      ByVal x As Single, ByVal y As Single)
    Dim ck As MSForms.CheckBox
    Set ck = f.controls.Add("Forms.CheckBox.1", "RLA_" & key & "_" & Replace(caption, "/", "_"))
    With ck
        .caption = caption
        .Left = x
        .Top = y
    End With
End Sub

' 繝ｬ繝吶Ν逕ｨ繧ｪ繝励す繝ｧ繝ｳ繝懊ち繝ｳ・・roupName=key・・
Private Sub AddRLAOpt(f As MSForms.Frame, ByVal key As String, ByVal caption As String, _
                      ByVal x As Single, ByVal y As Single)
    Dim ob As MSForms.OptionButton
    Set ob = f.controls.Add("Forms.OptionButton.1")
    With ob
        .caption = caption
        .groupName = key
        .Left = x
        .Top = y
    End With
End Sub

' 蜻ｼ縺ｳ蜃ｺ縺怜・縺ｧ蜿ら・縺励※縺・ｋ縺溘ａ縺ｮ莠呈鋤繝繝溘・・域里螳夐∈謚槭・迚ｹ縺ｫ險ｭ螳壹＠縺ｪ縺・ｼ・
Private Sub InitRLAdefaults()
    ' no-op
End Sub
'======================================================================

'=== 縺薙％縺九ｉ・壹・繝・ム讀懃ｴ｢邉ｻ縺ｮ莠呈鋤繝ｩ繝・ヱ繝ｼ =========================
' Local螳溯｣・ｒ譌｢縺ｫ蜈･繧後※縺・ｋ蜑肴署・・uildHeaderLookupLocal / ResolveColumnLocal / EnsureHeaderColumnLocal・・
' 蜻ｼ縺ｳ蜃ｺ縺怜・縺ｨ繧ｷ繧ｰ繝阪メ繝｣繧貞粋繧上○繧九◆繧√・阮・＞繝ｩ繝・ヱ繝ｼ縺縺醍畑諢・

' 1) BuildHeaderLookup 縺ｮ蛻･蜷阪Λ繝・ヱ繝ｼ
Private Function BuildHeaderLookup(ByVal ws As Worksheet) As Object
    Set BuildHeaderLookup = BuildHeaderLookupLocal(ws)
End Function

' 2) ResolveColumn 縺ｮ蛻･蜷阪Λ繝・ヱ繝ｼ
Private Function ResolveColumn(ByVal look As Object, ByVal key As String) As Long
    ResolveColumn = ResolveColumnLocal(look, key)
End Function

' 3) ResolveColOrCreate
'    隨ｬ1蜆ｪ蜈医く繝ｼ縺檎┌縺代ｌ縺ｰ繧ｨ繧､繝ｪ繧｢繧ｹ繧帝・↓謗｢縺励√◎繧後〒繧ら┌縺代ｌ縺ｰ隨ｬ1蜆ｪ蜈医く繝ｼ縺ｧ蛻励ｒ譁ｰ隕丈ｽ懈・
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

'=== ComboBox 縺ｫ驟榊・繧堤｢ｺ螳溘↓豬√＠霎ｼ繧譛蟆上・繝ｫ繝代・ ===
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



'=== 繝ｬ繧､繧｢繧ｦ繝郁・蜍輔ヵ繧｣繝・ヨ・亥ｮ牙・繧ｬ繝ｼ繝我ｻ倥″・・===
Private Sub FitLayout()
    On Error Resume Next

    ' 繝ｫ繝ｼ繝医・譛牙柑蟇ｸ豕包ｼ井ｸ矩剞縺ｧ繧ｯ繝ｩ繝ｳ繝暦ｼ・
    Dim iw As Single, iH As Single
    iw = Me.InsideWidth - 12
    iH = Me.InsideHeight - 60
    If iw < 240 Then iw = 240          ' 蟷・・荳矩剞
    If iH < 180 Then iH = 180          ' 鬮倥＆縺ｮ荳矩剞・遺・繧ｳ繧ｳ縺・0 莉･荳九↓縺ｪ繧九→ 380・・

    ' 繝ｫ繝ｼ繝・MultiPage
    If Not mp Is Nothing Then
        mp.Left = 6: mp.Top = 6
        mp.Width = iw
        mp.Height = iH
    End If

    ' 蜷・・繧ｹ繝医ヵ繝ｬ繝ｼ繝繧ょ酔讒倥↓繧ｯ繝ｩ繝ｳ繝・
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

    ' 縲碁哩縺倥ｋ縲阪・繧ｿ繝ｳ
    If Not btnCloseCtl Is Nothing Then
        btnCloseCtl.Left = Me.InsideWidth - btnCloseCtl.Width - 10
        btnCloseCtl.Top = 6 + iH + 8
    End If

    Me.ScrollBars = fmScrollBarsNone
    Me.ScrollHeight = Me.InsideHeight
End Sub


'=== 莠呈鋤繝繝溘・・壹ち繝夜・Μ繧ｻ繝・ヨ・井ｽ輔ｂ縺励↑縺・ｼ・=================
Private Sub ResetTabOrder()
End Sub
'============================================================

'=== 豌丞錐縺ｧ蛟呵｣懆｡後ｒ髮・ａ縺ｦ Variant 驟榊・縺ｫ縺励※霑斐☆ =====================
Private Function CollectCandidatesByNameLocal(ByVal ws As Worksheet, _
                                              ByVal look As Object, _
                                              ByVal pname As String) As Variant
    Dim nameCol As Long
    nameCol = ResolveColumnLocal(look, "Basic.Name")
    If nameCol = 0 Then nameCol = ResolveColumnLocal(look, "豌丞錐")
    If nameCol = 0 Then nameCol = ResolveColumnLocal(look, "Name")
    If nameCol = 0 Then Exit Function

    Dim key As String: key = NormName(pname)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, nameCol).End(xlUp).row
    Dim tmp As New Collection
    Dim r As Long, nm As String
    For r = 2 To lastRow
        nm = CStr(ws.Cells(r, nameCol).value)
        If NormName(nm) = key Then tmp.Add r
    Next

    If tmp.count = 0 Then Exit Function

    Dim a() As Long: ReDim a(1 To tmp.count)
    For r = 1 To tmp.count
        a(r) = CLng(tmp(r))
    Next
    CollectCandidatesByNameLocal = a
End Function





' 窶ｻ譌｢縺ｫ蜷悟錐縺ｮ髢｢謨ｰ縺後≠繧後・縺昴■繧峨ｒ菴ｿ縺｣縺ｦ縺上□縺輔＞
Private Function NormName(ByVal s As String) As String
    s = Replace(s, vbCrLf, "")
    s = Replace(s, " ", "")
    s = Replace(s, "縲", "")
    On Error Resume Next
    s = StrConv(s, vbNarrow)    ' 蜈ｨ隗停・蜊願ｧ抵ｼ育腸蠅・↓繧医ｊ螟ｱ謨励＠縺ｦ繧０K・・
    On Error GoTo 0
    NormName = LCase$(s)
End Function





'=== 譁・ｭ怜・縺ｮ豁｣隕丞喧・壼濠隗・蜈ｨ隗偵・菴吝・繧ｹ繝壹・繧ｹ繝ｻ螟ｧ蟆上ｒ蜷ｸ蜿弱＠縺ｦ辣ｧ蜷育畑繧ｭ繝ｼ縺ｫ縺吶ｋ ===
Private Function KeyNormalize(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)

    ' 謾ｹ陦後ｄ蜈ｨ隗偵せ繝壹・繧ｹ繧帝壼ｸｸ繧ｹ繝壹・繧ｹ縺ｫ
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(&H3000), " ") ' 蜈ｨ隗偵せ繝壹・繧ｹ竊貞濠隗・

    ' 騾｣邯壹せ繝壹・繧ｹ蝨ｧ邵ｮ・・燕蠕後ヨ繝ｪ繝
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)

    ' 蜈ｨ隗停・蜊願ｧ抵ｼ・SCII/謨ｰ蟄・繧ｫ繧ｿ繧ｫ繝翫↑縺ｩ蟇ｾ雎｡・・
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    ' 繧医￥縺ゅｋ繝上う繝輔Φ鬘槭・邨ｱ荳・・?・坂絶・-・・
    s = Replace(s, "・・, "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "窶・, "-")

    ' 螟ｧ譁・ｭ怜喧・郁恭蟄励ｆ繧峨℃蟇ｾ遲厄ｼ・
    s = UCase$(s)

    ' 辣ｧ蜷域凾縺ｯ繧ｹ繝壹・繧ｹ辟｡隕・
    s = Replace(s, " ", "")

    KeyNormalize = s
End Function

'=== 蜈･蜉帙Δ繝ｼ繝会ｼ壽律譛ｬ隱・蜊願ｧ偵・閾ｪ蜍募・譖ｿ ===============================

' 逕ｻ髱｢蜈ｨ菴薙・繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ縺ｫ IME 繝｢繝ｼ繝峨ｒ驕ｩ逕ｨ・亥・蟶ｰ・・
Private Sub SetupInputModesJP()
    ApplyInputModeJP Me

    If FnHasControl("txtAge") Then
    Debug.Print "[IME] txtAge =", Me.controls("txtAge").IMEMode
Else
    Debug.Print "[IME] txtAge = (not found)"
End If

If FnHasControl("txtEDate") Then
    Debug.Print "[IME] txtEDate =", Me.controls("txtEDate").IMEMode
Else
    Debug.Print "[IME] txtEDate = (not found)"
End If

If FnHasControl("txtPost_Note") Then
    Debug.Print "[IME] txtPost_Note =", Me.controls("txtPost_Note").IMEMode
Else
    Debug.Print "[IME] txtPost_Note = (not found)"
End If

End Sub

'=== 繝倥Ν繝代・・壹さ繝ｳ繝医Ο繝ｼ繝ｫ蟄伜惠繝√ぉ繝・け・郁｡晉ｪ∝屓驕ｿ縺ｮ蛻･蜷搾ｼ・===
Private Function FnHasControl(ByVal nm As String) As Boolean
    Dim c As MSForms.Control
    For Each c In Me.controls
        If StrComp(c.name, nm, vbTextCompare) = 0 Then
            FnHasControl = True
            Exit Function
        End If
    Next
    FnHasControl = False
End Function








'=== IME蛻・崛・壹さ繝ｳ繝・リ繧貞・蟶ｰ逧・↓蜃ｦ逅・ｼ・ultiPage蟇ｾ蠢懃沿・・==============
Private Sub ApplyInputModeJP(container As Object)
    Dim typ As String
    On Error Resume Next
    typ = TypeName(container)
    On Error GoTo 0

    If typ = "MultiPage" Then
    Dim pg As MSForms.page
    For Each pg In container.Pages
        ApplyInputModeJP pg
    Next
    Exit Sub
      End If


    ' Controls 繧呈戟縺溘↑縺・ｂ縺ｮ縺ｯ邨ゆｺ・ｼ・rame/Page 莉･螟悶・螳牙・蟇ｾ遲厄ｼ・
    If Not HasControls(container) Then Exit Sub

    Dim c As MSForms.Control
    For Each c In container.controls
        Select Case TypeName(c)
            Case "TextBox", "ComboBox"
                On Error Resume Next
                If ShouldBeNumericField(c) Then
                    c.IMEMode = fmIMEModeDisable     ' 蜊願ｧ定恭謨ｰ
                Else
                    c.IMEMode = fmIMEModeHiragana     ' 縺ｲ繧峨′縺ｪ
                End If
                On Error GoTo 0

            Case "Frame", "Page"
                ApplyInputModeJP c                   ' 蟄舌∈蜀榊ｸｰ

            Case "MultiPage"
                ' 蟄舌↓ MultiPage 縺後・繧我ｸ九′縺｣縺ｦ縺・ｋ蝣ｴ蜷医・ Pages 繧貞屓縺・
                Dim p As MSForms.page
                For Each p In c.Pages
                    ApplyInputModeJP p
                Next
        End Select
    Next c
End Sub


' 縺薙・繝輔か繝ｼ繝縺ｧ縲梧焚蟄玲ｬ・阪→隕九↑縺吝愛螳・
Private Function ShouldBeNumericField(c As MSForms.Control) As Boolean
    Dim nm As String: nm = LCase$(c.name & "")
    Dim tg As String: tg = LCase$(c.tag & "")

    ' 蜷榊燕縺ｧ蛻､螳夲ｼ亥ｿ・ｦ√↑繧峨％縺薙↓霑ｽ險假ｼ・
    If nm = "txtage" Or nm = "txtedate" Or nm = "txtonset" _
       Or nm = "txttenmwalk" Or nm = "txttug" Or nm = "txtfivests" _
       Or nm = "txtsemi" Or nm = "txtgripr" Or nm = "txtgripl" _
       Or nm = "txtbi" Or nm = "txtpid" Then
        ShouldBeNumericField = True: Exit Function
    End If

    ' Tag縺ｧ蛻､螳夲ｼ域里蟄倥・Tag繧呈ｴｻ逕ｨ・・
    If InStr(tg, "evaldate") > 0 Or InStr(tg, "onsetdate") > 0 Then ShouldBeNumericField = True: Exit Function
    If Left$(tg, 5) = "test." Or Left$(tg, 5) = "grip." Then ShouldBeNumericField = True: Exit Function
    If Right$(tg, 4) = ".age" Or tg = "bi.total" Then ShouldBeNumericField = True: Exit Function
End Function

' ・井ｻｻ諢擾ｼ我ｿ晏ｭ伜燕縺ｫ謨ｰ蟄玲ｬ・ｒ蜊願ｧ偵↓邨ｱ荳縺励※縺翫￥
Private Sub NormalizeNumericInputsToHalfwidth()
    Dim c As MSForms.Control
    For Each c In Me.controls
        Call NormalizeNumericInContainer(c)
    Next
End Sub

'=== 謨ｰ蛟､谺・・蜊願ｧ堤ｵｱ荳・哺ultiPage蟇ｾ蠢懃沿 ================================
Private Sub NormalizeNumericInContainer(container As Object)
    Dim typ As String
    On Error Resume Next
    typ = TypeName(container)
    On Error GoTo 0

    If typ = "MultiPage" Then
        Dim pg As MSForms.page
        For Each pg In container.Pages
            NormalizeNumericInContainer pg
        Next
        Exit Sub
    End If

    If Not HasControls(container) Then Exit Sub

    Dim c As MSForms.Control
    For Each c In container.controls
        Select Case TypeName(c)
            Case "TextBox", "ComboBox"
                If ShouldBeNumericField(c) Then
                    On Error Resume Next
                    c.text = StrConv(c.text, vbNarrow) ' 蜈ｨ隗停・蜊願ｧ・
                    On Error GoTo 0
                End If

            Case "Frame", "Page"
                NormalizeNumericInContainer c

            Case "MultiPage"
                Dim p As MSForms.page
                For Each p In c.Pages
                    NormalizeNumericInContainer p
                Next
        End Select
    Next c
End Sub


Private Function HasControls(o As Object) As Boolean
    On Error Resume Next
    Dim t As Long: t = o.controls.count
    HasControls = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
'====================================================================

'=== ComboBox 縺ｫ驟榊・繧堤｢ｺ螳溘↓豬√＠霎ｼ繧譛蟆上・繝ｫ繝代・ ===
Private Sub FillComboItems(ByRef cbo As MSForms.ComboBox, ByVal items As Variant)
    Dim k As Long
    cbo.Clear
    If IsArray(items) Then
        For k = LBound(items) To UBound(items)
            cbo.AddItem CStr(items(k))
        Next
    End If
    ' 譌｢螳夐∈謚槭↑縺・
    On Error Resume Next
    cbo.ListIndex = -1
End Sub



'=== BI/IADL 繧貞ｼｷ蛻ｶ蜀肴ｧ狗ｯ会ｼ亥ｿ・★荳ｭ霄ｫ繧定｡ｨ遉ｺ・・===============================
Private Function EnsureBI_IADL() As MSForms.MultiPage
    On Error Resume Next

    Trace "EnsureBI_IADL start", "BI/IADL"   ' 竊絶蔵縺薙％




    
    Dim mpADL As MSForms.MultiPage
    Dim nextTop As Single
    

 

 ' 1) 繝ｫ繝ｼ繝・MultiPage 縺ｨ縲梧律蟶ｸ逕滓ｴｻ蜍穂ｽ懊阪・繝ｼ繧ｸ繧堤音螳・
    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Function

    Dim pgMove As MSForms.page: Set pgMove = FindPageByCaption(mp, "譌･蟶ｸ逕滓ｴｻ蜍穂ｽ・)
    If pgMove Is Nothing Then
        If mp.Pages.count >= 3 Then
            Set pgMove = mp.Pages(2) ' 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ
        Else
            Exit Function
        End If
    End If

    ' 2) 繝壹・繧ｸ蜀・・繝帙せ繝・Frame 繧貞叙蠕暦ｼ育┌縺代ｌ縺ｰ菴懈・・・
    Dim host As MSForms.Frame, c As Control
    For Each c In pgMove.controls
        If TypeName(c) = "Frame" Then Set host = c: Exit For
    Next
    If host Is Nothing Then
        Set host = pgMove.controls.Add("Forms.Frame.1", "frMoveHost")
        host.Left = 6: host.Top = 6
        host.Width = pgMove.InsideWidth - 12
        host.Height = pgMove.InsideHeight - 12
    End If

    ' 3) host 蜀・・ MultiPage 縺縺代ｒ蜈ｨ豸亥悉・医⊇縺九・隗ｦ繧峨↑縺・ｼ・
    Dim i As Long
    For i = host.controls.count - 1 To 0 Step -1
        If TypeName(host.controls(i)) = "MultiPage" Then
            host.controls.Remove host.controls(i).name
        End If
    Next

    ' 4) mpADL 繧剃ｽ懈・・・譫壻ｿ晁ｨｼ・・:BI / 1:IADL / 2:襍ｷ螻・虚菴懶ｼ・
    Set mpADL = host.controls.Add("Forms.MultiPage.1", "mpADL")
    Trace "mpADL ready; pages=" & mpADL.Pages.count, "BI/IADL"

    With mpADL
        .Left = 12
        .Top = pad
        .Width = host.InsideWidth - 24
        .Height = 300
        .Style = fmTabStyleTabs
        AttachMPHook mpADL
    End With
    Do While mpADL.Pages.count < 3: mpADL.Pages.Add: Loop
    mpADL.Pages(0).caption = "繝舌・繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ"
    mpADL.Pages(1).caption = "IADL"
    mpADL.Pages(2).caption = "襍ｷ螻・虚菴・
    
    
Trace "EnsureBI_IADL end; pages=" & mpADL.Pages.count, "BI/IADL"
Set EnsureBI_IADL = mpADL

    ' 5) 襍ｷ螻・虚菴懊ち繝悶・UI
    mp.value = 0
    BuildKyoOnADL mpADL.Pages(2)


    '======================== BI・・0鬆・岼・・========================
    Dim pBI As MSForms.page: Set pBI = mpADL.Pages(0)
    ' 荳譌ｦ繧ｯ繝ｪ繧｢縺励※縺九ｉ菴懈・・育ｩｺ・城㍾隍・←縺｡繧峨↓繧ょｯｾ蠢懶ｼ・
    Dim iCtl As Long
    
     For iCtl = pBI.controls.count - 1 To 0 Step -1
    If Left(pBI.controls(iCtl).name, 5) = "lblBI" _
    Or Left(pBI.controls(iCtl).name, 5) = "cmbBI" _
    Or pBI.controls(iCtl).name = "txtBITotal" _
    Or pBI.controls(iCtl).name = "frBIHomeEnv" Then
        pBI.controls.Remove pBI.controls(iCtl).name
    End If
Next

    Dim biItems As Variant, biChoices As Variant
    biItems = Array("鞫る｣・, "霆翫＞縺・繝吶ャ繝臥ｧｻ荵・, "謨ｴ螳ｹ", "繝医う繝ｬ蜍穂ｽ・, "蜈･豬ｴ", _
                    "豁ｩ陦・霆翫＞縺咏ｧｻ蜍・, "髫取ｮｵ譏・剄", "譖ｴ陦｣", "謗剃ｾｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ", "謗貞ｰｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ")
    biChoices = Array()


    Dim yBI As Single, idx As Long
    yBI = 18

    Dim lblBI As MSForms.label, txtBI As MSForms.TextBox
    Set lblBI = pBI.controls.Add("Forms.Label.1", "lblBIHeader")
    With lblBI
        .caption = "繝舌・繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ・育せ・・
        .Left = 12: .Top = yBI: .Width = 160
    End With
    Set txtBI = pBI.controls.Add("Forms.TextBox.1", "txtBITotal")
    With txtBI
        .tag = "BI.Total"
        .Left = lblBI.Left + lblBI.Width + 8
        .Top = yBI - 3
        .Width = 60
    End With
    yBI = yBI + rowH

    For idx = LBound(biItems) To UBound(biItems)
        Dim lb As MSForms.label, cb As MSForms.ComboBox
        Set lb = pBI.controls.Add("Forms.Label.1", "lblBI_" & CStr(idx))
        With lb: .caption = CStr(biItems(idx)): .Left = 12: .Top = yBI: .Width = 160: End With

        Set cb = pBI.controls.Add("Forms.ComboBox.1", "cmbBI_" & CStr(idx))
        AttachBIHook cb
        With cb
            .tag = "BI." & CStr(biItems(idx))
            .Left = 190
            .Top = yBI - 3
            .Width = 200
            .Style = fmStyleDropDownList
        End With
               ' 鬆・岼縺斐→縺ｫ繝舌・繧ｵ繝ｫ讓呎ｺ悶・轤ｹ謨ｰ蛟呵｣懊ｒ險ｭ螳・
        Select Case idx
            ' 0: 鞫る｣・
            ' 3: 繝医う繝ｬ蜍穂ｽ・
            ' 6: 髫取ｮｵ譏・剄
            ' 7: 譖ｴ陦｣
            ' 8: 謗剃ｾｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ
            ' 9: 謗貞ｰｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ
            ' 竊・0 / 5 / 10 轤ｹ
            Case 0, 3, 6, 7, 8, 9
                AddItemsToCombo cb, Array("0", "5", "10")

            ' 2: 謨ｴ螳ｹ
            ' 4: 蜈･豬ｴ
            ' 竊・0 / 5 轤ｹ
            Case 2, 4
                AddItemsToCombo cb, Array("0", "5")

            ' 1: 霆翫＞縺・繝吶ャ繝臥ｧｻ荵・
            ' 5: 豁ｩ陦・霆翫＞縺咏ｧｻ蜍・
            ' 竊・0 / 5 / 10 / 15 轤ｹ
            Case 1, 5
                AddItemsToCombo cb, Array("0", "5", "10", "15")
        End Select

        yBI = yBI + rowH
    Next idx
    
   Dim frHomeEnv As MSForms.Frame
    Dim lblHomeNote As MSForms.label, txtHomeNote As MSForms.TextBox
    Dim homeItems As Variant, homeNames As Variant, h As Long
    Dim ckHome As MSForms.CheckBox
    Dim yHome As Single

Set frHomeEnv = pBI.controls.Add("Forms.Frame.1", "frBIHomeEnv")
With frHomeEnv
    .caption = "菴丞ｮ・憾豕・ｼ郁ｩｲ蠖薙・縺ｿ繝√ぉ繝・け・・
    .Left = 420
    .ZOrder 0
    .Top = yBI + 6
    .Width = 390
End With

homeItems = Array( _
    "邇・未縺ｾ縺ｧ縺ｮ谿ｵ蟾ｮ", _
    "荳翫′繧頑｡・ｼ域ｮｵ蟾ｮ・・, _
    "螻句・縺ｮ谿ｵ蟾ｮ", _
    "髫取ｮｵ縺ゅｊ", _
    "謇九☆繧翫≠繧・, _
    "繧ｹ繝ｭ繝ｼ繝励≠繧・, _
    "騾夊ｷｯ縺檎強縺・ _
)

homeNames = Array( _
    "chkBIHomeEnv_Entrance", _
    "chkBIHomeEnv_Genkan", _
    "chkBIHomeEnv_IndoorStep", _
    "chkBIHomeEnv_Stairs", _
    "chkBIHomeEnv_Handrail", _
    "chkBIHomeEnv_Slope", _
    "chkBIHomeEnv_NarrowPath" _
)

yHome = 18
For h = LBound(homeItems) To UBound(homeItems)
    Set ckHome = frHomeEnv.controls.Add("Forms.CheckBox.1", CStr(homeNames(h)))

    Dim col As Long, row As Long
    col = h Mod 2
    row = h \ 2

    With ckHome
        .caption = CStr(homeItems(h))
        .Left = 12 + col * (frHomeEnv.InsideWidth / 2)
        .Top = 18 + row * 20
        .Width = (frHomeEnv.InsideWidth / 2) - 18
        .tag = "BI.HomeEnv." & CStr(h)   '竊慎ag縺ｯ縺昴・縺ｾ縺ｾ
        .value = False
    End With
Next h

yHome = 18 + ((UBound(homeItems) + 1 + 1) \ 2) * 20

Set lblHomeNote = frHomeEnv.controls.Add("Forms.Label.1", "lblBIHomeEnvNote")
With lblHomeNote
    .caption = "蛯呵・
    .Left = 12
    .Top = yHome + 4
    .Width = 60
End With

Set txtHomeNote = frHomeEnv.controls.Add("Forms.TextBox.1", "txtBIHomeEnvNote")
With txtHomeNote
    .tag = "BI.HomeEnv.Note"           '竊慎ag縺ｯ縺昴・縺ｾ縺ｾ
    .Left = 12
    .Top = lblHomeNote.Top + 14
    .Width = frHomeEnv.InsideWidth - 18
    .Height = 100
    .multiline = True
    .EnterKeyBehavior = True
End With

'笘・鬮倥＆繧偵％縺薙〒遒ｺ螳夲ｼ亥崋螳・30繧偵ｄ繧√ｋ・・
frHomeEnv.Height = txtHomeNote.Top + txtHomeNote.Height + 12
  
    pBI.ScrollBars = fmScrollBarsNone
    pBI.ScrollHeight = yBI + 8


    '======================== IADL・・鬆・岼・・========================
    Dim pIADL As MSForms.page: Set pIADL = mpADL.Pages(1)
    For iCtl = pIADL.controls.count - 1 To 0 Step -1
        pIADL.controls.Remove pIADL.controls(iCtl).name
    Next

    Dim iadlItems As Variant, iadlChoices As Variant
    iadlItems = Array("隱ｿ逅・, "豢玲ｿｯ", "謗・勁", "雋ｷ縺・黄", "驥鷹姦邂｡逅・, "譛崎脈邂｡逅・, _
                      "雜｣蜻ｳ繝ｻ菴呎嚊豢ｻ蜍・, "遉ｾ莨壼盾蜉・亥､門・繝ｻ蝨ｰ蝓滓ｴｻ蜍包ｼ・, "繧ｳ繝溘Η繝九こ繝ｼ繧ｷ繝ｧ繝ｳ・磯崕隧ｱ繝ｻ莨夊ｩｱ・・)
    iadlChoices = Array("閾ｪ遶・, "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧", "蜈ｨ莉句勧")

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
        Set lb2 = pIADL.controls.Add("Forms.Label.1", "lblIADL_" & CStr(j))
        With lb2: .caption = CStr(iadlItems(j)): .Left = xIADL: .Top = yIADL: .Width = 120: End With

        Set cb2 = pIADL.controls.Add("Forms.ComboBox.1", "cmbIADL_" & CStr(j))
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
    Set lblINote = pIADL.controls.Add("Forms.Label.1", "lblIADLNote")
    With lblINote: .caption = "蛯呵・: .Left = 12: .Top = gridBottom + 12: .Width = 40: End With
    Set txtINote = pIADL.controls.Add("Forms.TextBox.1", "txtIADLNote")
    With txtINote
        .tag = "IADL.蛯呵・
        .Left = 60
        .Top = lblINote.Top - 3
        .Width = mpADL.Width - 84
        .Height = 60
        .multiline = True
        .EnterKeyBehavior = True
    End With

    '---- 鬮倥＆譖ｴ譁ｰ ----
    Dim bottomI As Single: bottomI = txtINote.Top + txtINote.Height + 18
    pIADL.ScrollBars = fmScrollBarsNone
    pIADL.ScrollHeight = bottomI

    If mpADL.Height < bottomI + 42 Then mpADL.Height = bottomI + 42
    
   
    nextTop = mpADL.Top + mpADL.Height + 10
    If Not hostMove Is Nothing Then hostMove.ScrollHeight = nextTop + 10

    '--- BI: 菴丞ｮ・憾豕√ｒ蜿ｳ蛛ｴ縺ｸ蝗ｺ螳夲ｼ域怙蠕後↓蠖薙※繧具ｼ壽紛蛻励〒謌ｻ縺輔ｌ縺ｪ縺・ｈ縺・↓・・--
On Error Resume Next
pBI.controls("frBIHomeEnv").Left = 600
pBI.controls("frBIHomeEnv").Top = 12
pBI.controls("frBIHomeEnv").ZOrder 0
On Error GoTo 0


    Set EnsureBI_IADL = mpADL
End Function
'=================================================================

'=== BI繧ｳ繝ｳ繝懊↓繧､繝吶Φ繝医ヵ繝・け繧貞ｼｵ繧・===
Private Sub AttachBIHook(ByRef cb As MSForms.ComboBox)
    If BIHooks Is Nothing Then Set BIHooks = New Collection
    Dim h As CboBIHook
    Set h = New CboBIHook
    h.Init Me, cb
    BIHooks.Add h
End Sub

'=== BI・夐・岼ﾃ鈴∈謚櫁い 竊・轤ｹ謨ｰ・・arthel讓呎ｺ悶↓豐ｿ縺・ｼ・==================
Private Function BIItemScore(ByVal itemName As String, ByVal level As String) As Long
    Select Case itemName
        Case "鞫る｣・                         ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "霆翫＞縺・繝吶ャ繝臥ｧｻ荵・            ' 15 / 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 15
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・:       BIItemScore = 10
                Case "荳驛ｨ莉句勧":               BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "謨ｴ螳ｹ"                         ' 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "繝医う繝ｬ蜍穂ｽ・                   ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "蜈･豬ｴ"                         ' 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "豁ｩ陦・霆翫＞縺咏ｧｻ蜍・              ' 15 / 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 15
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・:       BIItemScore = 10
                Case "荳驛ｨ莉句勧":               BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "髫取ｮｵ譏・剄"                     ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "譖ｴ陦｣"                         ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "謗剃ｾｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ"             ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
                Case Else:                     BIItemScore = 0
            End Select

        Case "謗貞ｰｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ"             ' 10 / 5 / 0
            Select Case level
                Case "閾ｪ遶・:                   BIItemScore = 10
                Case "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧": BIItemScore = 5
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
    Dim pBI As MSForms.page
    Dim idx As Long
    Dim cb As MSForms.ComboBox
    Dim total As Long
    Dim v As String
    Dim txt As MSForms.TextBox

    On Error Resume Next

    ' --- 縲後ヰ繝ｼ繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ縲阪ち繝悶ｒ謖√▽ MultiPage 繧呈爾縺・---
    Set mpADL = Nothing
    For Each ctrl In Me.controls
        If TypeOf ctrl Is MSForms.MultiPage Then
            Set mpADL = ctrl
            If mpADL.Pages.count > 0 Then
                If mpADL.Pages(0).caption = "繝舌・繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ" Then
                    Exit For
                End If
            End If
            Set mpADL = Nothing
        End If
    Next ctrl

    If mpADL Is Nothing Then Exit Sub

    Set pBI = mpADL.Pages(0)   ' 0繝壹・繧ｸ逶ｮ縺後ヰ繝ｼ繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ

    ' --- 10鬆・岼蛻・・轤ｹ謨ｰ繧貞腰邏斐↓蜷郁ｨ医☆繧具ｼ医さ繝ｳ繝懊・蛟､縺後◎縺ｮ縺ｾ縺ｾ轤ｹ謨ｰ・・---
    total = 0
    For idx = 0 To 9
        Set cb = Nothing
        Set cb = pBI.controls("cmbBI_" & CStr(idx))
        If Not cb Is Nothing Then
            v = Trim$(CStr(cb.value))
            If Len(v) > 0 Then
                total = total + CLng(val(v))
            End If
        End If
    Next idx

    ' --- 蜷郁ｨ医ｒ txtBITotal 縺ｫ蜿肴丐 ---
    Set txt = pBI.controls("txtBITotal")
    If Not txt Is Nothing Then
        txt.value = total
    End If
End Sub


'=== 莉ｻ諢上・TextBox縺ｫIME縺ｲ繧峨′縺ｪ繝輔ャ繧ｯ繧貞ｼｵ繧・===
Private Sub AttachImeHiragana(tb As MSForms.TextBox)
    If ImeHooks Is Nothing Then Set ImeHooks = New Collection
    Dim h As TxtImeHook
    Set h = New TxtImeHook
    h.Init tb
    ImeHooks.Add h
    ' 蠢ｵ縺ｮ縺溘ａ逶ｴ縺｡縺ｫ蜿肴丐
    On Error Resume Next
    tb.IMEMode = fmIMEModeHiragana
End Sub

'=== MultiPage 縺ｫ繝輔ャ繧ｯ繧貞ｼｵ繧・===
Private Sub AttachMPHook(mp As MSForms.MultiPage)
    If MPHs Is Nothing Then Set MPHs = New Collection
    Dim h As MPHook
    Set h = New MPHook
    h.Init Me, mp
    MPHs.Add h
End Sub

'=== IADL蛯呵・↓IME縺ｲ繧峨′縺ｪ繧貞・驕ｩ逕ｨ・磯・蠎ｦ蜻ｼ縺ｶ・・===
Public Sub ApplyImeToIADLNote()
    On Error Resume Next
     Dim mpA As MSForms.MultiPage, c As Control
     If hostMove Is Nothing Then Exit Sub

    For Each c In hostMove.controls
        If TypeName(c) = "MultiPage" Then
            If c.name = "mpADL" Then Set mpA = c: Exit For
        End If
    Next c
    If mpA Is Nothing Then Exit Sub

    Dim tb As MSForms.TextBox
    Set tb = mpA.Pages(1).controls("txtIADLNote") ' Page(1) = IADL
    If Not tb Is Nothing Then tb.IMEMode = fmIMEModeHiragana

   
End Sub

Private Sub EnsureMpPhysChangeHook_Once()
    If mMPPhysHookAttached Then Exit Sub

    Dim mpPhysObj As Object
    Set mpPhysObj = SafeGetControl(Me, "mpPhys")
    If mpPhysObj Is Nothing Then Exit Sub

    AttachMPHook mpPhysObj
    mMPPhysHookAttached = True
End Sub

Public Sub RunMMTBuildOnceOnMpPhysMMTActive()
    If mMMTFirstBuildDone Then Exit Sub

    Dim mpPhysObj As Object
    Set mpPhysObj = SafeGetControl(Me, "mpPhys")
    If mpPhysObj Is Nothing Then Exit Sub

    If CLng(mpPhysObj.value) <> 1 Then Exit Sub

    MMT_BuildChildTabs_Direct
    mMMTFirstBuildDone = True
    
    TryRunPendingMMTLoad
End Sub

Public Sub QueueMMTLoadAfterUI(ByVal ws As Worksheet, ByVal rowNum As Long)
    If ws Is Nothing Then Exit Sub
    If rowNum < 2 Then Exit Sub

    Set mPendingMMTWs = ws
    mPendingMMTRow = rowNum
    mPendingMMTLoad = True

    TryRunPendingMMTLoad
End Sub

Private Sub TryRunPendingMMTLoad()
    On Error GoTo EH

    If mApplyingPendingMMTLoad Then Exit Sub
    If Not mPendingMMTLoad Then Exit Sub
    If Not mMMTFirstBuildDone Then Exit Sub
    If mPendingMMTWs Is Nothing Then Exit Sub

    mApplyingPendingMMTLoad = True
    Call MMT.LoadMMTFromSheet(mPendingMMTWs, mPendingMMTRow, Me)

ExitHere:
    If Err.Number = 0 Then
        mPendingMMTLoad = False
        mPendingMMTRow = 0
        Set mPendingMMTWs = Nothing
    End If
    mApplyingPendingMMTLoad = False
    Exit Sub

EH:
    Debug.Print "[ERR][frmEval.TryRunPendingMMTLoad] Err=" & Err.Number & " Desc=" & Err.Description
    Resume ExitHere
    
End Sub

'=== 縺ｩ縺薙°縺ｫ谿九▲縺ｦ縺・ｋ mpADL 繧貞・驛ｨ豸医☆ ===
Private Sub RemoveAllMpADL()
    Dim i As Long, c As Control
    ' 繝輔か繝ｼ繝逶ｴ荳・
    For i = Me.controls.count - 1 To 0 Step -1
        If TypeName(Me.controls(i)) = "MultiPage" Then
            If Me.controls(i).name = "mpADL" Then
                Me.controls.Remove Me.controls(i).name
            End If
        End If
    Next i

    ' 繝ｫ繝ｼ繝・MultiPage・・p・峨・蜷・・繝ｼ繧ｸ蜀・
    Dim mp As MSForms.MultiPage, p As MSForms.page
    For Each c In Me.controls
        If TypeName(c) = "MultiPage" Then Set mp = c: Exit For
    Next c
    If Not mp Is Nothing Then
        For i = 0 To mp.Pages.count - 1
            For Each c In mp.Pages(i).controls
                If TypeName(c) = "MultiPage" Then
                    If c.name = "mpADL" Then mp.Pages(i).controls.Remove c.name
                End If
            Next c
        Next i
    End If
End Sub


'=== ADL繧ｿ繝門・縺ｮ3譫夂岼縲瑚ｵｷ螻・虚菴懊阪・繝ｼ繧ｸ繧堤ｵ・∩遶九※繧・======================
Private Sub BuildKyoOnADL(pg As MSForms.page)

    Dim fr As MSForms.Frame
    ' 譌｢蟄倥′縺ゅｌ縺ｰ蜀榊茜逕ｨ縲∫┌縺代ｌ縺ｰ菴懈・・・indOrAddFrame縺ｯ譌｢蟄倥・繝ｫ繝代・・・
    Set fr = FindOrAddFrame(pg, "frKyo")
    fr.caption = "襍ｷ螻・虚菴・
    ClearChildren fr
    
     ' 笘・ｿｽ蜉・壻ｽ咲ｽｮ縺ｨ繧ｵ繧､繧ｺ・医・繝ｼ繧ｸ縺・▲縺ｱ縺・↓蠎・￡繧具ｼ・
    With fr
        .Left = 12
        .Top = 12
        .Width = pg.parent.Width - 24   ' 竊・MultiPage縺ｮ蟷・↓霑ｽ蠕・
    End With
    

    Dim y As Single: y = 22
    Dim lb As MSForms.label, cb As MSForms.ComboBox, txt As MSForms.TextBox
    Dim choices As Variant

    ' 蛟呵｣懶ｼ壽里蟄倥・ PostureChoices() 縺後≠繧後・蛻ｩ逕ｨ縲∫┌縺代ｌ縺ｰ繝・ヵ繧ｩ繝ｫ繝・
    On Error Resume Next
    choices = PostureChoices()
    If Err.Number <> 0 Then
        choices = Array("閾ｪ遶・, "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧", "蜈ｨ莉句勧")
        Err.Clear
    End If
    On Error GoTo 0
    
    
' 蟇晁ｿ斐ｊ
Set lb = CreateLabel(fr, "蟇晁ｿ斐ｊ", COL_LX, y)
Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbKyo_Roll", True)
With cb: .Left = COL_LX + lblW + 60: .Top = y - 3: .Width = 120: End With
AddItemsToCombo cb, choices
y = y + rowH

' 襍ｷ縺堺ｸ翫′繧・
Set lb = CreateLabel(fr, "襍ｷ縺堺ｸ翫′繧・, COL_LX, y)
Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbKyo_SitUp", True)
With cb: .Left = COL_LX + lblW + 60: .Top = y - 3: .Width = 120: End With
AddItemsToCombo cb, choices
y = y + rowH

' 蠎ｧ菴堺ｿ晄戟
Set lb = CreateLabel(fr, "蠎ｧ菴堺ｿ晄戟", COL_LX, y)
Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbKyo_SitHold", True)
With cb: .Left = COL_LX + lblW + 60: .Top = y - 3: .Width = 120: End With
AddItemsToCombo cb, choices
y = y + rowH

    
  ' 蜿ｳ蛻暦ｼ夂ｫ九■荳翫′繧・/ 遶倶ｽ堺ｿ晄戟・亥ｷｦ蛻励→蜷後§陦刑=22,46・上が繝輔そ繝・ヨ+60・丞ｹ・20縺ｧ謠・∴繧具ｼ・
CreateLabel fr, "遶九■荳翫′繧・, COL_RX, 22
Dim cboUp As MSForms.ComboBox
Set cboUp = CreateCombo(fr, COL_RX + lblW + 60, 22, 120, , "POSTURE|遶九■荳翫′繧・)
cboUp.List = MakeList("閾ｪ遶・隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・荳驛ｨ莉句勧,蜈ｨ莉句勧")

CreateLabel fr, "遶倶ｽ堺ｿ晄戟", COL_RX, 46
Dim cboStand As MSForms.ComboBox
Set cboStand = CreateCombo(fr, COL_RX + lblW + 60, 46, 120, , "POSTURE|遶倶ｽ堺ｿ晄戟")
cboStand.List = MakeList("閾ｪ遶・隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・荳驛ｨ莉句勧,蜈ｨ莉句勧")

    ' 蛯呵・
    Set lb = CreateLabel(fr, "蛯呵・, COL_LX, y)
    Set txt = fr.controls.Add("Forms.TextBox.1", "txtKyoNote")
    With txt
        .Left = COL_LX + 40
        .Top = y - 3
        .Width = fr.InsideWidth - (COL_LX + 40) - 12
        .Height = 120
        .multiline = True
        .EnterKeyBehavior = True
    End With
    y = y + txt.Height + 12

    fr.Height = y + 12
End Sub

'ADL-襍ｷ螻・虚菴懃畑・壹さ繝ｳ繝懊↓蛟呵｣懊ｒ繧ｻ繝・ヨ
Private Sub AddItemsToCombo(cb As MSForms.ComboBox, items As Variant)
    Dim k As Long
    On Error Resume Next
    cb.Clear
    cb.Style = fmStyleDropDownList
    For k = LBound(items) To UBound(items)
        cb.AddItem CStr(items(k))
    Next k
End Sub











'======================================================================
' 襍ｷ螻・虚菴懶ｼ医瑚ｺｫ菴捺ｩ溯・・玖ｵｷ螻・虚菴懊阪・繝ｼ繧ｸ・・ 螳牙ｮ壹ユ繝ｳ繝励Ξ・亥ｙ閠・≠繧奇ｼ・
' 繝ｻ騾比ｸｭ縺ｫ繧ｿ繝悶ｒ謖ｿ蜈･縺励※繧ょ｣翫ｌ縺ｪ縺・ｼ咾aption讀懃ｴ｢
' 繝ｻ逕滓・・・uild・峨→謨ｴ蛻暦ｼ・ayout・峨ｒ蛻・屬・壽僑蠑ｵ譎ゅ・莠区腐繧呈怙蟆丞喧
' 繝ｻ菫晏ｭ・隱ｭ霎ｼ縺ｯ Tag="POSTURE|窶ｦ" 縺ｧ譌｢蟄倥Ο繧ｸ繝・け縺ｫ縺昴・縺ｾ縺ｾ荵励ｋ
'======================================================================



'窶補・繝輔か繝ｼ繝蜀・・ MultiPage 繧定・蜍墓､懷・・亥錐蜑阪↓萓晏ｭ倥＠縺ｪ縺・ｼ・
Private Function FindMainMultiPage() As MSForms.MultiPage
    Dim c As MSForms.Control
    For Each c In Me.controls
        If TypeOf c Is MSForms.MultiPage Then
            Set FindMainMultiPage = c
            Exit Function
        End If
    Next
End Function

'窶補・Caption 縺ｫ謖・ｮ壽枚蟄怜・繧貞性繧繝壹・繧ｸ繧定ｿ斐☆・育┌縺代ｌ縺ｰ Nothing・・
Private Function FindPageByCaption(mp As MSForms.MultiPage, cap As String) As MSForms.page
    Dim pg As MSForms.page
    For Each pg In mp.Pages
        If InStr(pg.caption, cap) > 0 Then
            Set FindPageByCaption = pg
            Exit Function
        End If
    Next
End Function

'窶補・Page 蜀・〒 Frame 繧貞叙蠕暦ｼ育┌縺代ｌ縺ｰ菴懈・・・
Private Function FindOrAddFrame(pg As MSForms.page, nm As String) As MSForms.Frame
    Dim c As MSForms.Control
    For Each c In pg.controls
        If TypeOf c Is MSForms.Frame Then
            If StrComp(c.name, nm, vbTextCompare) = 0 Then
                Set FindOrAddFrame = c
                Exit Function
            End If
        End If
    Next
    Set FindOrAddFrame = pg.controls.Add("Forms.Frame.1", nm, True)
End Function

'窶補・蟄舌さ繝ｳ繝医Ο繝ｼ繝ｫ蜈ｨ蜑企勁・育函謌仙燕縺ｫ荳蠎ｦ縺縺托ｼ・
Private Sub ClearChildren(fr As MSForms.Frame)
    Dim i As Long
    For i = fr.controls.count - 1 To 0 Step -1
        fr.controls.Remove fr.controls(i).name
    Next
End Sub


Private Function PostureItems() As Variant
    PostureItems = Array("蟇晁ｿ斐ｊ", "襍ｷ縺堺ｸ翫′繧・, "蠎ｧ菴堺ｿ晄戟", "遶九■荳翫′繧・, "遶倶ｽ堺ｿ晄戟")
End Function



Private Function PostureChoices() As Variant
    PostureChoices = Array("閾ｪ遶・, "隕句ｮ医ｊ・育屮隕紋ｸ具ｼ・, "荳驛ｨ莉句勧", "蜈ｨ莉句勧")
End Function
'======================================================================

'窶補・蛻晏屓縺縺鯛懃函謌絶昴☆繧具ｼ・ctivate 縺九ｉ蜻ｼ縺ｰ繧後ｋ・・
Private Sub BuildPostureUI()

Debug.Print "BuildPostureUI CALLED", Join(PostureItems, " / ")

    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Sub

    ' 隧ｲ蠖薙・繝ｼ繧ｸ繧・Caption 縺ｧ蜿門ｾ暦ｼ育┌縺代ｌ縺ｰ菴懊ｋ・・
    Dim pg As MSForms.page
    Set pg = FindPageByCaption(mp, CAP_POSTURE_PAGE)
    If pg Is Nothing Then
        Set pg = mp.Pages.Add
        pg.caption = CAP_POSTURE_PAGE
    End If
    mp.value = pg.Index   ' 縺薙・繧ｿ繝悶ｒ蜑埼擇縺ｸ

    ' 繝輔Ξ繝ｼ繝蜿門ｾ・
    Dim fr As MSForms.Frame: Set fr = FindOrAddFrame(pg, "frPosture")

    ' 譌｢蟄倥ｒ繧ｯ繝ｪ繧｢ 竊・陦檎函謌・
    ClearChildren fr
    CreatePostureRows fr    ' 繝ｩ繝吶Ν/繧ｳ繝ｳ繝・蛯呵・ｒ菴懊ｋ・亥ｺｧ讓吶・ Layout 縺ｧ隱ｿ謨ｴ・・

    ' 逕滓・逶ｴ蠕後↓荳蠎ｦ繝ｬ繧､繧｢繧ｦ繝・
    LayoutPosture
End Sub

'窶補・繝ｩ繝吶Ν・・さ繝ｳ繝懆｡鯉ｼ句ｙ閠・ｬ・ｒ菴懊ｋ・井ｽ咲ｽｮ縺ｯ縺薙％縺ｧ縺ｯ豎ｺ繧√↑縺・ｼ・
Private Sub CreatePostureRows(fr As MSForms.Frame)
    Dim items As Variant:   items = PostureItems()
    Dim choices As Variant: choices = PostureChoices()

    Dim i As Long
    For i = LBound(items) To UBound(items)
        ' 繝ｩ繝吶Ν
        Dim lb As MSForms.label
        Set lb = fr.controls.Add("Forms.Label.1", "lblPost_" & CStr(i), True)
        lb.caption = CStr(items(i))

        ' 繧ｳ繝ｳ繝懶ｼ井ｿ晏ｭ・隱ｭ霎ｼ縺ｫ菴ｿ縺・Tag 繧剃ｻ倅ｸ趣ｼ・
        Dim cb As MSForms.ComboBox
        Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbPost_" & CStr(i), True)
        cb.Style = fmStyleDropDownList
        cb.tag = POSTURE_TAG_PREFIX & CStr(items(i))

        ' 驕ｸ謚櫁い縺ｮ險ｭ螳夲ｼ壼・騾夐未謨ｰ縺後≠繧後・蜆ｪ蜈医∫┌縺代ｌ縺ｰ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ
        On Error Resume Next
        SetComboItems cb, choices
        If Err.Number <> 0 Then
            Dim k As Long: Err.Clear
            For k = LBound(choices) To UBound(choices)
                cb.AddItem CStr(choices(k))
            Next
        End If
        On Error GoTo 0
    Next i

    '窶補・蛯呵・Λ繝吶Ν・九ユ繧ｭ繧ｹ繝医・繝・け繧ｹ・・陦檎嶌蠖難ｼ・
    Dim lbNote As MSForms.label
    Set lbNote = fr.controls.Add("Forms.Label.1", "lblPost_Note", True)
    lbNote.caption = "蛯呵・

    Dim txtNote As MSForms.TextBox
    Set txtNote = fr.controls.Add("Forms.TextBox.1", "txtPost_Note", True)
    With txtNote
        .multiline = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .IMEMode = fmIMEModeHiragana   ' 竊・譌･譛ｬ隱槫・蜉幢ｼ亥・隗抵ｼ峨ｒ譏守､ｺ
        .tag = POSTURE_TAG_PREFIX & "蛯呵・
    End With
End Sub

'窶補・菴咲ｽｮ繝ｻ繧ｵ繧､繧ｺ隱ｿ謨ｴ・・esize 豈弱↓蜻ｼ縺ｶ・丞・逕滓・縺励↑縺・ｼ・
Private Sub LayoutPosture()
    Dim mp As MSForms.MultiPage: Set mp = FindMainMultiPage()
    If mp Is Nothing Then Exit Sub

    Dim pg As MSForms.page: Set pg = FindPageByCaption(mp, CAP_POSTURE_PAGE)
    If pg Is Nothing Then Exit Sub

    Dim fr As MSForms.Frame: Set fr = FindOrAddFrame(pg, "frPosture")

    ' 繝輔Ξ繝ｼ繝縺ｮ蝓ｺ貅門ｯｸ豕包ｼ・ultiPage 縺・0 縺ｮ蝣ｴ蜷医・ Form 縺九ｉ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・・
    fr.Visible = True
    fr.Left = 6: fr.Top = 6
    fr.Width = Application.Max(120, IIf(mp.Width > 0, mp.Width, Me.InsideWidth) - 12)
    fr.Height = Application.Max(120, IIf(mp.Height > 0, mp.Height, Me.InsideHeight) - 30)
    fr.ZOrder 0

    ' 繝ｬ繧､繧｢繧ｦ繝茨ｼ・蛻暦ｼ・
    Dim items As Variant: items = PostureItems()
    Dim cols As Long: cols = POSTURE_COLS
    Dim colW As Single: colW = Application.Max(120, (fr.Width - 24) / cols)
    Dim rows As Long: rows = (UBound(items) - LBound(items) + 1 + cols - 1) \ cols

    Dim startY As Single: startY = 12
    Dim rowH As Single:   rowH = 28
    Dim labelW As Single: labelW = Application.Max(60, colW - 110)

    Dim i As Long, c As Long, r As Long, x As Single, y As Single
    For i = LBound(items) To UBound(items)
        c = (i - LBound(items)) \ rows
        r = (i - LBound(items)) Mod rows
        x = 12 + c * colW
        y = startY + r * rowH

        With fr.controls("lblPost_" & CStr(i))
            .Left = x
            .Top = y + 3
            .Width = labelW
            .Visible = True
        End With
        With fr.controls("cmbPost_" & CStr(i))
            .Left = x + labelW + 6
            .Top = y
            .Width = 100
            .Visible = True
        End With
    Next i

    '窶補・蛯呵・・菴咲ｽｮ縺ｨ繧ｵ繧､繧ｺ・・陦悟・ 竕・邏・0px・・
    Dim noteTop As Single: noteTop = startY + rows * rowH + 10
    With fr.controls("lblPost_Note")
        .Left = 12
        .Top = noteTop + 2
        .Width = 40
        .Visible = True
    End With
    With fr.controls("txtPost_Note")
        .Left = 12 + 40 + 6
        .Top = noteTop
        .Width = fr.Width - 24 - 46
        .Height = 90          ' 竊・縺縺・◆縺・陦檎嶌蠖・
        .Visible = True
    End With

    ' 繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ・亥ｙ閠・・荳狗ｫｯ縺ｾ縺ｧ繧偵き繝舌・・・
    fr.ScrollBars = fmScrollBarsNone
    fr.ScrollHeight = noteTop + 90 + 12
    If fr.Height < fr.ScrollHeight Then fr.ScrollBars = fmScrollBarsVertical
    
    ' Call UserForm_Resize  ' NOTE: 逶ｴ蜻ｼ縺ｳ縺吶ｋ縺ｨ mpPhys 縺ｮ鬮倥＆縺悟・險育ｮ励＆繧後悟・菴薙′遏ｭ縺・榊・逋ｺ貅舌↓縺ｪ繧九◆繧∫ｦ∵ｭ｢

    
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

Private Sub HookRomMirrorButtons_Once()
    If mRomMirrorHooks Is Nothing Then Set mRomMirrorHooks = New Collection
    If mRomMirrorHooks.count > 0 Then Exit Sub
    HookRomMirrorButtonsInContainer Me
End Sub

Private Sub HookRomMirrorButtonsInContainer(ByVal container As Object)
    On Error Resume Next

    Dim c As Object
    If TypeName(container) = "MultiPage" Then
        Dim pg As MSForms.page
        For Each pg In container.Pages
            HookRomMirrorButtonsInContainer pg
        Next
        Exit Sub
    End If

    For Each c In container.controls
        If TypeName(c) = "CommandButton" Then
            If CStr(c.tag) = "ROM_MIRROR" Then
                Dim h As clsRomMirrorBtnHook
                Set h = New clsRomMirrorBtnHook
                Set h.btn = c
                mRomMirrorHooks.Add h
            End If
        End If

        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            HookRomMirrorButtonsInContainer c
        End If
    Next
End Sub



'====================
' 繝倥Ν繝代・髢｢謨ｰ
'====================

Private Function RequiredOk() As Boolean
    On Error Resume Next
    RequiredOk = (Len(Trim$(Me.controls("txtPID").text)) > 0) _
        And (Len(Trim$(Me.controls("txtName").text)) > 0) _
        And IsNumeric(Me.controls("txtAge").text) _
        And (val(Me.controls("txtAge").text) >= 0)
End Function

Private Sub RefreshSaveEnabled()
    If Not btnSaveCtl Is Nothing Then btnSaveCtl.Enabled = RequiredOk()
End Sub

Private Sub txtPID_Change():  RefreshSaveEnabled: End Sub

Private Sub txtHdrName_Change()
     EnsureNameSuggestList

     Me.controls("txtName").text = Me.controls("frHeader").controls("txtHdrName").text
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



    Set host = Me.controls("frHeader")
    Set tb = host.controls("txtHdrName")



    ' 蛟呵｣懊Μ繧ｹ繝育｢ｺ菫・
    On Error Resume Next
        Dim i As Long
        Set lb = Nothing
           For i = Me.controls.count - 1 To 0 Step -1
           If Me.controls(i).name = "lstNameSuggest" Then
        Set lb = Me.controls(i)
        Exit For
    End If
Next i

    On Error GoTo 0

    If lb Is Nothing Then
        EnsureNameSuggestList
        Set lb = Me.controls("lstNameSuggest")
    End If


    key = Trim$(tb.text)
    keyN = NormalizeName(key)

    lb.Clear
    lb.Visible = False
    If Len(keyN) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(EVAL_SHEET_NAME)

    ' 豌丞錐蛻暦ｼ域里蟄倥Ο繧ｸ繝・け縺ｨ蜷後§謗｢縺玲婿・・
    cName = FindHeaderColLocal(ws, "Basic.Name")
    If cName = 0 Then cName = FindHeaderColLocal(ws, "豌丞錐")
    If cName = 0 Then cName = FindHeaderColLocal(ws, "Name")
    If cName = 0 Then Exit Sub

    ' ID蛻暦ｼ医≠繧後・菴ｵ險假ｼ・
    cID = FindHeaderColLocal(ws, "Basic.ID")
    If cID = 0 Then cID = FindHeaderColLocal(ws, "ID")
    If cID = 0 Then cID = FindHeaderColLocal(ws, "PID")


    lastRow = ws.Cells(ws.rows.count, cName).End(xlUp).row

    ' 2蛻励↓縺励※縲・蛻礼岼・・D・峨・髱櫁｡ｨ遉ｺ驕狗畑・郁｡ｨ遉ｺ譁・ｭ怜・縺ｫ菴ｵ險倥☆繧具ｼ・
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
           MsgBox "蜷悟ｧ灘酔蜷阪・蛟呵｣懊′隍・焚縺ゅｊ縺ｾ縺吶ょｿ・ｦ√↑繧迂D縺ｧ邨槭ｊ霎ｼ縺ｿ縺励※縺上□縺輔＞縲・, vbInformation
           mDupNameWarned = True
      End If
    End If



End Sub

Private Function FindHeaderColLocal(ws As Worksheet, headerText As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol
        If CStr(ws.Cells(1, c).value) = headerText Then
            FindHeaderColLocal = c
            Exit Function
        End If
    Next c
End Function



Private Sub txtAge_Change():  RefreshSaveEnabled: End Sub





'========================
' 繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ辟｡縺励・繝帙せ繝医ヵ繝ｬ繝ｼ繝
'========================
Private Function CreateScrollHost(pg As MSForms.page) As MSForms.Frame
    Dim host As MSForms.Frame
    Set host = pg.controls.Add("Forms.Frame.1")

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
' 繧ｷ繝ｼ繝茨ｼ・valData・・窶ｻfrmEval繝ｭ繝ｼ繧ｫ繝ｫ迚・
'========================
Private Function EnsureEvalData() As Worksheet
    Const sh As String = "EvalData"
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sh)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = sh
    End If

    Set EnsureEvalData = ws
End Function

Private Sub EnsureJapaneseHeaderRow(ws As Worksheet)
    If Application.WorksheetFunction.CountA(ws.rows(2)) = 0 Then
        Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        Dim c As Long
        For c = 1 To lastCol
            Select Case CStr(ws.Cells(1, c).value)
                Case "PatientID":        ws.Cells(2, c).value = "ID"
                Case "EvalDate":         ws.Cells(2, c).value = "隧穂ｾ｡譌･"
                Case "Basic.Name":       ws.Cells(2, c).value = "豌丞錐"
                Case "Basic.Age":        ws.Cells(2, c).value = "蟷ｴ鮨｢"
                Case "Basic.Gender":     ws.Cells(2, c).value = "諤ｧ蛻･"
                Case "Basic.PrimaryDx":  ws.Cells(2, c).value = "荳ｻ險ｺ譁ｭ"
                Case "Basic.OnsetDate":  ws.Cells(2, c).value = "逋ｺ逞・律"
                Case "Basic.Living":     ws.Cells(2, c).value = "逕滓ｴｻ迥ｶ豕・
                Case "Basic.CareLevel":  ws.Cells(2, c).value = "隕∽ｻ玖ｭｷ蠎ｦ"
                ' 蠢・ｦ√↓蠢懊§縺ｦ霑ｽ蜉
            End Select
        Next
    End If
End Sub

'========================
' 逶ｴ霑題｡梧爾邏｢・・ID・・
'========================
Private Function FindLastRowByPID(ByVal pid As String, ByVal ws As Worksheet) As Long
    Dim colPID As Variant, colDate As Variant, colTS As Variant
    colPID = Application.Match("PatientID", ws.rows(1), 0)
    colDate = Application.Match("EvalDate", ws.rows(1), 0)
    colTS = Application.Match("Timestamp", ws.rows(1), 0)

    If IsError(colPID) Or IsError(colDate) Then Exit Function

    Dim last As Long: last = ws.Cells(ws.rows.count, 1).End(xlUp).row

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
' RLA 繝ｬ繝吶Ν蜿門ｾ・
'========================
Private Function GetRLAGroupLevel(ByVal grp As String) As String
    Dim c As MSForms.Control, p As MSForms.page, fr As MSForms.Control, ob As MSForms.Control
    For Each c In hostWalk.controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each fr In p.controls
                    If TypeName(fr) = "Frame" Then
                        For Each ob In fr.controls
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
' 蜿朱寔
'========================
Private Function CollectFormData() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As MSForms.Control, oc As MSForms.Control, ic As MSForms.Control

    Dim j As Long
    For Each c In Me.controls
        Select Case TypeName(c)
            Case "MultiPage"
                For j = 0 To c.Pages.count - 1
                    Dim p As MSForms.page: Set p = c.Pages(j)
                    Dim co As MSForms.Control
                    For Each co In p.controls
                        CollectOne d, co
                        If TypeName(co) = "Frame" Then
                            For Each ic In co.controls: CollectOne d, ic: Next
                        ElseIf TypeName(co) = "MultiPage" Then
                            Dim p2 As MSForms.page, fr As MSForms.Control, it As MSForms.Control
                            For Each p2 In co.Pages
                                For Each fr In p2.controls
                                    CollectOne d, fr
                                    If TypeName(fr) = "Frame" Then
                                        For Each it In fr.controls: CollectOne d, it: Next
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

    Dim frames As Collection: Set frames = FindAllFramesByCaptionPart("RLA豁ｩ陦悟捉譛・)
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
        Case "TextBox", "ComboBox": d(ctl.tag) = ctl.text
        Case "CheckBox"
            If ctl.tag <> "AssistiveGroup" And ctl.tag <> "RiskGroup" Then
                d(ctl.tag) = IIf(ctl.value, "譛・, "辟｡")
            End If
    End Select
End Sub

Private Function AggregateChecks(ByVal groupTag As String) As String
    Dim picks As String, c As MSForms.Control, p As MSForms.page, fr As MSForms.Control, cc As MSForms.Control
    For Each c In Me.controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each fr In p.controls
                    If TypeName(fr) = "Frame" Then
                        For Each cc In fr.controls
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
    For Each c In f.controls
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
    Dim c As MSForms.Control, p As MSForms.page, oc As MSForms.Control
    For Each c In Me.controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                For Each oc In p.controls
                    If TypeName(oc) = "Frame" Then
                        If InStr(1, oc.caption, part, vbTextCompare) > 0 Then col.Add oc
                    ElseIf TypeName(oc) = "MultiPage" Then
                        Dim p2 As MSForms.page, oc2 As MSForms.Control
                        For Each p2 In oc.Pages
                            For Each oc2 In p2.controls
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
' 蜈･蜉帙メ繧ｧ繝・け
'========================
Private Function CheckRange(frm As Object, ByVal nm As String, ByVal lo As Double, ByVal hi As Double, ByVal message As String, ByRef sb As String) As Boolean
    If Not FnHasControl(nm) Then CheckRange = True: Exit Function
    Dim t As String: t = Trim$(frm.controls(nm).text & "")
    If t = "" Then CheckRange = True: Exit Function
    If Not IsNumeric(t) Then sb = sb & "繝ｻ" & message & vbCrLf: CheckRange = False: Exit Function
    Dim v As Double: v = CDbl(t)
    If v < lo Or v > hi Then sb = sb & "繝ｻ" & message & vbCrLf: CheckRange = False: Exit Function
    CheckRange = True
End Function

Private Function ValidateForm(ByRef errmsg As String) As Boolean
    Dim ok As Boolean: ok = True
    Dim sb As String: sb = ""

    If FnHasControl("txtName") Then If Trim$(Me.controls("txtName").text) = "" Then ok = False: sb = sb & "繝ｻ豌丞錐繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・ & vbCrLf
    If FnHasControl("txtAge") Then
        If Trim$(Me.controls("txtAge").text) = "" Or Not IsNumeric(Me.controls("txtAge").text) Then
            ok = False: sb = sb & "繝ｻ蟷ｴ鮨｢繧呈焚蛟､縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・ & vbCrLf
        ElseIf val(Me.controls("txtAge").text) < 0 Or val(Me.controls("txtAge").text) > 120 Then
            ok = False: sb = sb & "繝ｻ蟷ｴ鮨｢縺ｯ0・・20縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・ & vbCrLf
        End If
    End If

    ' 隧穂ｾ｡譌･繝√ぉ繝・け
    If FnHasControl("txtEDate") Then
        If Not IsDate(Me.controls("txtEDate").text) Then
            ok = False: sb = sb & "繝ｻ隧穂ｾ｡譌･繧呈ｭ｣縺励＞譌･莉假ｼ・yyy/mm/dd 遲会ｼ峨〒蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・ & vbCrLf
        End If
    End If

    ok = ok And CheckRange(Me, "txtTenMWalk", 0, 300, "10m豁ｩ陦鯉ｼ育ｧ抵ｼ峨・0・・00縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)
    ok = ok And CheckRange(Me, "txtTUG", 0, 300, "TUG・育ｧ抵ｼ峨・0・・00縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)
    ok = ok And CheckRange(Me, "txtFiveSTS", 0, 300, "5蝗樒ｫ九■荳翫′繧奇ｼ育ｧ抵ｼ峨・0・・00縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)
    ok = ok And CheckRange(Me, "txtSemi", 0, 300, "繧ｻ繝溘ち繝ｳ繝・Β・育ｧ抵ｼ峨・0・・00縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)
    ok = ok And CheckRange(Me, "txtGripR", 0, 120, "謠｡蜉・蜿ｳ・・g・峨・0・・20縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)
    ok = ok And CheckRange(Me, "txtGripL", 0, 120, "謠｡蜉・蟾ｦ・・g・峨・0・・20縺ｧ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, sb)

    errmsg = IIf(ok, "", sb)
    ValidateForm = ok
End Function

Private Sub btnSaveCtl_Click()
    Call SyncAgeFromBirth
    Me.controls("txtName").text = Me.controls("txtHdrName").text
    SaveEvaluation_Append_From Me
End Sub



'=== frmEval・壼燕蝗櫁ｪｭ霎ｼ繝懊ち繝ｳ 螳悟・雋ｼ繧頑崛縺・============================
Private Sub btnLoadPrevCtl_Click()

Me.controls("txtName").text = Me.controls("txtHdrName").text



    Call modEvalIOEntry.LoadEvaluation_ByName_From(Me)


Me.Repaint

    Exit Sub

End Sub

Public Sub HandleHdrLoadPrevClick()
    
    Call btnLoadPrevCtl_Click
End Sub

Private Sub cmdHdrLoadPrev_Click()
    Call btnLoadPrevCtl_Click
End Sub


' 荳九°繧蛾■縺｣縺ｦ豌丞錐荳閾ｴ縺ｮ縲梧怙譁ｰ?譛螟ｧ5莉ｶ縲阪ｒ髮・ａ縲・
' 莉ｶ謨ｰ=1縺ｪ繧峨◎繧後ｒ霑斐＠縲・莉･荳翫↑繧臥分蜿ｷ驕ｸ謚槭・InputBox繧貞・縺・
Private Function FindRowByNameWithPickLocal(ws As Worksheet, nameText As String, Optional maxCount As Long = 5) As Long
    Dim colName As Long, colDate As Long
    colName = modEvalIOEntry.FindColByHeaderExact(ws, "豌丞錐")
    If colName = 0 Then colName = modEvalIOEntry.FindColByHeaderExact(ws, "蛻ｩ逕ｨ閠・錐")
    If colName = 0 Then colName = modEvalIOEntry.FindColByHeaderExact(ws, "蜷榊燕")
    If colName = 0 Then Exit Function

    colDate = modEvalIOEntry.FindColByHeaderExact(ws, "隧穂ｾ｡譌･")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "險倬鹸譌･")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "譖ｴ譁ｰ譌･")
    If colDate = 0 Then colDate = modEvalIOEntry.FindColByHeaderExact(ws, "菴懈・譌･")

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, colName).End(xlUp).row
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
                disp = disp & i & ") ・域律莉倥↑縺暦ｼ・ Row:" & rows(i) & vbCrLf
            End If
        Else
            disp = disp & i & ") Row:" & rows(i) & vbCrLf
        End If
    Next i

    Dim idx As Variant
    idx = Application.InputBox( _
            prompt:="蜷悟ｧ灘酔蜷阪・逶ｴ霑・ & cnt & "莉ｶ・域怙譁ｰ竊貞商縺・ｼ・" & vbCrLf & disp & vbCrLf & _
                    "逡ｪ蜿ｷ繧貞・蜉幢ｼ・-" & cnt & "縲√く繝｣繝ｳ繧ｻ繝ｫ=荳ｭ豁｢・・, _
            Type:=1)
    If idx = False Then Exit Function
    If idx >= 1 And idx <= cnt Then
        FindRowByNameWithPickLocal = rows(CLng(idx))
    End If
End Function



'=== 蛟呵｣應ｸ隕ｧ繧定｡ｨ遉ｺ縺励※逡ｪ蜿ｷ蜈･蜉帙〒1縺､驕ｸ繧薙〒繧ゅｉ縺・======================
Public Function PickCandidateRowByNameLocal(ByVal ws As Worksheet, _
                                             ByVal look As Object, _
                                             ByVal candidates As Variant, _
                                             ByVal pname As String) As Long
    
    Debug.Print "[ENTER] PickCandidateRowByNameLocal", Timer

    
    If IsEmpty(candidates) Then Exit Function
    If Not IsArray(candidates) Then Exit Function

    Dim pidCol As Long, ageCol As Long, dtCol As Long
    pidCol = RCol(ws, look, "Basic.ID", "ID", "蛟倶ｺｺID")
ageCol = RCol(ws, look, "Basic.Age", "蟷ｴ鮨｢")
dtCol = RCol(ws, look, "Basic.EvalDate", "隧穂ｾ｡譌･", "險倬鹸譌･", "譖ｴ譁ｰ譌･", "菴懈・譌･")

    If dtCol = 0 Then dtCol = ResolveColumnLocal(look, "隧穂ｾ｡譌･")
    If dtCol = 0 Then dtCol = ResolveColumnLocal(look, "EvalDate")

    Dim lb As Long, ub As Long, cnt As Long
    lb = LBound(candidates): ub = UBound(candidates): cnt = ub - lb + 1
    Debug.Print "[CANDS] name=", pname, " cnt=", cnt, " range=", lb & "-" & ub



    Dim i As Long, r As Long, msg As String, disp As Long
    msg = "蜷悟ｧ灘酔蜷阪′隕九▽縺九ｊ縺ｾ縺励◆縲りｪｭ縺ｿ霎ｼ繧繝・・繧ｿ繧帝∈謚槭＠縺ｦ縺上□縺輔＞・育分蜿ｷ繧貞・蜉幢ｼ峨・ & vbCrLf & vbCrLf
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
            msg = msg & "(隧穂ｾ｡譌･縺ｪ縺・"
        End If
        If ageCol > 0 Then msg = msg & " | 蟷ｴ鮨｢:" & Trim$(CStr(ws.Cells(r, ageCol).value))
        If pidCol > 0 Then msg = msg & " | ID:" & Trim$(CStr(ws.Cells(r, pidCol).value))
        msg = msg & vbCrLf
    Next i

   ' 蜈･蜉帛叙蠕暦ｼ域焚蛟､縺ｮ縺ｿ・・
Dim sel As Variant
sel = Application.InputBox(msg, "蛟呵｣憺∈謚・- " & pname, Type:=1)

' Cancel / 繧ｨ繝ｩ繝ｼ / 髱樊焚蛟､ / 遨ｺ 繧貞ｼｾ縺擾ｼ育洒邨｡縺ｧ蛻､螳夲ｼ・
If VarType(sel) = vbBoolean Then Exit Function   ' Cancel
If IsError(sel) Then Exit Function               ' 縺ｾ繧後↓ CVErr
If Not IsNumeric(sel) Then Exit Function
If Len(CStr(sel)) = 0 Then Exit Function


    Dim n As Long: n = CLng(sel)
    If n < 1 Or n > cnt Then Exit Function

    PickCandidateRowByNameLocal = candidates(lb + n - 1)
End Function


Private Function NzTxt(tb As MSForms.TextBox) As String
    On Error Resume Next
    NzTxt = ""
    If Not tb Is Nothing Then NzTxt = Trim$(tb.text)
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



'=== 縺薙％縺九ｉ陬懷勧髢｢謨ｰ鄒､・・rmEval 繝ｭ繝ｼ繧ｫ繝ｫ・・============================
Private Function FxGetText(ByVal ctrlName As String) As String
    On Error Resume Next
    FxGetText = Trim$(Me.controls(ctrlName).text)
End Function

Private Sub FxSetText(ByVal ctrlName As String, ByVal value As String)
    On Error Resume Next
    Me.controls(ctrlName).text = value
End Sub

Private Function GetOrCreateEvalSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        ws.name = "EvalData"
    End If
    Set GetOrCreateEvalSheet = ws
End Function

Private Function BuildHeaderLookupLocal(ByVal ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 'TextCompare

    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
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

    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
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

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, idCol).End(xlUp).row
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

'=== 豌丞錐縺ｧ譛譁ｰ陦後ｒ謗｢縺呻ｼ・D迚医・螳悟・莠呈鋤繝ｭ繧ｸ繝・け・・=====================
Private Function FindLastRowByNameLocal(ByVal ws As Worksheet, _
                                        ByVal look As Object, _
                                        ByVal pname As String) As Long
    ' 隕句・縺励・蛻嶺ｽ咲ｽｮ・育┌縺代ｌ縺ｰ菴懈・縺励※謠・∴繧具ｼ唔D迚医→蜷後§譁ｹ驥晢ｼ・
    Dim nameCol As Long: nameCol = ResolveColOrCreate(ws, look, "Basic.Name", "豌丞錐", "Name")
    Dim tsCol As Long:   tsCol = ResolveColOrCreate(ws, look, "Timestamp")
    Dim dtCol As Long:   dtCol = ResolveColOrCreate(ws, look, "Basic.EvalDate", "隧穂ｾ｡譌･", "EvalDate")

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, nameCol).End(xlUp).row
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
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row
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
    s = Replace(s, "縲", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeKeyLocal = LCase$(s)
End Function
'=== 陬懷勧髢｢謨ｰ鄒､ 縺薙％縺ｾ縺ｧ ===============================================

Private Sub btnCloseCtl_Click()

    RequestQuitExcelAskAndCloseForm
    Unload Me
End Sub


Private Sub SetImeRecursive(container As Object)
    Dim ctl As Control
    For Each ctl In container.controls
        If TypeName(ctl) = "TextBox" Then ctl.IMEMode = fmIMEModeHiragana

        Select Case TypeName(ctl)
            Case "Frame", "UserForm"
                SetImeRecursive ctl

            Case "MultiPage"
                Dim p As MSForms.page
                For Each p In ctl.Pages          ' 笘・Pages 繧貞・謖吶☆繧九・縺碁㍾隕・
                    SetImeRecursive p
                Next
        End Select
    Next
End Sub

Private Sub FixRestNRS_Once()
    Dim L As Control, c As Control
    On Error Resume Next
    Set c = Me.controls("cmbNRS_Move")                  ' 蜍穂ｽ懈凾NRS縺ｮ繧ｳ繝ｳ繝・
    If c Is Nothing Then Exit Sub

    ' 霑代＞鬮倥＆縺ｫ縺ゅｋ繝ｩ繝吶Ν繧呈爾縺呻ｼ郁ｦ九▽縺九ｉ縺ｪ縺代ｌ縺ｰ譁ｰ隕上↓菴懊ｋ・・
    For Each L In Me.controls
        If TypeName(L) = "Label" Then
            If L.caption = "螳蛾撕譎・RS" Or (Abs(L.Top - c.Top) <= 20 And L.Left < c.Left) Then
                Exit For
            End If
        End If
    Next
    If L Is Nothing Then
        Set L = c.parent.controls.Add("Forms.Label.1", "lblNRS_Rest", True)
    End If

    ' 陦ｨ遉ｺ繝ｻ菴咲ｽｮ繝ｻ繧ｵ繧､繧ｺ繧堤｢ｺ螳・
    L.caption = "螳蛾撕譎・RS"
    L.Visible = True
    L.WordWrap = False
    L.Width = 72
    L.Top = c.Top
    L.Left = c.Left - (L.Width + 6)
    L.ZOrder 0
End Sub










'=== 霑ｽ蜉・哺ultiPage縺ｮ譌｢螳壹・繝ｼ繧ｸ繧呈祉髯､・・Page*" 繧貞・驛ｨ蜑企勁・・===
Private Sub CleanDefaultPages(mp As MSForms.MultiPage)
    On Error Resume Next
    Dim i As Long
    ' 繧ｭ繝｣繝励す繝ｧ繝ｳ縺・"Page" 縺ｧ蟋九∪繧九・繝ｼ繧ｸ繧貞ｾ後ｍ縺九ｉ蜑企勁
    For i = mp.Pages.count - 1 To 0 Step -1
        If Left$(mp.Pages(i).caption, 4) = "Page" Then
            mp.Pages.Remove i
        End If
    Next
End Sub


' 謖・ｮ壹く繝｣繝励す繝ｧ繝ｳ縺ｮ繝懊ち繝ｳ繧貞・蟶ｰ縺ｧ謗｢縺・
Private Function FindButtonByCaption(container As Object, ByVal cap As String) As MSForms.CommandButton
    Dim c As Object, hit As MSForms.CommandButton
    For Each c In container.controls
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

    ' 縺ｾ縺壹く繝｣繝励す繝ｧ繝ｳ驛ｨ蛻・ｸ閾ｴ縺ｧ
    '=== 縺ｾ縺壹く繝｣繝励す繝ｧ繝ｳ荳閾ｴ・亥ｮ壽焚縺ｧ蜴ｳ蟇・↓・・===
Const CAP_SAVE As String = "繧ｷ繝ｼ繝医∈菫晏ｭ・          ' 竊・cap=[ ] 縺ｮ荳ｭ霄ｫ縺縺・
Const CAP_LOAD As String = "蜑榊屓縺ｮ蛟､繧定ｪｭ縺ｿ霎ｼ繧"    ' 竊・cap=[ ] 縺ｮ荳ｭ霄ｫ縺縺・

Debug.Print "[TEST] 菫晏ｭ詫ike:", TypeName(FindButtonByCaptionLike(Me, "菫晏ｭ・))
Debug.Print "[TEST] 隱ｭ縺ｿ霎ｼlike:", TypeName(FindButtonByCaptionLike(Me, "隱ｭ縺ｿ霎ｼ"))


Set btnHdrSave = FindButtonByCaptionLike(Me, CAP_SAVE)
Set btnHdrLoadPrev = FindButtonByCaptionLike(Me, CAP_LOAD)


    If btnHdrLoadPrev Is Nothing Then Set btnHdrLoadPrev = FindButtonByCaptionLike(Me, "隱ｭ")

    ' 隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ譛荳頑ｮｵ縺ｮ蜿ｳ蛛ｴ・偵▽繧呈治逕ｨ
    If (btnHdrSave Is Nothing) Or (btnHdrLoadPrev Is Nothing) Then
        FindHeaderButtonsByPosition Me, btnHdrSave, btnHdrLoadPrev
    End If

    Debug.Print "[HOOK] save=", IIf(btnHdrSave Is Nothing, "NG", "OK"), _
                "    load=", IIf(btnHdrLoadPrev Is Nothing, "NG", "OK")
End Sub





'=== 繝懊ち繝ｳ繧偵く繝｣繝励す繝ｧ繝ｳ驛ｨ蛻・ｸ閾ｴ縺ｧ謗｢縺呻ｼ・rame/MultiPage/Page 繧貞・蟶ｰ・・==
Public Function FindButtonByCaptionLike(container As Object, _
                                        ByVal needle As String) As Object
SafeExit:
        mAgeBusy = False

    Dim c  As Object
    Dim pg As MSForms.page
    Dim hit As Object

    ' MultiPage 縺ｯ Pages 邨檎罰縺ｧ貎懊ｋ・医％縺薙′驥崎ｦ・ｼ・
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

    ' 縺昴ｌ莉･螟悶・ Controls 繧定ｵｰ譟ｻ
    For Each c In container.controls
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


End Function


'=== 菴咲ｽｮ縺ｧ繝倥ャ繝陦後・繝懊ち繝ｳ繧呈鏡縺・ｼ域怙荳頑ｮｵ縺ｮ蜿ｳ蛛ｴ 2 蛟九ｒ謗｡逕ｨ・・==
Private Sub GatherButtons(container As Object, ByRef arr As Collection)
    Dim c  As Object
    Dim pg As MSForms.page

    ' MultiPage 縺ｯ Pages 繧貞・蟶ｰ
    If TypeName(container) = "MultiPage" Then
        For Each pg In container.Pages
            GatherButtons pg, arr
        Next
        Exit Sub
    End If

    ' 縺昴ｌ莉･螟悶・ Controls
    For Each c In container.controls
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
    If arr.count = 0 Then Exit Sub

    Dim minTop As Single: minTop = 1E+20
    For i = 1 To arr.count
        If arr(i).Top < minTop Then minTop = arr(i).Top
    Next

    Dim cand As New Collection
    For i = 1 To arr.count
        If arr(i).Top <= minTop + 10 Then cand.Add arr(i)   '譛荳頑ｮｵﾂｱ10px
    Next
    ' 繧ｭ繝ｼ繝ｯ繝ｼ繝峨〒蜆ｪ蜈亥愛螳・
    For i = 1 To cand.count
        Dim cap As String: cap = Replace(CStr(cand(i).caption), vbCrLf, "")
        If InStr(cap, "菫晏ｭ・) > 0 Then Set bSave = cand(i)
        If InStr(cap, "隱ｭ") > 0 Or InStr(cap, "霎ｼ") > 0 Then Set bLoad = cand(i)
    Next
    ' 縺ｾ縺遨ｺ縺・※縺・◆繧牙承蛛ｴ・偵▽繧貞牡繧雁ｽ薙※
    Dim right1 As MSForms.CommandButton, right2 As MSForms.CommandButton
    For i = 1 To cand.count
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



'=== 荳隕ｧ繧貞・縺吝・蜿｣・亥・髢具ｼ・===
Public Sub Debug_ListButtons()
    DumpButtonsProc Me
End Sub

'=== 繝懊ち繝ｳ蛻玲嫌・・ultiPage蟇ｾ蠢懃沿・・===
Private Sub DumpButtonsProc(container As Object)
    Dim c As Control
    Dim pg As MSForms.page

    If TypeName(container) = "MultiPage" Then
        'MultiPage 縺ｯ Pages 驟堺ｸ九ｒ蝗槭☆
        For Each pg In container.Pages
            For Each c In pg.controls
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
        '騾壼ｸｸ縺ｮ繧ｳ繝ｳ繝・リ・・serForm / Frame / Page 縺ｪ縺ｩ・・
        For Each c In container.controls
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



' ========= 縺薙％縺九ｉ雋ｼ繧具ｼ・mdLoadPrev_Click 蜈ｨ菴難ｼ・========
Private Sub cmdLoadPrev_Click()
Debug.Print "[ENTER] cmdLoadPrev_Click", Timer

    Call modEvalIOEntry.LoadEvaluation_ByName_From(Me)

End Sub





Private Sub UserForm_Activate()
   
    Static done As Boolean
    Dim scrH As Single
    Dim h As Single
    
    Application.WindowState = xlMaximized
    
    If done Then Exit Sub
    done = True
    
    Call Align_BIHomeEnv_Once
    
    
    Me.controls("txtHdrName").SetFocus
    

End Sub




'==== UserForm 繧ｳ繝ｼ繝峨Δ繧ｸ繝･繝ｼ繝ｫ・・raDynPainBox 縺瑚ｼ峨▲縺ｦ縺・ｋ繝輔か繝ｼ繝・・===

Private Sub UserForm_Initialize()

Me.StartUpPosition = 0
Me.Width = 1072
Me.Height = 632.15
Me.Left = Application.Left + (Application.Width - Me.Width) / 2: Me.Top = Application.Top + (Application.Height - Me.Height) / 2



    Dim scrH As Single
    Dim h As Single
    Dim mp2 As Object
    Dim mp2Parent As Object
    Dim mpPhysObj As Object
    Dim pgPhys0 As Object
    Dim frPhys8 As Object
    Dim mpROMObj As Object
    Dim mp1 As Object
    Dim pg1 As Object
    Dim fr32 As Object
    Dim btnLoadPrev As Object
    scrH = Application.UsableHeight
    If scrH < 500 Then
        h = 530
    Else
        h = 690
    End If

    Me.Height = h
    DoEvents

    Call LegacyInit
    EnsureMpPhysChangeHook_Once
#If APP_DEBUG Then
    Debug.Print "[PostInit] CtlCount=" & Me.controls.count
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
    
    TidyPainBoxes        ' 竊・蜿ｳ蛻・隱伜屏繝ｻ霆ｽ貂帛屏蟄・縺ｮ諱剃ｹ・・鄂ｮ
    TidyPainCourse       ' 竊・笘・ｿｽ蜉・夂ｵ碁℃繝ｻ譎る俣縺ｮ螟牙喧縺ｮ諱剃ｹ・・鄂ｮ
    Me.WidenAndTidyPainCourse
 
   
    

    Me.TidyPainUI_Once

If Not mPainTidyBusy Then
    'TidyPainUI_Once
    Me.Height = h
    DoEvents
    ClearPainUI Me   ' 竊・襍ｷ蜍墓凾縺ｯ遨ｺ縺ｧ髢句ｧ具ｼ郁ｪｭ縺ｿ霎ｼ縺ｿ縺ｯ謇句虚縺ｧ・・
End If

        BuildWalkUI_All

    Dim cogRoot As MSForms.Frame
    Set cogRoot = GetCogRootFrame()
    If Not cogRoot Is Nothing Then cogRoot.caption = ""
    BuildCogMentalUI_Simple
    BuildCog_CognitionCore      '竊・隱咲衍6鬆・岼繧堤函謌・
    BuildCog_DementiaBlock      '竊・隱咲衍逞・・遞ｮ鬘橸ｼ句ｙ閠・ｒ逕滓・
    BuildCog_BPSD
    BuildCog_MentalBlock
    BuildDailyLogTab
    
        ApplyDailyLogImeSettings
    
    Set mDailyList = New clsDailyLogList
    'Set mDailyList.lb = Me.Controls("lstDailyLogList")


    
    BuildDailyLog_ExtractButton Me
    BuildDailyLog_SaveButton Me
    Me.controls("txtEDate").value = Date
    Dim txtDailyDate As Object
    Set txtDailyDate = DailyLogCtl("txtDailyDate")
    If Not txtDailyDate Is Nothing Then txtDailyDate.value = Date

    If Not mPlacedGlobalSave Then
    PlaceGlobalSaveButton_Once
    mPlacedGlobalSave = True
    End If
    


    
    On Error Resume Next
    Me.controls("btnSaveCtl").Visible = False
    On Error GoTo 0
    
    'Me.Height = Application.UsableHeight - 40
    
    




If mBaseLayoutDone Then
    'Apply_AlignRoot_All
End If


'--- Fix: 蟄信ultiPage隕句・繧悟ｯｾ遲厄ｼ・025-12-13 OK繧ｹ繝翫ャ繝励す繝ｧ繝・ヨ蝗ｺ螳夲ｼ・


On Error Resume Next
Set mp2 = SafeGetControl(Me, "MultiPage2")
If Not mp2 Is Nothing Then
    Set mp2Parent = mp2.parent
    If Not mp2Parent Is Nothing Then
        mp2Parent.Height = mp2.Height
    End If
End If
On Error GoTo 0

Dim frame12 As Object
Set frame12 = SafeGetControl(Me, "Frame12")
If Not frame12 Is Nothing Then frame12.Height = 508.1
On Error Resume Next
Set mpPhysObj = SafeGetControl(Me, "mpPhys")
If Not mpPhysObj Is Nothing Then
    Set pgPhys0 = mpPhysObj.Pages(0)
    If Not pgPhys0 Is Nothing Then
       Set frPhys8 = SafeGetControl(pgPhys0, "Frame8")
        If Not frPhys8 Is Nothing Then
            Set mpROMObj = frPhys8.controls("mpROM")
            frPhys8.Height = mpPhysObj.Height
            If Not mpROMObj Is Nothing Then
                mpROMObj.Height = frPhys8.InsideHeight - mpROMObj.Top
            End If
        End If
    End If
End If
On Error GoTo 0


    Call BuildEvalShell_Once
    
    Call CreateHeaderButtons_Once

    Tidy_DailyLog_Once

    
    

    Fix_Page8_DailyLog_Once
    Fix_Page6_Walk_FrameScroll_Once
    ApplyScroll_MP1_Page3_7_Once
    
    Call Preview_NameToHeader
    Me.controls("txtName").Visible = False
    Me.controls("txtPID").Visible = False

    
    AddHeaderArchiveDeleteButton

      RearrangeHeaderTopAreaLayout
      
AddPrintButton_TestEval

Call Ensure_MonthlyDraftBox_UnderFraDailyLog

Set mHdrNameSink = New cHdrNameSink
mHdrNameSink.Hook Me.controls("frHeader").controls("txtHdrName")


Call Align_LoadPrevButton_NextToHdrKana(Me)
Call Ensure_LoadPrevButton_Once(Me)
Call HookRomMirrorButtons_Once

'--- hook header "LoadPrev" button (MUST be after Ensure_LoadPrevButton_Once) ---
Set mHdrLoadPrevHook = New clsHdrBtnHook
Set mHdrLoadPrevHook.btn = Me.controls("frHeader").controls("cmdHdrLoadPrev")
mHdrLoadPrevHook.tag = "LoadPrev"
Set mHdrLoadPrevHook.owner = Me
DoEvents

    
 If Not mBasicInfoTidyDone Then
    mBasicInfoTidyDone = True
    Call TidyBasicInfo_TwoColumns
 End If
    
 EnsureBasicInfoEnterFixedRouteReady
 
  Call Fix_InnerScrollBars
 
End Sub

Private Sub ApplyDailyLogImeSettings()
    Dim dailyFra As Object

    Set dailyFra = GetDailyLogFrame()
    If dailyFra Is Nothing Then Set dailyFra = SafeGetControl(Me, "fraDailyLog")
    If dailyFra Is Nothing Then Exit Sub

    SetControlImeHiragana dailyFra, "txtDailyStaff"
    SetControlImeHiragana dailyFra, "txtDailyTraining"
    SetControlImeHiragana dailyFra, "txtDailyReaction"
    SetControlImeHiragana dailyFra, "txtDailyAbnormal"
    SetControlImeHiragana dailyFra, "txtDailyPlan"
End Sub

Private Sub SetControlImeHiragana(ByVal owner As Object, ByVal controlName As String)
    Dim tb As Object

    Set tb = SafeGetControl(owner, controlName)
    If tb Is Nothing Then Exit Sub

    tb.IMEMode = fmIMEModeHiragana
End Sub


Private Sub Fix_InnerScrollBars()

On Error Resume Next

Me.controls("MultiPage1").Pages("Page1").controls("Frame1").ScrollBars = fmScrollBarsNone
Me.controls("MultiPage1").Pages("Page1").controls("Frame1").KeepScrollBarsVisible = fmScrollBarsNone

Me.controls("MultiPage1").Pages("Page3").controls("Frame8").ScrollBars = fmScrollBarsNone
Me.controls("MultiPage1").Pages("Page3").controls("Frame8").KeepScrollBarsVisible = fmScrollBarsNone

Me.controls("MultiPage1").Pages("Page3").controls("Frame9").ScrollBars = fmScrollBarsNone
Me.controls("MultiPage1").Pages("Page3").controls("Frame9").KeepScrollBarsVisible = fmScrollBarsNone

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
'=== frmEval 縺ｮ繧ｳ繝ｼ繝峨Δ繧ｸ繝･繝ｼ繝ｫ縺ｫ雋ｼ繧具ｼ医％縺薙∪縺ｧ・・===





Private Sub LegacyInit()



    Set mMPHooks = New Collection
    Set mTxtHooks = New Collection
    Set mRomMirrorHooks = New Collection

SetupLayout

    Me.caption = "隧穂ｾ｡繝輔か繝ｼ繝"
Me.ScrollBars = fmScrollBarsNone

' 逕ｻ髱｢繧ｵ繧､繧ｺ縺ｫ蜷医ｏ縺帙※荳企剞繧偵°縺代ｋ・域怙蟆丞､画峩・・
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
Me.StartUpPosition = 1      ' CenterOwner・井ｻｻ諢擾ｼ壻ｸｭ螟ｮ縺ｫ蜃ｺ縺励◆縺・ｴ蜷茨ｼ・



    ' 繝ｫ繝ｼ繝・MultiPage
    Set mp = Me.controls.Add("Forms.MultiPage.1")
    With mp
        .Left = 6: .Top = 6
        .Width = Me.InsideWidth - 12
        .Height = Me.InsideHeight - 60
        .Style = fmTabStyleTabs
    End With
    ' 螳牙・遲厄ｼ壽怙菴・繝壹・繧ｸ繧剃ｿ晁ｨｼ・育腸蠅・ｷｮ蟇ｾ遲厄ｼ・
    If mp.Pages.count = 0 Then mp.Pages.Add: mp.Pages.Add

    ' 蜈磯ｭ2譫壹・譌｢蟄・
mp.Pages(0).caption = "蝓ｺ譛ｬ諠・ｱ"
mp.Pages(1).caption = "蟋ｿ蜍｢隧穂ｾ｡"

' 竊舌％縺薙〒縲瑚ｺｫ菴捺ｩ溯・隧穂ｾ｡縲阪ｒ蜈医↓霑ｽ蜉
Dim pgPhys  As MSForms.page: Set pgPhys = mp.Pages.Add: pgPhys.caption = "霄ｫ菴捺ｩ溯・隧穂ｾ｡"

' 谿九ｊ縺ｮ隕ｪ繧ｿ繝・
Dim pgMove  As MSForms.page: Set pgMove = mp.Pages.Add: pgMove.caption = "譌･蟶ｸ逕滓ｴｻ蜍穂ｽ・
Dim pgTests As MSForms.page: Set pgTests = mp.Pages.Add: pgTests.caption = "繝・せ繝医・隧穂ｾ｡"
Dim pgWalk  As MSForms.page: Set pgWalk = mp.Pages.Add: pgWalk.caption = "豁ｩ陦瑚ｩ穂ｾ｡"
Dim pgCog   As MSForms.page: Set pgCog = mp.Pages.Add: pgCog.caption = "隱咲衍繝ｻ邊ｾ逾・

' --- 繝帙せ繝亥､画焚縺ｮ螳｣險・医％縺薙ｒ霑ｽ蜉・・--
' --- 繝帙せ繝亥､画焚縺ｮ螳｣險 ---
Dim hostBasic As MSForms.Frame
Dim hostPost  As MSForms.Frame   ' 竊・譌ｧ hostBody 繧剃ｽｿ縺｣縺ｦ縺・◆繧臥ｽｮ謠・
Dim hostPhys  As MSForms.Frame
Dim hostMove  As MSForms.Frame
Dim hostTests As MSForms.Frame
Dim hostWalk  As MSForms.Frame
Dim hostCog   As MSForms.Frame


' --- 縺薙％縺九ｉ譌｢蟄倥・繝帙せ繝育函謌・---
Set hostBasic = CreateScrollHost(mp.Pages(0))
Set hostPost = CreateScrollHost(mp.Pages(1))
Set hostPhys = CreateScrollHost(pgPhys)
Set hostMove = CreateScrollHost(pgMove)
Set hostTests = CreateScrollHost(pgTests)
Set hostWalk = CreateScrollHost(pgWalk)
Set hostCog = CreateScrollHost(pgCog)



    
 Call modPhysEval.EnsurePhysicalFunctionTabs_Under(Me, mp)
    
    
    
    
' --- 霑ｷ蟄舌・ mpADL 繧貞・豸亥悉・井ｿ晞匱・・---
Trace "RemoveAllMpADL()", "Initialize"
RemoveAllMpADL
Trace "RemoveAllMpADL() done", "Initialize"


' --- mpADL 縺ｮ蝎ｨ繧堤畑諢擾ｼ医・繝ｼ繧ｸ0縺ｨ1繧剃ｽ懊ｋ・・---
Trace "EnsureBI_IADL() call", "Initialize"
Dim mpADL As MSForms.MultiPage
Set mpADL = EnsureBI_IADL()
Trace "EnsureBI_IADL() returned; pages=" & mpADL.Pages.count, "Initialize"


' --- 襍ｷ螻・虚菴懊・繝ｼ繧ｸ縺ｮUI繧剃ｽ懊ｋ ---
Trace "BuildKyoOnADL Page(2) start", "Initialize"
BuildKyoOnADL mpADL.Pages(2)
Trace "BuildKyoOnADL Page(2) end", "Initialize"



' 蛻晄悄陦ｨ遉ｺ縺ｯBI縺ｧOK縺ｪ繧・0 縺ｮ縺ｾ縺ｾ
mpADL.value = 0


' ・井ｻｻ諢擾ｼ迂ADL蛯呵・・IME蜀崎ｨｭ螳・
ApplyImeToIADLNote

' --- 荳九↓邯壹￥ nextTop 縺ｮ譖ｴ譁ｰ縺ｪ縺ｩ ---
Trace "update nextTop start", "Initialize"

nextTop = mpADL.Top + mpADL.Height + 10
hostMove.ScrollHeight = nextTop + 10

Trace "update nextTop done; mpADL.Top=" & mpADL.Top & _
      ", mpADL.Height=" & mpADL.Height & _
      ", nextTop=" & nextTop, "Initialize"


Dim y As Single, cboTmp As MSForms.ComboBox



    '================ 繝・せ繝医・隧穂ｾ｡ ================
    Trace "TESTS start", "Init"
    
    nextTop = pad
    Dim fTests As MSForms.Frame: Set fTests = CreateFrameP(hostTests, "繝・せ繝医・隧穂ｾ｡・育ｧ・蟆乗焚2譯√・謠｡蜉・蟆乗焚1譯√・0莉･荳奇ｼ・, 150)
Dim colL As Single, colR As Single
Dim row1 As Single, row2 As Single, row3 As Single
Dim memoW As Single

colL = 16
colR = 420
row1 = 18
row2 = 130
row3 = 242
memoW = 300

CreateLabel fTests, "10m豁ｩ陦・, colL, row1, 110
CreateLabel fTests, "遘・, colL, row1 + 22, 52
CreateTextBox fTests, colL + 58, row1 + 20, 80, 0, False, "txtTenMWalk", "Test.10m"
CreateLabel fTests, "遘・, colL + 144, row1 + 22, 20
CreateLabel fTests, "謇隕・, colL, row1 + 50, 52
CreateTextBox fTests, colL + 58, row1 + 48, memoW, 40, True, "txtMemo_10mWalk", "Memo.10m"

CreateLabel fTests, "TUG", colR, row1, 110
CreateLabel fTests, "遘・, colR, row1 + 22, 52
CreateTextBox fTests, colR + 58, row1 + 20, 80, 0, False, "txtTUG", "Test.TUG"
CreateLabel fTests, "遘・, colR + 144, row1 + 22, 20
CreateLabel fTests, "謇隕・, colR, row1 + 50, 52
CreateTextBox fTests, colR + 58, row1 + 48, memoW, 40, True, "txtMemo_TUG", "Memo.TUG"

CreateLabel fTests, "5蝗樒ｫ九■荳翫′繧・, colL, row2, 110
CreateLabel fTests, "遘・, colL, row2 + 22, 52
CreateTextBox fTests, colL + 58, row2 + 20, 80, 0, False, "txtFiveSTS", "Test.5xSTS"
CreateLabel fTests, "遘・, colL + 144, row2 + 22, 20
CreateLabel fTests, "謇隕・, colL, row2 + 50, 52
CreateTextBox fTests, colL + 58, row2 + 48, memoW, 40, True, "txtMemo_STS5", "Memo.5xSTS"

CreateLabel fTests, "繧ｻ繝溘ち繝ｳ繝・Β", colR, row2, 110
CreateLabel fTests, "遘・, colR, row2 + 22, 52
CreateTextBox fTests, colR + 58, row2 + 20, 80, 0, False, "txtSemi", "Test.Semi"
CreateLabel fTests, "遘・, colR + 144, row2 + 22, 20
CreateLabel fTests, "謇隕・, colR, row2 + 50, 52
CreateTextBox fTests, colR + 58, row2 + 48, memoW, 40, True, "txtMemo_SemiTandem", "Memo.Semi"

CreateLabel fTests, "謠｡蜉帛承", colL, row3, 110
CreateLabel fTests, "kg", colL, row3 + 22, 52
CreateTextBox fTests, colL + 58, row3 + 20, 80, 0, False, "txtGripR", "Grip.R"
CreateLabel fTests, "kg", colL + 144, row3 + 22, 20
CreateLabel fTests, "謇隕・, colL, row3 + 50, 52
CreateTextBox fTests, colL + 58, row3 + 48, memoW, 40, True, "txtMemo_GripR", "Memo.GripR"

CreateLabel fTests, "謠｡蜉帛ｷｦ", colR, row3, 110
CreateLabel fTests, "kg", colR, row3 + 22, 52
CreateTextBox fTests, colR + 58, row3 + 20, 80, 0, False, "txtGripL", "Grip.L"
CreateLabel fTests, "kg", colR + 144, row3 + 22, 20
CreateLabel fTests, "謇隕・, colR, row3 + 50, 52
CreateTextBox fTests, colR + 58, row3 + 48, memoW, 40, True, "txtMemo_GripL", "Memo.GripL"

ResizeFrameToContent fTests, row3 + 112


    Trace "TESTS end", "Init"

    
    '================ 豁ｩ陦瑚ｩ穂ｾ｡・郁・遶句ｺｦ / RLA・・================
Trace "WALK start", "Init"

Set mpWalk = hostWalk.controls.Add("Forms.MultiPage.1")

' 螟画焚螳｣險縺ｯ With 縺ｮ螟悶〒OK
Dim w As Single, h As Single

With mpWalk
    .Left = 0
    .Top = 0

    ' 蜀・ｯｸ繝吶・繧ｹ縺ｧ邂怜・縺励∽ｸ矩剞繧剃ｻ倥￠縺ｦ繧ｯ繝ｩ繝ｳ繝・
    w = hostWalk.InsideWidth - 12:  If w < 200 Then w = 200
    h = hostWalk.InsideHeight - 12: If h < 150 Then h = 150

    .Width = w
    .Height = h
    .Style = fmTabStyleTabs
End With

' 繝壹・繧ｸ(0),(1)繧定ｧｦ繧句燕縺ｫ2譫壻ｿ晁ｨｼ
Do While mpWalk.Pages.count < 2
    mpWalk.Pages.Add
Loop
mpWalk.Pages(0).caption = "閾ｪ遶句ｺｦ"
mpWalk.Pages(1).caption = "RLA"


    Set hostWalkGait = mpWalk.Pages(0).controls.Add("Forms.Frame.1", "hostWalkGait")
    With hostWalkGait
    .caption = ""
    .Left = 0: .Top = 0
    .Width = mpWalk.Width - 12
    Dim tGait As Single: tGait = mpWalk.Height - 30
    If tGait < 120 Then tGait = 120     ' 竊・縺薙％繧・t 竊・tGait 縺ｫ
    .Height = tGait                      ' 竊・縺薙％繧・t 竊・tGait 縺ｫ
    .ScrollBars = fmScrollBarsNone
End With


    nextTop = pad
        Set fGait = CreateFrameP(hostWalkGait, "豁ｩ陦瑚ｩ穂ｾ｡・郁・遶句ｺｦ・・, 90)
    
    fGait.name = "fGait"
    
    y = 22
    Dim rowH As Single: rowH = 28
    
    CreateLabel fGait, "豁ｩ陦瑚・遶句ｺｦ", COL_LX, y
    Dim cboGait As MSForms.ComboBox: Set cboGait = CreateCombo(fGait, COL_LX + lblW, y, 500, , "Gait.閾ｪ遶句ｺｦ")
    cboGait.List = MakeList("螳悟・閾ｪ遶・菫ｮ豁｣閾ｪ遶具ｼ郁｣懷勧蜈ｷ菴ｿ逕ｨ・・逶｣隕悶・隕句ｮ医ｊ,霆ｽ莉句勧・・5%譛ｪ貅・・荳ｭ遲牙ｺｦ莉句勧・・5-50%・・驥堺ｻ句勧・・0%莉･荳奇ｼ・蜈ｨ莉句勧")
    ResizeFrameToContent fGait, y + rowH

    Dim hostWalkRLA As MSForms.Frame
    Set hostWalkRLA = mpWalk.Pages(1).controls.Add("Forms.Frame.1")
    With hostWalkRLA: .caption = "": .Left = 0: .Top = 0: .Width = mpWalk.Width - 12: .Height = mpWalk.Height - 30: .ScrollBars = fmScrollBarsNone: End With

    Dim mpRLA As MSForms.MultiPage
    Set mpRLA = hostWalkRLA.controls.Add("Forms.MultiPage.1")
    With mpRLA
        .Left = 0: .Top = 0
        .Width = hostWalkRLA.Width - 6
        .Height = hostWalkRLA.Height - 6
        .Style = fmTabStyleTabs
    End With
    mpRLA.Pages(0).caption = "遶玖・譛滂ｼ・C-TSt・・
    mpRLA.Pages(1).caption = "驕願・譛滂ｼ・Sw-TSw・・

    Dim hostRLAStance As MSForms.Frame
    Set hostRLAStance = mpRLA.Pages(0).controls.Add("Forms.Frame.1")
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
    Dim fRLA1 As MSForms.Frame: Set fRLA1 = CreateFrameP(hostRLAStance, "RLA豁ｩ陦悟捉譛滂ｼ・C / LR / MSt / TSt・・, 280)
    Build_RLA_ChecksPart fRLA1, "stance": ResizeFrameToContent fRLA1, 260

    Dim hostRLASwing As MSForms.Frame
    Set hostRLASwing = mpRLA.Pages(1).controls.Add("Forms.Frame.1")
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
    Dim fRLA2 As MSForms.Frame: Set fRLA2 = CreateFrameP(hostRLASwing, "RLA豁ｩ陦悟捉譛滂ｼ・Sw / ISw / MSw / TSw・・, 280)
    Build_RLA_ChecksPart fRLA2, "swing": ResizeFrameToContent fRLA2, 260

    
    
    Trace "WALK end", "Init"



    '================ 隱咲衍繝ｻ邊ｾ逾・================
   Trace "COG start", "Init"
    
    nextTop = pad
    Dim fCog As MSForms.Frame: Set fCog = CreateFrameP(hostCog, "隱咲衍讖溯・繝ｻ邊ｾ逾樣擇", 110)
    y = 22
    CreateLabel fCog, "隱咲衍讖溯・繝ｬ繝吶Ν", COL_LX, y
    Dim cboCog As MSForms.ComboBox: Set cboCog = CreateCombo(fCog, COL_LX + lblW, y, 160, , "Cognition.Level")
    cboCog.List = MakeList("豁｣蟶ｸ,霆ｽ蠎ｦ菴惹ｸ・荳ｭ遲牙ｺｦ菴惹ｸ・鬮伜ｺｦ菴惹ｸ・)
    CreateLabel fCog, "邊ｾ逾樣擇", COL_RX, y
    Dim cboPsy As MSForms.ComboBox: Set cboPsy = CreateCombo(fCog, COL_RX + lblW, y, 160, , "Psych.Status")
    cboPsy.List = MakeList("螳牙ｮ・荳榊ｮ牙だ蜷・謚代≧縺､蛯ｾ蜷・縺昴・莉・)
    CreateLabel fCog, "蛯呵・, COL_LX, y + 28
    CreateTextBox fCog, COL_LX + lblW, y + 26, 610, 50, True, , "Cognition.蛯呵・
    ResizeFrameToContent fCog, y + 26 + 50
    
    Trace "COG end", "Init"




    '--- 荳矩Κ縺ｮ縲碁哩縺倥ｋ縲・---
    Trace "CLOSE start", "Init"

Set btnCloseCtl = Me.controls.Add("Forms.CommandButton.1")
With btnCloseCtl
    .caption = "髢峨§繧・
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



    
    

    '================ 蝓ｺ譛ｬ諠・ｱ・壼・髱｢繝ｬ繧､繧｢繧ｦ繝・================
    nextTop = pad
  Dim fBasic As MSForms.Frame: Set fBasic = CreateFrameP(hostBasic, "蝓ｺ譛ｬ諠・ｱ", 360)
    Set fBasicRef = fBasic
    y = 22

    ' --- 譛荳頑ｮｵ・壹Θ繝ｼ繝・ぅ繝ｪ繝・ぅ陦・---
    Dim chkDelta As MSForms.CheckBox
    Set chkDelta = CreateCheck(fBasic, "螟画峩轤ｹ縺ｮ縺ｿ菫晏ｭ假ｼ育ｩｺ谺・・蜑榊屓蛟､繧貞ｼ慕ｶ吶℃・・, COL_LX, 6, "chkDeltaOnly", "Delta.Only")
    chkDelta.AutoSize = True: chkDelta.WordWrap = False

    Set btnLoadPrevCtl = fBasic.controls.Add("Forms.CommandButton.1")
    With btnLoadPrevCtl
        .caption = "蜑榊屓縺ｮ蛟､繧定ｪｭ縺ｿ霎ｼ繧": .Accelerator = "L"
        .Width = 180: .Height = 24: .name = "btnLoadPrevCtl"
    End With
    Set btnSaveCtl = fBasic.controls.Add("Forms.CommandButton.1")
    With btnSaveCtl
        .caption = "繧ｷ繝ｼ繝医∈菫晏ｭ・: .Accelerator = "S"
        .Width = 120: .Height = 24: .name = "btnSaveCtl"
    End With
    PositionTopRightButtons fBasic
    nL y, 1

    ' 陦・・唔D / 隧穂ｾ｡譌･ / 隧穂ｾ｡閠・ｼ亥ｰ上＆繧・ｼ・
    CreateLabel fBasic, "ID", COL_LX, y
    Dim tbPID As MSForms.TextBox
    Set tbPID = CreateTextBox(fBasic, COL_LX + lblW, y, 120, 0, False, "txtPID", "PatientID")
    btnLoadPrevCtl.Left = tbPID.Left + tbPID.Width + 12
    btnLoadPrevCtl.Top = tbPID.Top
    CreateLabel fBasic, "隧穂ｾ｡譌･", COL_RX, y
    Dim tbED As MSForms.TextBox: Set tbED = CreateTextBox(fBasic, COL_RX + lblW, y, 120, 0, False, "txtEDate", "EvalDate")
    tbED.text = Format(Date, "yyyy/mm/dd")
    CreateLabel fBasic, "隧穂ｾ｡閠・, COL_RX + lblW + 130, y
    Dim tbEva As MSForms.TextBox: Set tbEva = CreateTextBox(fBasic, COL_RX + lblW + 180, y, 90, 0, False, "txtEvaluator", "Basic.Evaluator")
    tbEva.Font.Size = 8
    nL y

    ' 陦・・壽ｰ丞錐 / 蟷ｴ鮨｢ / 諤ｧ蛻･
    CreateLabel fBasic, "豌丞錐", COL_LX, y
    CreateTextBox fBasic, COL_LX + lblW, y, 200, 0, False, "txtName", "Basic.Name"
    CreateLabel fBasic, "蟷ｴ鮨｢", COL_RX, y
    CreateTextBox fBasic, COL_RX + lblW, y, 60, 0, False, "txtAge", "Basic.Age"
    CreateLabel fBasic, "諤ｧ蛻･", COL_RX + lblW + 70, y
    Dim cboSex As MSForms.ComboBox: Set cboSex = CreateCombo(fBasic, COL_RX + lblW + 110, y, 90, "cboSex", "Basic.Gender")
    cboSex.List = MakeList("逕ｷ諤ｧ,螂ｳ諤ｧ,縺昴・莉・荳肴・")
    nL y
    
    ' 陦・.5・夂函蟷ｴ譛域律
    CreateLabel fBasic, "逕溷ｹｴ譛域律", COL_RX, y
    Dim tbBirth As MSForms.TextBox
    Set tbBirth = CreateTextBox(fBasic, COL_RX + lblW, y, 120, 0, False, "txtBirth", "Basic.BirthDate")
    tbBirth.IMEMode = fmIMEModeOff


    nL y
   
      
   



    ' 陦・・壻ｸｻ險ｺ譁ｭ / 逋ｺ逞・律
    CreateLabel fBasic, "荳ｻ險ｺ譁ｭ", COL_LX, y
    CreateTextBox fBasic, COL_LX + lblW, y, 260, 0, False, "txtDx", "Basic.PrimaryDx"
    CreateLabel fBasic, "逋ｺ逞・律", COL_RX, y
    CreateTextBox fBasic, COL_RX + lblW, y, 120, 0, False, "txtOnset", "Basic.OnsetDate"
    nL y

    ' 陦・・夂函豢ｻ迥ｶ豕・/ 隕∽ｻ玖ｭｷ蠎ｦ
    CreateLabel fBasic, "逕滓ｴｻ迥ｶ豕・, COL_LX, y
   CreateTextBox fBasic, COL_LX + lblW, y, 220, 50, True, "txtLiving", "Basic.Living"
    CreateLabel fBasic, "隕∽ｻ玖ｭｷ蠎ｦ", COL_RX, y
    Dim cboLev As MSForms.ComboBox: Set cboLev = CreateCombo(fBasic, COL_RX + lblW, y, 150, "cboCare", "Basic.CareLevel")
    cboLev.List = MakeList("隕∵髪謠ｴ1,隕∵髪謠ｴ2,隕∽ｻ玖ｭｷ1,隕∽ｻ玖ｭｷ2,隕∽ｻ玖ｭｷ3,隕∽ｻ玖ｭｷ4,隕∽ｻ玖ｭｷ5")
    nL y

    ' 陦・・夐囿螳ｳ鬮倬ｽ｢閠・ｼ剰ｪ咲衍逞・ｫ倬ｽ｢閠・ｼ医Λ繝吶Ν蛟句挨蟷・ｼ・
    Dim LBLW_LONG_LEFT As Long:  LBLW_LONG_LEFT = 150
    Dim LBLW_LONG_RIGHT As Long: LBLW_LONG_RIGHT = 170

    CreateLabel fBasic, "髫懷ｮｳ鬮倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", COL_LX, y, LBLW_LONG_LEFT
    Dim cboEL As MSForms.ComboBox
    Set cboEL = CreateCombo(fBasic, COL_LX + LBLW_LONG_LEFT + 6, y, 180, "cboElder", "Basic.ElderlyLevel")
    cboEL.List = MakeList("閾ｪ遶・J1,J2,A1,A2,B1,B2,C1,C2")

    CreateLabel fBasic, "隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", COL_RX, y, LBLW_LONG_RIGHT
    Dim cboDL As MSForms.ComboBox
    Set cboDL = CreateCombo(fBasic, COL_RX + LBLW_LONG_RIGHT + 6, y, 160, "cboDementia", "Basic.DementiaLevel")
    cboDL.List = MakeList("閾ｪ遶・I,IIa,IIb,IIIa,IIIb,IV,M")
    nL y, 1

    ' 陬懷勧蜈ｷ・上Μ繧ｹ繧ｯ・・蛻暦ｼ・
    nL y, 1

    Dim frRisk As MSForms.Frame
    Dim ASSISTIVE_CSV As String: ASSISTIVE_CSV = "譚・繧ｷ繝ｫ繝舌・繧ｫ繝ｼ,豁ｩ陦悟勣,霆翫＞縺・遏ｭ荳玖い陬・・,莉句勧繝吶Ν繝・謇九☆繧・繧ｹ繝ｭ繝ｼ繝・
    Dim RISK_CSV As String: RISK_CSV = "霆｢蛟・隱､蝴･,螟ｱ遖・隍･逖｡,菴取・､・蠕伜ｾ・縺帙ｓ螯・ADL菴惹ｸ・

    Set frRisk = BuildCheckFrame(fBasic, "繝ｪ繧ｹ繧ｯ", COL_RX, y, 370, MakeList(RISK_CSV), "RiskGroup")



    Dim nextY As Single
    nextY = frRisk.Top + frRisk.Height
    y = nextY + 8

    BuildAssistiveChecksInWalkEval ASSISTIVE_CSV

    ' Needs・亥ｷｦ蜿ｳ2繧ｫ繝ｩ繝・・
    Dim needsH As Single: needsH = 36
    CreateLabel fBasic, "謔｣閠・eeds", COL_LX, y
    CreateTextBox fBasic, COL_LX + lblW, y, 270, needsH, True, "txtNeedsPt", "Needs.Patient"
    CreateLabel fBasic, "螳ｶ譌蒐eeds", COL_RX, y
    CreateTextBox fBasic, COL_RX + lblW, y, 270, needsH, True, "txtNeedsFam", "Needs.Family"
    y = y + needsH + 10
    
    

    ResizeFrameToContent fBasic, y
    
    
    ' 繝ｬ繧､繧｢繧ｦ繝茨ｼ・ち繝夜・
    FitLayout
    If Not fBasicRef Is Nothing Then PositionTopRightButtons fBasicRef
    ResetTabOrder

    ' 蛻晄悄縺ｮ菫晏ｭ倥・繧ｿ繝ｳ豢ｻ諤ｧ迥ｶ諷九ｒ蜿肴丐
    RefreshSaveEnabled
    Me.controls("btnSaveCtl").Enabled = True  ' 竊・荳譌ｦ縲∽ｿ晏ｭ倥・繧ｿ繝ｳ繧貞ｸｸ譎よ怏蜉ｹ蛹・
    
    

    '================ 蟋ｿ蜍｢隧穂ｾ｡・亥､牙ｽ｢隕∝屏縺ｯ蛯呵・∈・・================
    nextTop = pad
    Dim fPost As MSForms.Frame
    Set fPost = CreateFrameP(hostPost, "蟋ｿ蜍｢隧穂ｾ｡・亥､牙ｽ｢隕∝屏縺ｯ蛯呵・∈・・, 200)

    Dim yP As Single: yP = 22

    ' 1陦檎岼
    CreateCheck fPost, "鬆ｭ驛ｨ蜑肴婿遯∝・", COL_LX, yP, , "Posture.鬆ｭ驛ｨ蜑肴婿遯∝・"
    CreateCheck fPost, "蜀・レ", COL_LX + 150, yP, , "Posture.蜀・レ"
    CreateCheck fPost, "蛛ｴ蠑ｯ", COL_LX + 300, yP, , "Posture.蛛ｴ蠑ｯ"
    CreateLabel fPost, "鬪ｨ逶､蛯ｾ譁・, COL_RX, yP - 2
    Dim cboPel As MSForms.ComboBox
    Set cboPel = CreateCombo(fPost, COL_RX + lblW, yP - 2, 120, , "Posture.鬪ｨ逶､蛯ｾ譁・)
    cboPel.List = MakeList("蜑榊だ,蠕悟だ,豁｣蟶ｸ,荳肴・")
    nL yP

    ' 2陦檎岼
    CreateCheck fPost, "菴灘ｹｹ蝗樊雷", COL_LX, yP, , "Posture.菴灘ｹｹ蝗樊雷"
    CreateCheck fPost, "蜿榊ｼｵ閹・, COL_LX + 150, yP, , "Posture.蜿榊ｼｵ閹・

    ' 蛯呵・ｼ亥承荳奇ｼ・
    CreateLabel fPost, "蛯呵・, COL_RX, yP - 2
    CreateTextBox fPost, COL_RX + lblW + 10, yP - 4, 190, 50, True, "", "Posture.蛯呵・

    nL yP, 3
    ResizeFrameToContent fPost, yP

    nextTop = fPost.Top + fPost.Height + 6

    '================ 髢｢遽諡倡ｸｮ・亥承繝ｻ蟾ｦ・会ｼ句ｙ閠・================
    Dim fCon As MSForms.Frame: Set fCon = CreateFrameP(hostPost, "髢｢遽諡倡ｸｮ・亥承繝ｻ蟾ｦ・会ｼ句ｙ閠・, 180)
    Dim y0 As Single: y0 = 22
    y = y0

    ' 繧ｬ繧､繝芽ｦ句・縺・
    CreateLabel fCon, "驛ｨ菴・, COL_LX, y
    CreateLabel fCon, "蜿ｳ", COL_LX + 90 + 20, y
    CreateLabel fCon, "蟾ｦ", COL_LX + 90 + 20 + 60, y
    nL y

    CreateCheck fCon, "鬆ｸ驛ｨ・亥ｷｦ蜿ｳ縺ｪ縺暦ｼ・, COL_LX, y, "", "Contracture.鬆ｸ驛ｨ": nL y

    ' 驛ｨ菴阪＃縺ｨ縺ｫR/L繝√ぉ繝・け
    CreateContractureRLRow fCon, y, "閧ｩ髢｢遽", "Contracture.閧ｩ"
    CreateContractureRLRow fCon, y, "閧倬未遽", "Contracture.閧・
    CreateContractureRLRow fCon, y, "謇矩未遽", "Contracture.謇矩未遽"
    CreateContractureRLRow fCon, y, "閧｡髢｢遽", "Contracture.閧｡髢｢遽"
    CreateContractureRLRow fCon, y, "閹晞未遽", "Contracture.閹晞未遽"
    CreateContractureRLRow fCon, y, "雜ｳ髢｢遽", "Contracture.雜ｳ髢｢遽"

    ' 蛯呵・ｼ亥承蛛ｴ荳企Κ縺ｫ・・
    CreateLabel fCon, "蛯呵・, COL_RX, y0 - 2
    CreateTextBox fCon, COL_RX + lblW, y0 - 4, 250, 80, True, "", "Contracture.蛯呵・

    ' 鬮倥＆隱ｿ謨ｴ
    ResizeFrameToContent fCon, Application.WorksheetFunction.Max(y, y0 + 80)


  SetupInputModesJP
  
'== ROM蜀・・遨ｺ繝壹・繧ｸ(Page14/15)繧定ｵｷ蜍墓凾縺ｫ閾ｪ蜍募炎髯､ ==
Dim ctlZ As Object, mpZ As MSForms.MultiPage, iZ As Long, capZ As String
For Each ctlZ In Me.controls
    If TypeName(ctlZ) = "MultiPage" Then
        Set mpZ = ctlZ
        For iZ = mpZ.Pages.count - 1 To 0 Step -1
            capZ = CStr(mpZ.Pages(iZ).caption)
            If capZ = "Page14" Or capZ = "Page15" Then mpZ.Pages.Remove iZ
        Next iZ
    End If
Next ctlZ


'== 蛯呵・ｬ・ｼ医Λ繝吶Ν・句､ｧ縺阪＞TextBox・峨ｒ髱櫁｡ｨ遉ｺ ==

Dim mpN As Object, pgN As Object, pN As Object, cN As Object

' --- ROM繝壹・繧ｸ迚ｹ螳夲ｼ・ROM" 縺ｾ縺溘・ "荳ｻ隕・未遽"・・---
Set mpN = Nothing: Set pgN = Nothing
For Each cN In Me.controls
    If TypeName(cN) = "MultiPage" Then
        Set mpN = cN
        Exit For
    End If
Next cN

If Not mpN Is Nothing Then
    For Each pN In mpN.Pages
    If InStr(1, CStr(pN.caption), "ROM", vbTextCompare) > 0 _
       Or InStr(1, CStr(pN.caption), "荳ｻ隕・未遽", vbTextCompare) > 0 Then
        Set pgN = pN: Exit For
    End If
Next pN

End If



If Not pgN Is Nothing Then
    Dim stk As Collection: Set stk = New Collection
    Dim parent As Object, ctl As Object
    Dim nLbl As Object, nTB As Object, isML As Boolean

    ' 蟄舌さ繝ｳ繝・リ繧貞性繧√※豺ｱ縺募━蜈医〒謗｢邏｢
    stk.Add pgN
    Do While stk.count > 0
        Set parent = stk(1): stk.Remove 1
        On Error Resume Next
        For Each ctl In parent.controls
            On Error GoTo 0
            Select Case TypeName(ctl)
                Case "Frame", "MultiPage", "Page"
                    stk.Add ctl                      ' 蟄舌ｒ縺溘←繧・
                Case "Label"
                    If InStr(1, CStr(ctl.caption), "蛯呵・, vbTextCompare) > 0 Then
                        Set nLbl = ctl              ' 縲悟ｙ閠・蛯呵・ｬ・阪ｒ謐墓拷
                    End If
                Case "TextBox"
                    isML = False
                    On Error Resume Next: isML = ctl.multiline: On Error GoTo 0
                    If isML Or ctl.Height >= 80 Or ctl.Width >= 400 Then
                        If nTB Is Nothing Then
                            Set nTB = ctl
                        ElseIf ctl.Height * ctl.Width > nTB.Height * nTB.Width Then
                            Set nTB = ctl          ' 譛螟ｧ繧ｵ繧､繧ｺ繧貞ｙ閠・→縺励※謗｡逕ｨ
                        End If
                    End If
            End Select
        Next ctl
    Loop

    ' 隕九▽縺九▲縺溘ｉ髱櫁｡ｨ遉ｺ
    If Not nLbl Is Nothing Then nLbl.Visible = False
    If Not nTB Is Nothing Then nTB.Visible = False
End If


'=== 蛯呵・Λ繝吶Ν縺ｨ蛯呵・ユ繧ｭ繧ｹ繝医ｒ蜀榊ｸｰ逧・↓髱櫁｡ｨ遉ｺ・亥・繧ｳ繝ｳ繝・リ蟇ｾ蠢懶ｼ・===
Dim qH As New Collection, parentH As Object, ctlH As Object, iH As Long
Dim noteTB As Object, areaMax As Double




' ROM繝壹・繧ｸ繧偵◎縺ｮ蝣ｴ縺ｧ迚ｹ螳壹＠縺ｦ縺九ｉ髢句ｧ・
Dim rootPg As Object, c0 As Object, mp0 As Object, i0 As Long
Set rootPg = Nothing
For Each c0 In Me.controls
    If TypeName(c0) = "MultiPage" Then
        Set mp0 = c0
        For i0 = 0 To mp0.Pages.count - 1
            If InStr(1, CStr(mp0.Pages(i0).caption), "ROM", vbTextCompare) > 0 _
               Or InStr(1, CStr(mp0.Pages(i0).caption), "荳ｻ隕・未遽", vbTextCompare) > 0 Then
                Set rootPg = mp0.Pages(i0): Exit For
            End If
        Next i0
        If Not rootPg Is Nothing Then Exit For
    End If
Next c0
If rootPg Is Nothing Then Exit Sub
qH.Add rootPg

Do While qH.count > 0
    Set parentH = qH(1): qH.Remove 1

    If TypeName(parentH) = "MultiPage" Then
        ' MultiPage 縺ｯ Controls 繧呈戟縺溘↑縺・・縺ｧ Pages 繧貞句挨縺ｫ霎ｿ繧・
        For iH = 0 To parentH.Pages.count - 1
            qH.Add parentH.Pages(iH)
        Next iH
    Else
        On Error Resume Next
        For Each ctlH In parentH.controls
            On Error GoTo 0
           Select Case TypeName(ctlH)
    Case "Frame", "Page"
        qH.Add ctlH                      ' 蟄舌さ繝ｳ繝・リ繧定ｾｿ繧・
    Case "MultiPage"
        For iH = 0 To ctlH.Pages.count - 1
            qH.Add ctlH.Pages(iH)        ' 繝壹・繧ｸ繧貞句挨縺ｫ霎ｿ繧・
        Next iH
    Case "Label"
        If InStr(1, CStr(ctlH.caption), "蛯呵・, vbTextCompare) > 0 Then
            Set nLbl = ctlH
            Set noteTB = ctlH
            ctlH.Visible = False         ' 蛯呵・Λ繝吶Ν繧呈ｶ医☆
        End If
    Case "TextBox"
        Dim mlH As Boolean: mlH = False
        On Error Resume Next: mlH = ctlH.multiline: On Error GoTo 0
        If mlH Or ctlH.Height >= 80 Or ctlH.Width >= 400 Then
            ctlH.Visible = False         ' 螟ｧ縺阪＞繝・く繧ｹ繝・OX繧呈ｶ医☆
        End If
End Select

        Next ctlH
    End If
Loop

If Not noteTB Is Nothing Then noteTB.Visible = False
'Debug.Print "[ROM_Note] hidden lbl=" & (Not (nLbl Is Nothing)) & "  tb=" & (Not (noteTB Is Nothing))

'=== /蛯呵・撼陦ｨ遉ｺ ===

Call ROM_Fix_TextBoxHeight_Recursive_OnROM_Once
Call ROM_CheckBoxes_Up12_OnROM_Recursive_Once_V2
Call MMT_BuildChildTabs_Direct


Dim c As Control
For Each c In Me.controls
    If TypeName(c) = "Label" Then
        If c.caption = "NRS" Then
            c.caption = "螳蛾撕譎・RS"
            Exit For
        End If
    End If
Next

'--- 蜍穂ｽ懈凾NRS繧定・蜍戊ｿｽ蜉・・蝗槭□縺托ｼ・---
Dim srcLbl As MSForms.label, srcCmb As MSForms.ComboBox
Dim lbl As MSForms.label, cmb As MSForms.ComboBox
Dim ct As Control

' 螳蛾撕譎・RS繝ｩ繝吶Ν繧堤音螳・
For Each ct In Me.controls
    If TypeName(ct) = "Label" Then
        If ct.caption = "螳蛾撕譎・RS" Then
            Set srcLbl = ct
            Exit For
        End If
    End If
Next

If Not srcLbl Is Nothing Then
    ' 螳蛾撕譎・RS縺ｮ蜿ｳ蛛ｴ縺ｫ縺ゅｋ譌｢蟄呂ombo繧呈耳螳夲ｼ亥酔縺倬ｫ倥＆ﾂｱ6・・
    For Each ct In Me.controls
        If TypeName(ct) = "ComboBox" Then
            If Abs(ct.Top - srcLbl.Top) <= 20 And ct.Left > srcLbl.Left Then
                Set srcCmb = ct: Exit For
            End If
        End If
    Next

    ' 譌｢縺ｫ菴懈・貂医∩縺ｪ繧我ｽ輔ｂ縺励↑縺・
    For Each ct In Me.controls
        If TypeName(ct) = "Label" And ct.name = "lblNRS_Move" Then Set lbl = ct
        If TypeName(ct) = "ComboBox" And ct.name = "cmbNRS_Move" Then Set cmb = ct
    Next

    If lbl Is Nothing Then
       Set lbl = srcLbl.parent.controls.Add("Forms.Label.1", "lblNRS_Move", True)
        lbl.caption = "蜍穂ｽ懈凾NRS"
    End If

    If cmb Is Nothing Then
       Set cmb = srcLbl.parent.controls.Add("Forms.ComboBox.1", "cmbNRS_Move", True)
        cmb.Style = fmStyleDropDownList
        Dim i As Long
        For i = 0 To 10: cmb.AddItem CStr(i): Next i
    End If

    ' 菴咲ｽｮ豎ｺ繧・ｼ亥ｮ蛾撕譎・RS縺ｮ縲御ｸ九阪↓驟咲ｽｮ・・
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

' Align Rest NRS row above Move NRS to avoid drift
If Not srcLbl Is Nothing Then
    srcLbl.Left = lbl.Left
    srcLbl.Top = lbl.Top - srcLbl.Height - gap
End If
If Not srcCmb Is Nothing Then
    srcCmb.Left = cmb.Left
    srcCmb.Top = srcLbl.Top - 2
End If

End If
  
End Sub


Sub ShowFrame12()
    Dim f As Control, t As Control
    For Each f In frmEval.controls
        If TypeName(f) = "Frame" Then
            For Each t In f.controls
                If TypeName(t) = "TextBox" And t.name = "TextBox2" Then
                    f.ZOrder 0                 '荳逡ｪ謇句燕縺ｫ
                    f.caption = "笘・％繧後′Frame12笘・ '隕九▽縺代ｄ縺吶￥縺吶ｋ
                    Beep
                    Exit Sub
                End If
            Next
        End If
    Next
    MsgBox "TextBox2 縺ｮ隕ｪ繝輔Ξ繝ｼ繝縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・
End Sub













Public Sub AddPainQualUI()
    Dim host As MSForms.Frame
    Dim cap As MSForms.label
    Dim lb  As MSForms.ListBox
    Dim items As Variant
    Dim i As Long

    ' 譌｢蟄倥メ繧ｧ繝・け・磯㍾隍・函謌宣亟豁｢・・
    Dim t As Object
    Set t = FindCtlDeep(Me, "lstPainQual")
    If Not t Is Nothing Then Exit Sub


    ' 譌｢蟄朗RS縺ｮ隕ｪ繝輔Ξ繝ｼ繝縺ｫ霑ｽ蜉縺吶ｋ
    Set host = GetPainHost()

    ' 隕句・縺・
    Set cap = host.controls.Add("Forms.Label.1", "lblPainQual", True)
    cap.caption = "逞帙∩縺ｮ諤ｧ雉ｪ・郁､・焚驕ｸ謚槫庄・・
    cap.Left = 12
    cap.Top = 12
    cap.AutoSize = True

    ' 繝ｪ繧ｹ繝医・繝・け繧ｹ・郁､・焚驕ｸ謚橸ｼ・
    Set lb = host.controls.Add("Forms.ListBox.1", "lstPainQual", True)
    lb.Left = 12
    lb.Top = cap.Top + cap.Height + 6
    lb.Width = 240
    lb.Height = 96
    lb.MultiSelect = fmMultiSelectMulti
    lb.IntegralHeight = False

    ' 驕ｸ謚櫁い
    items = Array("驤咲李", "蛻ｺ縺吶ｈ縺・↑逞帙∩", "縺励・繧・, "轣ｼ辭ｱ諢・, "繧ｺ繧ｭ繧ｺ繧ｭ", _
                  "邱繧∽ｻ倥￠諢・, "蝨ｧ霑ｫ諢・, "蠑輔″縺､繧・, "髮ｻ謦・李", "縺薙ｏ縺ｰ繧・, "縺代＞繧後ｓ", "縺昴・莉・)
    For i = LBound(items) To UBound(items)
        lb.AddItem items(i)
    Next i
End Sub
















Public Sub AddPainFactorsUI()
    Dim host As MSForms.Frame
    Dim fr As MSForms.Frame
    Dim cap As MSForms.label
    Dim i As Long, y As Single

    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainFactors")
    If Not t Is Nothing Then Exit Sub

    ' 霑ｽ蜉蜈医・逍ｼ逞帙ち繝悶・繝輔Ξ繝ｼ繝・・RS縺ｮ隕ｪ・・
    Set host = GetPainHost()

    ' 隕句・縺・
    Set cap = host.controls.Add("Forms.Label.1", "lblPainFactors", True)
    cap.caption = "隱伜屏繝ｻ霆ｽ貂帛屏蟄・
    cap.Left = 270
    cap.Top = 280
    cap.AutoSize = True

    ' 繧ｳ繝ｳ繝・リ繝輔Ξ繝ｼ繝
    Set fr = host.controls.Add("Forms.Frame.1", "fraPainFactors", True)
    fr.Left = cap.Left
    fr.Top = cap.Top + cap.Height + 4
    fr.Width = 360
    fr.Height = 120

    ' 蟾ｦ蛻暦ｼ夊ｪ伜屏・・rovoking・・
    Dim provItems As Variant, relItems As Variant
    provItems = Array( _
        Array("chkPainProv_Move", "蜍穂ｽ・), _
        Array("chkPainProv_Posture", "蟋ｿ蜍｢"), _
        Array("chkPainProv_Walk", "豁ｩ陦・), _
        Array("chkPainProv_Lift", "謖√■荳翫￡"), _
        Array("chkPainProv_Cough", "蜥ｳ/縺上＠繧・∩") _
    )

    ' 蜿ｳ蛻暦ｼ夊ｻｽ貂幢ｼ・elieving・・
    relItems = Array( _
        Array("chkPainRelief_Rest", "螳蛾撕"), _
        Array("chkPainRelief_Heat", "貂ｩ辭ｱ"), _
        Array("chkPainRelief_Cold", "蜀ｷ蜊ｴ"), _
        Array("chkPainRelief_Med", "譛崎脈"), _
        Array("chkPainRelief_Brace", "繧ｳ繝ｫ繧ｻ繝・ヨ") _
    )

    ' 蟾ｦ蛻鈴・鄂ｮ
    y = 8
    For i = LBound(provItems) To UBound(provItems)
        Dim cb As MSForms.CheckBox
        Set cb = fr.controls.Add("Forms.CheckBox.1", CStr(provItems(i)(0)), True)
        cb.caption = CStr(provItems(i)(1))
        cb.Left = 12
        cb.Top = y
        y = y + cb.Height + 2
    Next i

    ' 蜿ｳ蛻鈴・鄂ｮ
    y = 8
    For i = LBound(relItems) To UBound(relItems)
        Dim cb2 As MSForms.CheckBox
        Set cb2 = fr.controls.Add("Forms.CheckBox.1", CStr(relItems(i)(0)), True)
        cb2.caption = CStr(relItems(i)(1))
        cb2.Left = fr.Width \ 2 + 8
        cb2.Top = y
        y = y + cb2.Height + 2
    Next i
End Sub

Public Sub AddVASUI()
    Dim host As MSForms.Frame
    Dim cap As MSForms.label
    Dim fr As MSForms.Frame
    Dim tb As MSForms.TextBox
    Dim sb As MSForms.ScrollBar

    ' 譌｢蟄倥メ繧ｧ繝・け・磯㍾隍・函謌宣亟豁｢・・
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraVAS")
    If Not t Is Nothing Then Exit Sub
    
    ' 霑ｽ蜉蜈茨ｼ晉名逞帙ち繝悶・繝輔Ξ繝ｼ繝・・RS縺ｮ隕ｪ・・
    Set host = GetPainHost()

    ' 隕句・縺・
    Set cap = host.controls.Add("Forms.Label.1", "lblVAS", True)
    cap.caption = "VAS・・?100・・
    cap.Left = 640
    cap.Top = 280
    cap.AutoSize = True

    ' 繧ｳ繝ｳ繝・リ
    Set fr = host.controls.Add("Forms.Frame.1", "fraVAS", True)
    fr.Left = cap.Left
    fr.Top = cap.Top + cap.Height + 4
    fr.Width = 122
    fr.Height = 64

    ' 繝・く繧ｹ繝医・繝・け繧ｹ・域焚蛟､ 0?100・・
    Set tb = fr.controls.Add("Forms.TextBox.1", "txtVAS", True)
    tb.Left = 8
    tb.Top = 10
    tb.Width = 40
    tb.text = "0"

    ' 繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ繝舌・・域ｨｪ・・?100
    Set sb = fr.controls.Add("Forms.ScrollBar.1", "sldVAS", True)
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
    EvalCtl("txtVAS").text = CStr(v)

End Sub


Public Sub WireVAS()
    On Error Resume Next
    Set mVAS = EvalCtl("sldVAS")
    If mVAS Is Nothing Then Exit Sub


End Sub





Public Sub AddPainCourseUI()
    Dim host As MSForms.Frame
    Dim lb As MSForms.label
    Dim cb As MSForms.ComboBox
    Dim tb As MSForms.TextBox
    Dim i As Long

    ' 譌｢蟄倥メ繧ｧ繝・け・磯㍾隍・函謌宣亟豁｢・・
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainCourse")
    If Not t Is Nothing Then Exit Sub

    ' 霑ｽ蜉蜈茨ｼ晉名逞帙ち繝悶・繝輔Ξ繝ｼ繝・・RS縺ｮ隕ｪ・・
    Set host = GetPainHost()

    ' 隕句・縺・
    Set lb = host.controls.Add("Forms.Label.1", "lblPainCourse", True)
    lb.caption = "逞帙∩縺ｮ邨碁℃繝ｻ譎る俣螟牙喧"
    lb.Left = 12
    lb.Top = 280
    lb.AutoSize = True

    ' 繧ｳ繝ｳ繝・リ
    Dim fr As MSForms.Frame
    Set fr = host.controls.Add("Forms.Frame.1", "fraPainCourse", True)
    fr.Left = lb.Left
    fr.Top = lb.Top + lb.Height + 4
    fr.Width = 610
    fr.Height = 78

    ' 逋ｺ逞・凾譛・
    Set lb = fr.controls.Add("Forms.Label.1", "lblPainOnset", True)
    lb.caption = "逋ｺ逞・凾譛・
    lb.Left = 12: lb.Top = 10: lb.AutoSize = True

    Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbPainOnset", True)
    cb.Left = lb.Left + 60: cb.Top = 8: cb.Width = 140
    cb.List = Array("諤･諤ｧ・・1騾ｱ・・, "莠懈･諤ｧ・・3縺区怦・・, "諷｢諤ｧ・・縺区怦?・・, "蜀咲㏍・丞・逋ｺ", "荳肴・")

    ' 謖∫ｶ壽凾髢・
    Set lb = fr.controls.Add("Forms.Label.1", "lblPainDuration", True)
    lb.caption = "謖∫ｶ・
    lb.Left = 260: lb.Top = 10: lb.AutoSize = True

    Set tb = fr.controls.Add("Forms.TextBox.1", "txtPainDuration", True)
    tb.Left = lb.Left + 36: tb.Top = 8: tb.Width = 40: tb.text = ""

    Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbPainDurationUnit", True)
    cb.Left = tb.Left + tb.Width + 6: cb.Top = 8: cb.Width = 70
    cb.List = Array("譌･", "騾ｱ", "縺区怦", "蟷ｴ")

    ' 譌･蜀・､牙虚
    Set lb = fr.controls.Add("Forms.Label.1", "lblPainDayPeriod", True)
    lb.caption = "譌･蜀・､牙虚"
    lb.Left = 12: lb.Top = 38: lb.AutoSize = True

    Set cb = fr.controls.Add("Forms.ComboBox.1", "cmbPainDayPeriod", True)
    cb.Left = lb.Left + 54: cb.Top = 36: cb.Width = 260
    cb.List = Array("譛昴↓蠑ｷ縺・, "譏ｼ縺ｫ蠑ｷ縺・, "螟懊↓蠑ｷ縺・, "蜈･豬ｴ蠕後↓霆ｽ貂・, "豢ｻ蜍募ｾ後↓蠅玲が", "荳螳壹〒螟牙喧縺ｪ縺・)
End Sub








Public Sub AddPainSiteUI()
    Dim host As MSForms.Frame
    Dim lb As MSForms.label
    Dim fr As MSForms.Frame
    Dim lst As MSForms.ListBox
    Dim i As Long
    Dim items As Variant

    ' 譌｢蟄倥メ繧ｧ繝・け・磯㍾隍・函謌宣亟豁｢・・
    Dim t As Object
    Set t = FindCtlDeep(Me, "fraPainSite")
    If Not t Is Nothing Then Exit Sub

    ' 霑ｽ蜉蜈茨ｼ晉名逞帙ち繝悶・繝輔Ξ繝ｼ繝・・RS縺ｮ隕ｪ・・
    Set host = GetPainHost()

    ' 隕句・縺・
    Set lb = host.controls.Add("Forms.Label.1", "lblPainSite", True)
    lb.caption = "逍ｼ逞幃Κ菴搾ｼ郁､・焚驕ｸ謚槫庄・・
    lb.Left = 12
    lb.Top = 380
    lb.AutoSize = True

    ' 繝輔Ξ繝ｼ繝
    Set fr = host.controls.Add("Forms.Frame.1", "fraPainSite", True)
    fr.Left = lb.Left
    fr.Top = lb.Top + lb.Height + 4
    fr.Width = 360
    fr.Height = 140

    ' 繝ｪ繧ｹ繝茨ｼ郁､・焚驕ｸ謚橸ｼ・
    Set lst = fr.controls.Add("Forms.ListBox.1", "lstPainSite", True)
    lst.Left = 12
    lst.Top = 10
    lst.Width = fr.Width - 24
    lst.Height = fr.Height - 20
    lst.MultiSelect = fmMultiSelectMulti
    lst.IntegralHeight = False

    ' 驛ｨ菴榊呵｣・
    items = Array( _
        "鬆ｭ驛ｨ", "鬆ｸ驛ｨ", "閧ｩ", "閧ｩ逕ｲ驛ｨ", "荳願・", "閧・, "蜑崎・", "謇矩ｦ・, "謇・謖・, _
        "閭ｸ驛ｨ", "閭碁Κ荳企Κ", "閭碁Κ荳矩Κ・郁・・・, _
        "鬪ｨ逶､驛ｨ/莉呵・驛ｨ", _
        "閧｡", "螟ｧ閻ｿ", "閹・, "荳玖・", "雜ｳ鬥・, "雜ｳ/雜ｾ" _
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

    Set fr = GetPainHost()
    
    ' 蜿ら・蜿門ｾ・
    On Error Resume Next
    Set lbQ = fr.controls("lstPainQual")
    Set lbS = SafeGetControl(fr, "fraPainSite").controls("lstPainSite")
    Set frF = SafeGetControl(fr, "fraPainFactors")
    vas = SafeGetControl(fr, "fraVAS").controls("txtVAS").text
    onset = SafeGetControl(fr, "fraPainCourse").controls("cmbPainOnset").text
    dura = SafeGetControl(fr, "fraPainCourse").controls("txtPainDuration").text
    unit = SafeGetControl(fr, "fraPainCourse").controls("cmbPainDurationUnit").text
    day = SafeGetControl(fr, "fraPainCourse").controls("cmbPainDayPeriod").text
    On Error GoTo 0

    ' 逞帙∩縺ｮ諤ｧ雉ｪ
    s = "縲千名逞帙∪縺ｨ繧√・
    s = s & " 諤ｧ雉ｪ: "
    If Not lbQ Is Nothing Then
        For i = 0 To lbQ.ListCount - 1
            If lbQ.Selected(i) Then s = s & lbQ.List(i) & "・・
        Next
        If Right$(s, 1) = "・・ Then s = Left$(s, Len(s) - 1)
    End If

    ' 驛ｨ菴・
    s = s & "・憺Κ菴・ "
    If Not lbS Is Nothing Then
        Dim tmpS As String: tmpS = ""
        For i = 0 To lbS.ListCount - 1
            If lbS.Selected(i) Then tmpS = tmpS & lbS.List(i) & "・・
        Next
        If tmpS <> "" Then tmpS = Left$(tmpS, Len(tmpS) - 1)
        s = s & tmpS
    End If

    ' 隱伜屏繝ｻ霆ｽ貂帛屏蟄・
    s = s & "・懷屏蟄・ "
    If Not frF Is Nothing Then
        Dim tmpF As String: tmpF = ""
        For Each c In frF.controls
            If TypeName(c) = "CheckBox" Then
                If c.value = True Then tmpF = tmpF & c.caption & "・・
            End If
        Next
        If tmpF <> "" Then tmpF = Left$(tmpF, Len(tmpF) - 1)
        s = s & tmpF
    End If

    ' 邨碁℃・儀AS
    If onset <> "" Then s = s & "・懃匱逞・ " & onset
    If dura <> "" Or unit <> "" Then s = s & "・懈戟邯・ " & dura & unit
    If day <> "" Then s = s & "・懈律蜀・ " & day
    If vas <> "" Then s = s & "・弖AS: " & vas

    ' 繝｡繝｢縺ｸ蜿肴丐
    On Error Resume Next
   EvalCtl("txtPainMemo").text = s
End Sub



Private Sub mBtnPainSum_Click()
    SummarizePainUI
End Sub


' 逍ｼ逞帙ち繝・ Frame12 )縺ｫ谿九▲縺ｦ縺・ｋ譌ｧUI繧帝勁蜴ｻ縺吶ｋ
Public Sub RemoveLegacyPainUI()
    Dim f As MSForms.Frame, n As Variant
    Set f = GetPainHost()
    If f Is Nothing Then Exit Sub

    ' Probe縺ｧ[LEGACY?]縺ｨ蛻､螳壹＆繧後◆繧ゅ・縺縺大炎髯､・域眠UI繧НRS/蛯呵・・谿九☆・・
    For Each n In Array("Label85", "TextBox1", "Label86", "Label87", "ComboBox39", "TextBox2", "txtPainMemo_lbl", "txtPainMemo", "lblNRS_Move", "cmbNRS_Move")
        On Error Resume Next
        f.controls.Remove CStr(n)
        If Err.Number = 0 Then Debug.Print "[removed]", n Else Debug.Print "[skip]", n, "err", Err.Number
        Err.Clear
        On Error GoTo 0
    Next n

    Debug.Print "[done] RemoveLegacyPainUI"
   
    
End Sub




Public Sub ArrangePainLayout()
If GetPainHost Is Nothing Then Exit Sub

    Dim f As MSForms.Frame
   Set f = GetPainHost()
    If f Is Nothing Then Exit Sub

    ' 荳頑ｮｵ・壼ｷｦ・晄ｧ雉ｪ縲∝承・抃AS
    With f.controls("lblPainQual")
        .Left = 12: .Top = 12: .ZOrder 0
    End With
    With f.controls("lstPainQual")
        .Left = 12
        .Top = f.controls("lblPainQual").Top + f.controls("lblPainQual").Height + 4
        .Width = 360: .Height = 120
        .ZOrder 0
    End With
    With f.controls("lblVAS")
        .Left = 420: .Top = 12: .ZOrder 0
    End With
    With SafeGetControl(f, "fraVAS")
        .Left = 420
        .Top = f.controls("lblVAS").Top + f.controls("lblVAS").Height + 4
        .ZOrder 0
    End With

    ' 荳ｭ谿ｵ・壼ｷｦ・晉ｵ碁℃縲∝承・晁ｪ伜屏繝ｻ霆ｽ貂・
    With f.controls("lblPainCourse")
        .Left = 12: .Top = 160: .ZOrder 0
    End With
    With SafeGetControl(f, "fraPainCourse")
        .Left = 12
        .Top = f.controls("lblPainCourse").Top + f.controls("lblPainCourse").Height + 4
        .Width = 360
        .ZOrder 0
    End With
    With f.controls("lblPainFactors")
        .Left = 420: .Top = 160: .ZOrder 0
    End With
    With SafeGetControl(f, "fraPainFactors")
        .Left = 420
        .Top = f.controls("lblPainFactors").Top + f.controls("lblPainFactors").Height + 4
        .Width = 330
        .ZOrder 0
    End With

    ' 荳区ｮｵ・壼ｷｦ・晞Κ菴阪∵怙荳区ｮｵ・晏ｙ閠・
    With f.controls("lblPainSite")
        .Left = 12: .Top = 300: .ZOrder 0
    End With
    With SafeGetControl(f, "fraPainSite")
        .Left = 12
        .Top = f.controls("lblPainSite").Top + f.controls("lblPainSite").Height + 4
        .Width = 360: .Height = 140
        .ZOrder 0
    End With
    With f.controls("txtPainMemo_lbl")
        .Top = 470: .ZOrder 0
    End With
    With f.controls("txtPainMemo")
        .Top = 492: .ZOrder 0
    End With
End Sub



Sub RemoveLegacyPainUI_Final()
    Dim fr As MSForms.Frame, c As Control
    Set fr = GetPainHost()
    If fr Is Nothing Then Exit Sub
    
    For Each c In fr.controls
        Select Case c.name
            Case "Label85", "TextBox1", "Label86", "Label87", _
                 "ComboBox39", "TextBox2", "txtPainMemo_lbl", "txtPainMemo"
                Debug.Print "[remove]", c.name
                fr.controls.Remove c.name
        End Select
    Next
    
    Debug.Print "[done] RemoveLegacyPainUI_Final"
End Sub









Public Sub MatchPainFrameHeights()
    Dim z As MSForms.Frame, pf As MSForms.Frame, ps As MSForms.Frame, lb As MSForms.label
    Set z = GetPainHost()
    If z Is Nothing Then Exit Sub
    Set pf = SafeGetControl(z, "fraPainFactors")
    Set ps = SafeGetControl(z, "fraPainSite")
    Set lb = z.controls("lblPainFactors")

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





'== frmEval 縺ｮ繧ｳ繝ｼ繝峨↓雋ｼ繧贋ｻ倥￠ ==
Public Sub TidyPainBoxes()
    Const gap As Single = 24

    Dim z As MSForms.Frame
    Dim ps As MSForms.Frame, pf As MSForms.Frame
    Dim lbPS As MSForms.label, lbPF As MSForms.label

    Set z = GetPainHost()
    If z Is Nothing Then Exit Sub
    Set ps = SafeGetControl(z, "fraPainSite")
    Set pf = SafeGetControl(z, "fraPainFactors")
    Set lbPS = z.controls("lblPainSite")
    Set lbPF = z.controls("lblPainFactors")

    ' 逍ｼ逞幃Κ菴阪・繝ｩ繝吶Ν繧呈棧逶ｴ荳翫↓謠・∴繧具ｼ井ｽ咲ｽｮ縺ｯ迴ｾ迥ｶ邯ｭ謖√↑繧我ｸ崎ｦ・ｼ・
    lbPS.Left = ps.Left: lbPS.Top = ps.Top - lbPS.Height - 4

    ' 隱伜屏繝ｻ霆ｽ貂帛屏蟄舌ｒ縲檎名逞幃Κ菴阪・蜿ｳ髫｣縲阪↓驟咲ｽｮ
    pf.Top = ps.Top
    pf.Left = ps.Left + ps.Width + gap
    pf.Height = ps.Height   ' 鬮倥＆荳閾ｴ

    ' 繝ｩ繝吶Ν繧よ棧縺ｮ逶ｴ荳翫↓
    lbPF.Left = pf.Left
    lbPF.Top = pf.Top - lbPF.Height - 4

   



End Sub





'--------------------------------------------
' 豎守畑: 繝輔Ξ繝ｼ繝蜀・さ繝ｳ繝医Ο繝ｼ繝ｫ蜿門ｾ暦ｼ・othing險ｱ螳ｹ・・
Private Function GetCtl(ByVal host As Object, ByVal name As String) As Object
    On Error Resume Next
    Set GetCtl = host.controls(name)
    On Error GoTo 0
End Function

' 逞帙∩縺ｮ邨碁℃繝ｻ譎る俣繝悶Ο繝・け繧呈￡荵・Ξ繧､繧｢繧ｦ繝・
Public Sub TidyPainCourse()
    Dim f As MSForms.Frame
    Dim frCourse As MSForms.Frame, lbCourse As MSForms.label
    Dim frSite As MSForms.Frame, frFactors As MSForms.Frame
    Dim L0 As Single, T0 As Single, m As Single, gap As Single
    Dim wLeftCol As Single, rightEdge As Single
    
     Set f = GetPainHost()
    If f Is Nothing Then Exit Sub

    ' 蜿ら・・亥ｭ伜惠縺励↑縺・ｴ蜷医・菴輔ｂ縺励↑縺・ｼ・
    Set lbCourse = GetCtl(f, "lblPainCourse")
    Set frCourse = GetCtl(f, "fraPainCourse")
    Set frSite = GetCtl(f, "fraPainSite")
    Set frFactors = GetCtl(f, "fraPainFactors")

    If lbCourse Is Nothing Or frCourse Is Nothing Then Exit Sub

    ' 繝ｬ繧､繧｢繧ｦ繝亥ｮ壽焚
    m = 12        ' 繝輔Ξ繝ｼ繝蜀・・蟾ｦ蜿ｳ繝槭・繧ｸ繝ｳ
    gap = 8       ' 繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ髢薙・繧ｮ繝｣繝・・
    L0 = frSite.Left                        ' 蟾ｦ蛻励・髢句ｧ倶ｽ咲ｽｮ・育名逞幃Κ菴阪→蜷医ｏ縺帙ｋ・・
    wLeftCol = frSite.Width                 ' 蟾ｦ蛻励・蟷・ｒ逍ｼ逞幃Κ菴阪→荳閾ｴ縺輔○繧・

    ' 縲瑚ｪ伜屏繝ｻ霆ｽ貂帛屏蟄舌阪ｒ蜿ｳ蛻励↓荳九￡縺溷燕謠舌〒縲∝ｷｦ蛻励・蟷・〒蠎・￡繧・
    lbCourse.Left = L0
    frCourse.Left = L0
    frCourse.Width = wLeftCol               ' 竊・蟾ｦ蛻励→蜷後§蟷・↓諱剃ｹ・喧
    ' 鬮倥＆縺ｯ荳ｭ縺ｮ繝ｬ繧､繧｢繧ｦ繝医↓萓晏ｭ倥ょｿ・ｦ√↑繧画怙蠕後↓閾ｪ蜍輔〒閭御ｸ郁ｪｿ謨ｴ繧ょ庄

    ' 邵ｦ菴咲ｽｮ・夂李縺ｿ縺ｮ諤ｧ雉ｪ縺ｮ荳具ｼ育樟蝨ｨ驟咲ｽｮ縺輔ｌ縺ｦ縺・ｋ菴咲ｽｮ縺九ｉ蟆代＠隧ｰ繧√ｋ・・
    ' 縺薙％縺ｧ縺ｯ縲檎李縺ｿ縺ｮ諤ｧ雉ｪ縲阪Μ繧ｹ繝医・荳狗ｫｯ・倶ｽ咏區縺ｧ蜷医ｏ縺帙ｋ
    Dim lstQual As MSForms.ListBox
    Set lstQual = GetCtl(f, "lstPainQual")
    If Not lstQual Is Nothing Then
        lbCourse.Top = lstQual.Top + lstQual.Height + 20
    Else
        ' 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・育樟蝨ｨ縺ｮTop繧貞ｰ企㍾・・
        lbCourse.Top = lbCourse.Top
    End If
    frCourse.Top = lbCourse.Top + lbCourse.Height + gap

    ' --- 蜀・Κ鬆・岼縺ｮ荳ｦ縺ｳ・域里蟄倥・逶ｸ蟇ｾ驟咲ｽｮ繧堤ｶｭ謖√＠縺､縺､蟷・□縺題ｿｽ蠕難ｼ・---
    '   [逋ｺ逞・凾譛歉 [謖∫ｶ・謨ｰ蛟､][蜊倅ｽ江     竊・谿ｵ逶ｮ
    '   [譌･蜀・､牙虚: ・ｿ・ｿ・ｿ・ｿ・ｿ・ｿ・ｿ・ｿ・ｿ・ｿ ] 竊・谿ｵ逶ｮ 蟷・＞縺｣縺ｱ縺・
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

    ' 1谿ｵ逶ｮ縺ｯ譌｢蟄倥・Left繧剃ｽｿ縺・▽縺､縲∝承遶ｯ縺後・縺ｿ蜃ｺ縺輔↑縺・ｈ縺・↓蠕ｮ隱ｿ謨ｴ
    Dim right1 As Single
    right1 = cmbUnit.Left + cmbUnit.Width
    If right1 > frCourse.Width - m Then
        cmbUnit.Left = frCourse.Width - m - cmbUnit.Width
        ' 謨ｰ蛟､繝ｻ繝ｩ繝吶Ν繧ょｷｦ縺ｸ隧ｰ繧√ｋ
        If Not txtDur Is Nothing Then txtDur.Left = cmbUnit.Left - gap - txtDur.Width
        If Not lblDur Is Nothing Then lblDur.Left = txtDur.Left - gap - lblDur.Width - 4   ' 竊・笘・ｰ代＠菴咏區
    Else
        ' 笘・壼ｸｸ譎ゅｂ霆ｽ縺丞ｷｦ縺ｫ菴咏區繧偵→繧具ｼ郁｢ｫ繧企亟豁｢・・
        If Not lblDur Is Nothing Then lblDur.Left = txtDur.Left - gap - lblDur.Width - 4
    End If

    ' 2谿ｵ逶ｮ・医梧律蜀・､牙虚縲阪さ繝ｳ繝懶ｼ峨・譫蜀・＞縺｣縺ｱ縺・↓
    If Not lblDay Is Nothing Then lblDay.Left = m
    If Not cmbDay Is Nothing Then
        cmbDay.Left = IIf(lblDay Is Nothing, m, lblDay.Left + lblDay.Width + gap)
        cmbDay.Width = frCourse.Width - m - cmbDay.Left
        If cmbDay.Width < 80 Then cmbDay.Width = 80
    End If

    ' 鬮倥＆繧定・蜍戊ｪｿ謨ｴ・域怙荳矩Κ縺ｮ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ荳狗ｫｯ・九・繝ｼ繧ｸ繝ｳ・・
    Dim bottomY As Single
    bottomY = 0
    Dim c As Control
    For Each c In frCourse.controls
        bottomY = Application.WorksheetFunction.Max(bottomY, c.Top + c.Height)
    Next
    frCourse.Height = bottomY + m
End Sub
'--------------------------------------------


Public Sub WidenAndTidyPainCourse()
    ' 逶ｸ莠貞他縺ｳ蜃ｺ縺励↑縺暦ｼ乗焔蜍募ｮ溯｡悟燕謠・
    Dim f As MSForms.Frame
   Set f = SafeGetControl(Me, "fraPainCourse")
    If f Is Nothing Then Exit Sub

    With f
        ' 蜿ら・
        Dim cmbOnset As MSForms.ComboBox
        Dim lblDur As MSForms.label
        Dim txtDur As MSForms.TextBox
        Dim cmbUnit As MSForms.ComboBox

        Set cmbOnset = .controls("cmbPainOnset")
        Set lblDur = .controls("lblPainDuration")
        Set txtDur = .controls("txtPainDuration")
        Set cmbUnit = .controls("cmbPainDurationUnit")

        ' 讓ｪ荳ｦ縺ｳ縺ｮ遒ｺ螳壹Ο繧ｸ繝・け・井ｻ雁屓縲悟ｮ檎挑縲阪↓縺ｪ縺｣縺溷ｼ上→蜷後§・・
        lblDur.Left = cmbOnset.Left + cmbOnset.Width + 12
        txtDur.Left = lblDur.Left + lblDur.Width + 12
        cmbUnit.Left = txtDur.Left + txtDur.Width + 8
       

      

        ' 蜿ｳ菴咏區繧・4pt遒ｺ菫・
        .Width = cmbUnit.Left + cmbUnit.Width + 24
    End With


End Sub


Public Sub TidyPainUI_Once()
    If mPainTidyBusy Then Exit Sub
    mPainTidyBusy = True
    On Error GoTo Clean



    Me.TidyPainBoxes   '窶ｻ蜀・Κ縺ｧ WidenAndTidyPainCourse 繧・蝗槭□縺大他縺ｶ

Clean:

    mPainTidyBusy = False
    Call FixPainCaptionsAndWidth   ' 竊・縺薙・陦後□縺題ｿｽ蜉・・・・
    
    '=== Pain headings finalize (once) ===
Dim f As MSForms.Frame
Dim L As MSForms.label

On Error Resume Next
'--- 逶ｴ荳・---
Set L = Me.controls("lblVAS"):           If Not L Is Nothing Then L.WordWrap = False: L.caption = "VAS・・・・00・・: L.WordWrap = False: L.AutoSize = False: L.Width = 120
Set L = Me.controls("lblPainQual"):      If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
Set L = Me.controls("lblPainCourse"):    If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
Set L = Me.controls("lblPainSite"):      If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150
Set L = Me.controls("lblPainFactors"):   If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150

'--- Frame3 蜀・---
Set f = SafeGetControl(Me, "Frame3")
If Not f Is Nothing Then
    Set L = f.controls("lblVAS"):         If Not L Is Nothing Then L.WordWrap = False: L.caption = "VAS・・・・00・・: L.WordWrap = False: L.AutoSize = False: L.Width = 120
    Set L = f.controls("lblPainQual"):    If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
    Set L = f.controls("lblPainCourse"):  If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
    Set L = f.controls("lblPainSite"):    If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150
    Set L = f.controls("lblPainFactors"): If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150
End If

'--- Frame12 蜀・---
Set f = SafeGetControl(Me, "Frame12")
If Not f Is Nothing Then
    Set L = f.controls("lblVAS"):         If Not L Is Nothing Then L.WordWrap = False: L.caption = "VAS・・・・00・・: L.WordWrap = False: L.AutoSize = False: L.Width = 120
    Set L = f.controls("lblPainQual"):    If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
    Set L = f.controls("lblPainCourse"):  If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 140
    Set L = f.controls("lblPainSite"):    If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150
    Set L = f.controls("lblPainFactors"): If Not L Is Nothing Then L.WordWrap = False: L.AutoSize = False: L.Width = 150
End If
On Error GoTo 0
'=== /Pain headings finalize ===


End Sub
Private Sub FixPainCaptionsAndWidth()
    Dim c As Control
    For Each c In Me.controls
        If TypeName(c) = "Frame" Then
            ' 蟾ｦ・夊ｦ句・縺暦ｼ育李縺ｿ縺ｮ諤ｧ雉ｪ窶ｦ・峨ｒ謚倥ｊ霑斐＆縺ｪ縺・ｹ・↓縺吶ｋ・郁ｦｪ蜿ｳ遶ｯ?24pt・・
            If InStr(c.caption, "逞・) > 0 And InStr(c.caption, "諤ｧ雉ｪ") > 0 Then
                On Error Resume Next
                c.Width = c.parent.InsideWidth - c.Left - 24
                On Error GoTo 0
            End If
            ' 蜿ｳ・啖AS縺ｮ陦ｨ險倥ｒ蝗ｺ螳・
            If InStr(c.caption, "VAS") > 0 Then
                c.caption = "VAS・・・・00・・
            End If
        End If
    Next
End Sub

Public Sub FixPainLabels_Final()
    Dim f As Control, c As Control, L As Object

    '--- 逶ｴ荳九・繝ｩ繝吶Ν繧貞・逅・---
    For Each c In Me.controls
        If TypeName(c) = "Label" Then
            If c.name = "lblPainQual" Then
                Set L = c
                On Error Resume Next
                CallByName L, "AutoSize", VbLet, True   ' 蠢・ｦ∝ｹ・ｒ蜿門ｾ・
                CallByName L, "AutoSize", VbLet, False  ' 蝗ｺ螳壹↓謌ｻ縺呻ｼ域釜霑斐＠髦ｲ豁｢・・
                On Error GoTo 0
            ElseIf c.name = "lblVAS" Then
                c.caption = "VAS・・・・00・・
            End If
        End If
    Next

    '--- 蜷Ёrame蜀・・繝ｩ繝吶Ν繧貞・逅・---
    For Each f In Me.controls
        If TypeName(f) = "Frame" Then
            For Each c In f.controls
                If TypeName(c) = "Label" Then
                    If c.name = "lblPainQual" Then
                        Set L = c
                        On Error Resume Next
                        CallByName L, "AutoSize", VbLet, True
                        CallByName L, "AutoSize", VbLet, False
                        On Error GoTo 0
                    ElseIf c.name = "lblVAS" Then
                        c.caption = "VAS・・・・00・・
                    End If
                End If
            Next
        End If
    Next
End Sub


Public Sub ListToneKeyCaptions()
    Dim c As Control
    For Each c In frmEval.controls
        On Error Resume Next
        If TypeName(c) = "CheckBox" Or TypeName(c) = "OptionButton" Or TypeName(c) = "Label" Then
            Dim cap As String: cap = CStr(c.caption)
            If InStr(cap, "MAS_") > 0 Or InStr(cap, "蜿榊ｰЮ") > 0 Then
                Debug.Print "[TONE-CTL]", TypeName(c), c.name, "|", cap
            End If
        End If
        On Error GoTo 0
    Next
End Sub

Private Function GetWalkBaseTop(ByVal f As MSForms.Frame) As Single
    Dim ctl As Object
    Dim bestTop As Single

    GetWalkBaseTop = -1
    If f Is Nothing Then Exit Function

    bestTop = 99999
    For Each ctl In f.controls
        If TypeName(ctl) = "ComboBox" Or TypeName(ctl) = "Label" Then
            If ctl.Top < bestTop Then bestTop = ctl.Top
        End If
    Next

    If bestTop < 99999 Then GetWalkBaseTop = bestTop
End Function

Private Function SafeGetPage(ByVal mp As Object, ByVal pageKey As Variant) As Object
    Dim i As Long
    Dim pg As Object
    Dim keyText As String
    Dim capText As String

    If mp Is Nothing Then Exit Function
    
    On Error Resume Next
    If IsNumeric(pageKey) Then
        i = CLng(pageKey)
        If i >= 0 And i < mp.Pages.count Then
            Set SafeGetPage = mp.Pages(i)
            Exit Function
        End If
    End If
    On Error GoTo 0

    keyText = Trim$(CStr(pageKey))
    If LenB(keyText) = 0 Then Exit Function

    For i = 0 To mp.Pages.count - 1
        Set pg = mp.Pages(i)

        If StrComp(CStr(pg.name), keyText, vbTextCompare) = 0 Then
            Set SafeGetPage = pg
            Exit Function
        End If

        capText = Trim$(CStr(pg.caption))
        If StrComp(capText, keyText, vbTextCompare) = 0 Then
            Set SafeGetPage = pg
            Exit Function
        End If
    Next i
End Function


Private Sub BuildWalkIndep_DistanceOutdoor()
    Dim f As MSForms.Frame
    Dim ctl As MSForms.Control
    Dim cmbBase As MSForms.ComboBox
    Dim lblDist As MSForms.label
    Dim lblOut As MSForms.label
    Dim cmbDist As MSForms.ComboBox
    Dim cmbOut As MSForms.ComboBox
    Dim leftLabel As Single
    Dim leftCombo As Single
    Dim wLabel As Single
    Dim wCombo As Single
    Dim topDist As Single
    Dim topOut As Single
    
    Set f = GetWalkAssistiveTargetFrame()
    If f Is Nothing Then Exit Sub

    For Each ctl In f.controls
        If TypeName(ctl) = "ComboBox" Then
            Set cmbBase = ctl
            Exit For
        End If
    Next
    
    If cmbBase Is Nothing Then Exit Sub
    cmbBase.tag = "WalkIndepLevel"

    leftLabel = 12
    leftCombo = cmbBase.Left
    wLabel = 60
    wCombo = cmbBase.Width
    topDist = cmbBase.Top + 24
    topOut = topDist + 24

    On Error Resume Next
    Set lblDist = f.controls("lblWalkDistance")
    Set cmbDist = f.controls("cmbWalkDistance")
    Set lblOut = f.controls("lblWalkOutdoor")
    Set cmbOut = f.controls("cmbWalkOutdoor")
    On Error GoTo 0

    If lblDist Is Nothing Then Set lblDist = CreateLabel(f, "", leftLabel, topDist, wLabel, "lblWalkDistance")
    If cmbDist Is Nothing Then Set cmbDist = CreateCombo(f, leftCombo, topDist, wCombo, "cmbWalkDistance", "WalkDistance")
    If lblOut Is Nothing Then Set lblOut = CreateLabel(f, "", leftLabel, topOut, wLabel, "lblWalkOutdoor")
    If cmbOut Is Nothing Then Set cmbOut = CreateCombo(f, leftCombo, topOut, wCombo, "cmbWalkOutdoor", "WalkOutdoor")

    If cmbDist.ListCount = 0 Then
      cmbDist.AddItem "5m譛ｪ貅"
      cmbDist.AddItem "5m莉･荳・0m譛ｪ貅"
      cmbDist.AddItem "10m莉･荳・0m譛ｪ貅"
      cmbDist.AddItem "50m莉･荳・
    End If

    If cmbOut.ListCount = 0 Then
      cmbOut.AddItem "螻句､匁ｭｩ陦悟庄"
      cmbOut.AddItem "螻句､悶ｂ遏ｭ霍晞屬縺ｪ繧牙庄"
      cmbOut.AddItem "螻句､悶・隕句ｮ医ｊ縺ｧ蜿ｯ"
      cmbOut.AddItem "螻句､悶・莉句勧縺悟ｿ・ｦ・
    End If

    lblDist.caption = "豁ｩ陦瑚ｷ晞屬"
    lblOut.caption = "螻句､匁ｭｩ陦・


    lblDist.Left = leftLabel: lblDist.Top = topDist: lblDist.Width = wLabel
    cmbDist.Left = leftCombo: cmbDist.Top = topDist: cmbDist.Width = wCombo
    lblOut.Left = leftLabel: lblOut.Top = topOut: lblOut.Width = wLabel
    cmbOut.Left = leftCombo: cmbOut.Top = topOut: cmbOut.Width = wCombo


    BuildWalkIndep_Stability
    BuildWalkIndep_Speed

End Sub



Private Sub BuildWalkIndep_Stability()
    Dim f As MSForms.Frame
    Dim top1 As Single, top2 As Single, top3 As Single, top4 As Single
    Dim chk As MSForms.CheckBox
    Dim leftPos As Single
    Dim baseTop As Single

    Set f = GetWalkAssistiveTargetFrame()
    If f Is Nothing Then Exit Sub

    baseTop = GetWalkBaseTop(f)
    If baseTop < 0 Then Exit Sub

    top1 = baseTop
    top2 = top1 + 24
    top3 = top2 + 24
    top4 = top3 + 24

    Dim i As Long
    For i = f.controls.count - 1 To 0 Step -1
        If StrComp(Left$(f.controls(i).name, 12), "chkWalkStab_", vbTextCompare) = 0 Then
            f.controls.Remove f.controls(i).name
        End If
    Next i


    If f.Height < top4 + 24 Then f.Height = top4 + 24
    
    leftPos = 12

    
    Set chk = f.controls.Add("Forms.CheckBox.1", "chkWalkStab_Furatsuki", True)
    chk.caption = "縺ｵ繧峨▽縺阪≠繧・: chk.Left = leftPos: chk.Top = top4: chk.Width = 90: chk.Height = 18
    
    leftPos = leftPos + chk.Width + 12
    

    
    Set chk = f.controls.Add("Forms.CheckBox.1", "chkWalkStab_Foot", True)
    chk.caption = "雜ｳ驕九・荳榊ｮ牙ｮ・: chk.Left = leftPos: chk.Top = top4: chk.Width = 100: chk.Height = 18
    
    leftPos = leftPos + chk.Width + 12
    
    
    Set chk = f.controls.Add("Forms.CheckBox.1", "chkWalkStab_Turn", True)
    chk.caption = "譁ｹ蜷題ｻ｢謠帑ｸ榊ｮ・: chk.Left = leftPos: chk.Top = top4: chk.Width = 100: chk.Height = 18
    
    leftPos = leftPos + chk.Width + 12
    
    
    
    Set chk = f.controls.Add("Forms.CheckBox.1", "chkWalkStab_Slow", True)
    chk.caption = "騾溷ｺｦ菴惹ｸ・: chk.Left = leftPos: chk.Top = top4: chk.Width = 80: chk.Height = 18
    
    leftPos = leftPos + chk.Width + 12


    Set chk = f.controls.Add("Forms.CheckBox.1", "chkWalkStab_FallRisk", True)
    chk.caption = "霆｢蛟偵Μ繧ｹ繧ｯ鬮倥＞": chk.Left = leftPos: chk.Top = top4: chk.Width = 110: chk.Height = 18
End Sub


Private Sub BuildWalkIndep_Speed()
    Dim f As MSForms.Frame
    Dim baseTop As Single
    Dim top5 As Single
    Dim cmb As MSForms.ComboBox

    Set f = GetWalkAssistiveTargetFrame()
    If f Is Nothing Then Exit Sub

    baseTop = GetWalkBaseTop(f)
    If baseTop < 0 Then Exit Sub

    top5 = baseTop + 96

    On Error Resume Next
    f.controls.Remove "cmbGaitSpeedDetail"
    On Error GoTo 0
    
    CreateLabel f, "豁ｩ陦碁溷ｺｦ", 12, top5, 60
    Set cmb = CreateCombo(f, 84, top5, 200, , "cmbGaitSpeedDetail")

    If cmb.ListCount = 0 Then
          cmb.AddItem "騾溘＞"
          cmb.AddItem "繧・ｄ騾溘＞"
          cmb.AddItem "譎ｮ騾・
          cmb.AddItem "繧・ｄ驕・＞"
          cmb.AddItem "驕・＞"
    End If

    If f.Height < top5 + 42 Then f.Height = top5 + 42
End Sub

Private Sub BuildWalk_AbnormalTab()
    Dim ctl As MSForms.Control
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page

    ' 豁ｩ陦瑚ｩ穂ｾ｡逕ｨ縺ｮ MultiPage2 繧呈爾縺・
    For Each ctl In Me.controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "MultiPage2" Then
                Set mp = ctl
                Exit For
            End If
        End If
    Next

    If mp Is Nothing Then Exit Sub

    ' 譌｢縺ｫ縲檎焚蟶ｸ豁ｩ陦後阪ち繝悶′縺ゅｌ縺ｰ菴輔ｂ縺励↑縺・
    On Error Resume Next
    Set pg = SafeGetPage(mp, "pgWalkAbnormal")
    On Error GoTo 0
    If Not pg Is Nothing Then Exit Sub

    ' 譁ｰ縺励＞繝壹・繧ｸ繧定ｿｽ蜉・・ndex=2諠ｳ螳夲ｼ夊・遶句ｺｦ, RLA 縺ｮ谺｡・・
    Set pg = mp.Pages.Add
    pg.name = "pgWalkAbnormal"
    pg.caption = "逡ｰ蟶ｸ豁ｩ陦・

    ' 縺ｲ縺ｨ縺ｾ縺壻ｸｭ縺ｫ遨ｺ縺ｮ繝輔Ξ繝ｼ繝縺縺醍ｽｮ縺・※縺翫￥・井ｸｭ霄ｫ縺ｯ蠕後〒菴懊ｋ・・
    Dim f As MSForms.Frame
    Set f = pg.controls.Add("Forms.Frame.1", "fraWalkAbnormal", True)
    With f
        .caption = "逡ｰ蟶ｸ豁ｩ陦後ヱ繧ｿ繝ｼ繝ｳ・医メ繧ｧ繝・け・・
        .Left = 6
        .Top = 6
        .Width = mp.Width - 24
        .Height = mp.Height - 24
    End With
End Sub


Private Sub BuildWalkAbnormal_Frames()
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim w As Single, h As Single

    ' MultiPage2・域ｭｩ陦瑚ｩ穂ｾ｡・峨ｒ蜿門ｾ・
    For Each ctl In Me.controls
        If TypeName(ctl) = "MultiPage" And ctl.name = "MultiPage2" Then
            Set mp = ctl
            Exit For
        End If
    Next
    If mp Is Nothing Then Exit Sub

    ' 逡ｰ蟶ｸ豁ｩ陦後・繝ｼ繧ｸ蜿門ｾ・
    On Error Resume Next
    Set pg = SafeGetPage(mp, "pgWalkAbnormal")
    On Error GoTo 0
    If pg Is Nothing Then Exit Sub

    ' 繝壹・繧ｸ縺ｮ繝ｯ繝ｼ繧ｯ繧ｨ繝ｪ繧｢繧ｵ繧､繧ｺ
    w = mp.Width - 24
    h = mp.Height - 24

    ' 譌｢蟄倥ヵ繝ｬ繝ｼ繝蜑企勁・亥・逕滓・逕ｨ・・
    For Each ctl In pg.controls
        If TypeName(ctl) = "Frame" Then
            pg.controls.Remove ctl.name
        End If
    Next

    ' --- A・夂援鮗ｻ逞ｺ邉ｻ ---
    Set f = pg.controls.Add("Forms.Frame.1", "fraWalkAbn_A", True)
    With f
        .caption = "A・夂援鮗ｻ逞ｺ繝ｻ閼ｳ陦邂｡髫懷ｮｳ繝代ち繝ｼ繝ｳ"
        .Left = 6
        .Top = 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- B・壹ヱ繝ｼ繧ｭ繝ｳ繧ｽ繝ｳ邉ｻ ---
    Set f = pg.controls.Add("Forms.Frame.1", "fraWalkAbn_B", True)
    With f
        .caption = "B・壹ヱ繝ｼ繧ｭ繝ｳ繧ｽ繝ｳ髢｢騾｣繝代ち繝ｼ繝ｳ"
        .Left = w / 2 + 6
        .Top = 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- C・壽紛蠖｢繝ｻ鬮倬ｽ｢閠・ｸ榊ｮ牙ｮ壽ｭｩ陦・---
    Set f = pg.controls.Add("Forms.Frame.1", "fraWalkAbn_C", True)
    With f
        .caption = "C・壽紛蠖｢繝ｻ鬮倬ｽ｢閠・ｸ榊ｮ牙ｮ壽ｭｩ陦・
        .Left = 6
        .Top = h / 2 + 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With

    ' --- D・壼鵠隱ｿ髫懷ｮｳ繝ｻ螟ｱ隱ｿ ---
    Set f = pg.controls.Add("Forms.Frame.1", "fraWalkAbn_D", True)
    With f
        .caption = "D・壼鵠隱ｿ髫懷ｮｳ繝ｻ螟ｱ隱ｿ繝代ち繝ｼ繝ｳ"
        .Left = w / 2 + 6
        .Top = h / 2 + 6
        .Width = w / 2 - 12
        .Height = h / 2 - 12
    End With
End Sub


Private Sub BuildWalkAbnormal_Checks()
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim ctl As MSForms.Control
    Dim f As MSForms.Frame
    Dim chk As MSForms.CheckBox
    Dim items As Variant
    Dim i As Long, topPos As Single
    
    ' MultiPage2 繧貞叙蠕・
    For Each ctl In Me.controls
        If TypeName(ctl) = "MultiPage" And ctl.name = "MultiPage2" Then
            Set mp = ctl
            Exit For
        End If
    Next
    If mp Is Nothing Then Exit Sub

    ' 逡ｰ蟶ｸ豁ｩ陦後・繝ｼ繧ｸ繧貞叙蠕・
    On Error Resume Next
    Set pg = SafeGetPage(mp, "pgWalkAbnormal")
    On Error GoTo 0
    If pg Is Nothing Then Exit Sub

    ' 4縺､縺ｮ繝輔Ξ繝ｼ繝繧帝・分縺ｫ蜃ｦ逅・
    Dim fNames As Variant
    fNames = Array("fraWalkAbn_A", "fraWalkAbn_B", "fraWalkAbn_C", "fraWalkAbn_D")
    
    Dim A_items As Variant
    Dim B_items As Variant
    Dim C_items As Variant
    Dim D_items As Variant

    ' ----------------------
    ' A・夂援鮗ｻ逞ｺ繝ｻ閼ｳ陦邂｡髫懷ｮｳ・・・・0・・
    ' ----------------------
    A_items = Array( _
        "縺吶ｊ雜ｳ豁ｩ陦・, _
        "縺ｶ繧灘屓縺玲ｭｩ陦・, _
        "蜿榊ｼｵ閹晄ｭｩ陦鯉ｼ郁・驕惹ｼｸ螻包ｼ・, _
        "繝医Ξ繝ｳ繝・Ξ繝ｳ繝悶Ν繧ｰ豁ｩ陦・, _
        "繝・Η繧ｷ繧ｧ繝ｳ繝梧ｭｩ陦・, _
        "荳句桙雜ｳ・医ヵ繝・ヨ繧ｹ繝ｩ繝・・・・, _
        "蜈ｱ蜷碁°蜍輔ヱ繧ｿ繝ｼ繝ｳ縺ｮ蠑ｷ縺・, _
        "鬪ｨ逶､蠕悟だ・丈ｽ灘ｹｹ蠕悟だ縺ｮ遶玖・", _
        "迚・・遶玖・縺ｮ闡励＠縺・洒邵ｮ", _
        "雜ｳ驛ｨ繧ｯ繝ｪ繧｢繝ｩ繝ｳ繧ｹ荳崎憶" _
    )

    ' ----------------------
    ' B・壹ヱ繝ｼ繧ｭ繝ｳ繧ｽ繝ｳ邉ｻ・・・・・・
    ' ----------------------
    B_items = Array( _
        "蟆丞綾縺ｿ豁ｩ陦・, _
        "蜑榊だ蟋ｿ蜍｢豁ｩ陦・, _
        "繝輔Μ繝ｼ繧ｺ・育ｪ∫┯蛛懈ｭ｢・・, _
        "豁ｩ陦碁幕蟋句峅髮｣・医せ繧ｿ繝ｼ繝・hesitation・・, _
        "遯・ｲ豁ｩ陦鯉ｼ医ヵ繧ｧ繧ｹ繝・ぅ繝阪・繧ｷ繝ｧ繝ｳ・・, _
        "豁ｩ蟷・ｸ帛ｰ・, _
        "謇九・謖ｯ繧頑ｶ亥､ｱ", _
        "譁ｹ蜷題ｻ｢謠帛峅髮｣", _
        "繝ｪ繧ｺ繝諤ｧ豸亥､ｱ" _
    )

    ' ----------------------
    ' C・壽紛蠖｢繝ｻ鬮倬ｽ｢閠・ｸ榊ｮ牙ｮ壽ｭｩ陦鯉ｼ・・・0・・
    ' ----------------------
    C_items = Array( _
        "繧医■繧医■豁ｩ陦鯉ｼ育ｭ句鴨菴惹ｸ具ｼ・, _
        "閹晄釜繧鯉ｼ・nee buckling・・, _
        "閧｡OA縺ｮ逍ｼ逞帶ｧ豁ｩ陦・, _
        "菴灘ｹｹ蟾ｦ蜿ｳ謠ｺ繧鯉ｼ医Ζ繧ｳ繝薙・蠕ｴ蛟呎ｧ假ｼ・, _
        "蜑崎ｶｳ驛ｨ闕ｷ驥阪′縺励↓縺上＞", _
        "豁ｩ蟷・・縺ｰ繧峨▽縺・, _
        "髱ｴ縺ｮ蠑輔″縺壹ｊ", _
        "譚悶・豁ｩ陦悟勣縺ｸ縺ｮ蠑ｷ縺・ｾ晏ｭ・, _
        "迚・・闕ｷ驥榊屓驕ｿ", _
        "逍ｼ逞帛屓驕ｿ諤ｧ縺ｮ逡ｰ蟶ｸ豁ｩ螳ｹ" _
    )

    ' ----------------------
    ' D・壼鵠隱ｿ髫懷ｮｳ繝ｻ螟ｱ隱ｿ・・・・・・
    ' ----------------------
    D_items = Array( _
        "螟ｱ隱ｿ諤ｧ豁ｩ陦鯉ｼ医Ρ繧､繝峨・繝ｼ繧ｹ・・, _
        "蜊・ｳ･雜ｳ豁ｩ陦・, _
        "繧ｹ繝・ャ繝斐Φ繧ｰ豁ｩ陦・, _
        "縺弱￥縺励ｃ縺上＠縺滓ｭｩ陦・, _
        "譁ｹ蜷題ｻ｢謠帶凾縺ｮ螟ｧ縺阪↑謠ｺ繧・, _
        "荳願い縺ｨ縺ｮ蜊碑ｪｿ荳崎憶", _
        "縺ｵ繧峨▽縺榊､ｧ", _
        "雜ｳ縺ｮ菴咲ｽｮ豎ｺ繧√′荳肴ｭ｣遒ｺ" _
    )

    ' --------- 縺薙％縺九ｉ逕滓・蜃ｦ逅・---------

    Dim listSets As Variant
    listSets = Array(A_items, B_items, C_items, D_items)

    Dim idx As Long, arr As Variant

    For idx = 0 To 3
        ' 蟇ｾ雎｡繝輔Ξ繝ｼ繝蜿門ｾ・
        Set f = pg.controls(fNames(idx))
        If f Is Nothing Then GoTo ContinueNext
        
        ' 譌｢蟄倥メ繧ｧ繝・け蜑企勁
        For Each ctl In f.controls
            If TypeName(ctl) = "CheckBox" Then
                f.controls.Remove ctl.name
            End If
        Next ctl
        
                ' 霑ｽ蜉
        arr = listSets(idx)
        topPos = 24

        Dim maxBottom As Single
        maxBottom = 0

        For i = LBound(arr) To UBound(arr)
            Set chk = f.controls.Add("Forms.CheckBox.1", fNames(idx) & "_chk" & CStr(i), True)
            With chk
                .caption = arr(i)
                .Left = 12
                .Top = topPos
                .Width = f.Width - 24
                .Height = 18
            End With

            ' 縺薙・繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ縺ｮ荳狗ｫｯ繧定ｨ倬鹸
            If chk.Top + chk.Height > maxBottom Then
                maxBottom = chk.Top + chk.Height
            End If

            topPos = topPos + 20
        Next i
        ' 荳ｭ霄ｫ縺ｫ蜷医ｏ縺帙※繝輔Ξ繝ｼ繝鬮倥＆繧呈怙菴朱剞縺ｾ縺ｧ莨ｸ縺ｰ縺呻ｼ医Ο繧ｰ莉倥″・・
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
    ' 豁ｩ陦・閾ｪ遶句ｺｦ繧ｿ繝厄ｼ郁ｷ晞屬繝ｻ螻句､悶・螳牙ｮ壽ｧ繝ｻ騾溷ｺｦ・・
    BuildWalkIndep_DistanceOutdoor
    
    ' 縲檎焚蟶ｸ豁ｩ陦後阪ち繝厄ｼ倶ｸｭ霄ｫ・・蛻・｡槭ヵ繝ｬ繝ｼ繝・九メ繧ｧ繝・け鄒､・・
    BuildWalk_AbnormalTab
    BuildWalkAbnormal_Frames
    BuildWalkAbnormal_Checks
    FixWalkRootFrameHeight
End Sub



Public Sub BuildCogMentalUI_Simple()
    Dim f As MSForms.Frame
    Dim c As MSForms.Control
    Dim mp As MSForms.MultiPage
    Dim fw As MSForms.Frame   '笘・豁ｩ陦後ち繝悶・繝輔Ξ繝ｼ繝

    ' 隕ｪ繝輔Ξ繝ｼ繝・郁ｪ咲衍讖溯・繝ｻ邊ｾ逾樣擇・・
    On Error Resume Next
    Set f = GetCogRootFrame()
    On Error GoTo 0

    If f Is Nothing Then
        MsgBox "隱咲衍讖溯・繝輔Ξ繝ｼ繝・・age7 > Frame7 > Frame30・峨′隕九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If

    ' 笘・ｭｩ陦後ち繝・Frame6)縺ｨ蜷後§菴咲ｽｮ繝ｻ繧ｵ繧､繧ｺ縺ｫ蜷医ｏ縺帙ｋ
    On Error Resume Next
    Set fw = SafeGetControl(Me, "Frame6")
    On Error GoTo 0
    If fw Is Nothing Then Exit Sub
    If Not fw Is Nothing Then
        f.Left = fw.Left
        f.Top = fw.Top
        f.Width = fw.Width
        f.Height = fw.Height
    End If

    ' 縺・▲縺溘ｓ荳ｭ霄ｫ繧貞・驛ｨ繧ｯ繝ｪ繧｢・亥・縺ｮ繝ｩ繝吶Ν・上さ繝ｳ繝懶ｼ乗里蟄倥・繝ｫ繝√・繝ｼ繧ｸ繧ょ性繧√※・・
    Do While f.controls.count > 0
        f.controls.Remove f.controls(0).name
    Loop

    ' 蟄信ultiPage繧定ｿｽ蜉・郁ｪ咲衍讖溯・ / 邊ｾ逾樣擇 縺ｮ2繧ｿ繝悶・縺ｿ・・
    Set mp = f.controls.Add("Forms.MultiPage.1", "mpCogMental", True)
           With mp
        .Left = 6
        .Top = 0          ' 竊・縺薙％繧・6 竊・0 縺ｫ
        .Width = f.Width - 12
        .Height = f.Height - 12
        .Style = fmTabStyleTabs
        .TabOrientation = fmTabOrientationTop
    End With



    ' 譌｢蟄倥・繝ｼ繧ｸ繧貞・驛ｨ豸医＠縺ｦ縺九ｉ2繝壹・繧ｸ菴懈・
    Do While mp.Pages.count > 0
        mp.Pages.Remove 0
    Loop

    mp.Pages.Add
    mp.Pages.Add

    mp.Pages(0).caption = "隱咲衍讖溯・"
    mp.Pages(0).name = "pgCognition"

    mp.Pages(1).caption = "邊ｾ逾樣擇"
    mp.Pages(1).name = "pgMental"
End Sub




Public Sub BuildCog_CognitionCore()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim c As MSForms.Control
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    
    ' 隕ｪ繝輔Ξ繝ｼ繝・郁ｪ咲衍・・
    On Error Resume Next
    Set f = GetCogRootFrame()
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "繝輔か繝ｼ繝縺後≠繧翫∪縺帙ｓ", vbExclamation
        Exit Sub
    End If
    
    ' 蟄舌・繝ｫ繝√・繝ｼ繧ｸ
    Set mp = SafeGetControl(f, "mpCogMental")
    If mp Is Nothing Then
        MsgBox "mpCogMental 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 隱咲衍讖溯・繝壹・繧ｸ
    Set pg = SafeGetPage(mp, "pgCognition")
    If pg Is Nothing Then
        MsgBox "pgCognition 繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 譌｢蟄倥さ繝ｳ繝医Ο繝ｼ繝ｫ繧偵け繝ｪ繧｢・医ｄ繧顔峩縺礼畑・・
    Do While pg.controls.count > 0
        pg.controls.Remove pg.controls(0).name
    Loop
    
    ' 繝ｬ繧､繧｢繧ｦ繝亥渕貅・
    Dim rowTop As Single, rowGap As Single
    Dim col1Left As Single, col2Left As Single
    Dim lblW As Single, cmbW As Single
    
    rowTop = 18
    rowGap = 24
    col1Left = 12
    col2Left = 260
    lblW = 60
    cmbW = 140
    
    ' 蜈ｱ騾壹〒菴ｿ縺・ｩ穂ｾ｡繝ｪ繧ｹ繝・
    Dim i As Long
    Dim arr4()
    
    arr4 = Array("豁｣蟶ｸ", "繧・ｄ菴惹ｸ・, "菴惹ｸ・, "闡玲・縺ｫ菴惹ｸ・)
    
    '窶補・1陦檎岼・夊ｨ俶・・乗ｳｨ諢・窶補・
    ' 險俶・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogMemory", True)
    With lbl
        .caption = "險俶・"
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogMemory", True)
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
    
    ' 豕ｨ諢・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogAttention", True)
    With lbl
        .caption = "豕ｨ諢・
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogAttention", True)
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
    
    '窶補・2陦檎岼・夊ｦ句ｽ楢ｭ假ｼ丞愛譁ｭ 窶補・
    rowTop = rowTop + rowGap
    
    ' 隕句ｽ楢ｭ・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogOrientation", True)
    With lbl
        .caption = "隕句ｽ楢ｭ・
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogOrientation", True)
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
    
    ' 蛻､譁ｭ
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogJudgement", True)
    With lbl
        .caption = "蛻､譁ｭ"
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogJudgement", True)
    With cmb
        .Left = col2Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "濶ｯ螂ｽ"
        .AddItem "繧・ｄ荳榊ｮ牙ｮ・
        .AddItem "荳榊ｮ牙ｮ・
    End With
    
    '窶補・3陦檎岼・夐≠陦鯉ｼ剰ｨ隱・窶補・
    rowTop = rowTop + rowGap
    
    ' 驕り｡・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogExecutive", True)
    With lbl
        .caption = "驕り｡・
        .Left = col1Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogExecutive", True)
    With cmb
        .Left = col1Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "濶ｯ螂ｽ"
        .AddItem "繧・ｄ荳榊ｮ牙ｮ・
        .AddItem "荳榊ｮ牙ｮ・
    End With
    
    ' 險隱・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblCogLanguage", True)
    With lbl
        .caption = "險隱・
        .Left = col2Left
        .Top = rowTop
        .Width = lblW
        .Height = 18
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbCogLanguage", True)
    With cmb
        .Left = col2Left + lblW + 6
        .Top = rowTop - 2
        .Width = cmbW
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "蝠城｡後↑縺・
        .AddItem "繧・ｄ髫懷ｮｳ"
        .AddItem "髫懷ｮｳ鬘戊送"
    End With
End Sub



Public Sub BuildCog_DementiaBlock()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim fraTop As Single
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    Dim txt As MSForms.TextBox
    
    ' 隕ｪ繝輔Ξ繝ｼ繝
    On Error Resume Next
    Set f = GetCogRootFrame()
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "繝輔か繝ｼ繝縺後≠繧翫∪縺帙ｓ", vbExclamation
        Exit Sub
    End If
    
    ' 蟄舌・繝ｫ繝√・繝ｼ繧ｸ
    Set mp = SafeGetControl(f, "mpCogMental")
    If mp Is Nothing Then
        MsgBox "mpCogMental 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 隱咲衍讖溯・繝壹・繧ｸ
    Set pg = SafeGetPage(mp, "pgCognition")
    If pg Is Nothing Then
        MsgBox "pgCognition 繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
       ' 縺・▲縺溘ｓ縲∵里蟄倥・隱咲衍逞・ヶ繝ｭ繝・け繧呈ｶ医☆・医ｄ繧顔峩縺礼畑・・
    Dim i As Long
    For i = pg.controls.count - 1 To 0 Step -1
        With pg.controls(i)
            If .name = "lblDementiaType" _
               Or .name = "cmbDementiaType" _
               Or .name = "lblDementiaNote" _
               Or .name = "txtDementiaNote" Then
                pg.controls.Remove .name
            End If
        End With
    Next i

    
    ' 荳翫・隱咲衍6鬆・岼繝悶Ο繝・け縺ｮ縺吶＄荳九↓驟咲ｽｮ・医□縺・◆縺・陦鯉ｼ倶ｽ咏區縺ｶ繧謎ｸ九￡繧具ｼ・
    fraTop = 18 + 3 * 24 + 18   ' 18(譛蛻・ + 3陦・24 + 菴咏區
    
    ' 隕句・縺励Λ繝吶Ν縲瑚ｪ咲衍逞・・遞ｮ鬘槭・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblDementiaType", True)
    With lbl
        .caption = "隱咲衍逞・・遞ｮ鬘・
        .Left = 12
        .Top = fraTop
        .Width = 90
        .Height = 18
    End With
    
    ' 險ｺ譁ｭ蜷阪さ繝ｳ繝・
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbDementiaType", True)
    With cmb
        .Left = lbl.Left + lbl.Width + 6
        .Top = fraTop - 2
        .Width = 160
        .Height = 18
        .Style = fmStyleDropDownList
        .AddItem "縺ｪ縺・/ 荳肴・"
        .AddItem "繧｢繝ｫ繝・ワ繧､繝槭・蝙・
        .AddItem "陦邂｡諤ｧ"
        .AddItem "繝ｬ繝薙・蟆丈ｽ灘梛"
        .AddItem "蜑埼ｭ蛛ｴ鬆ｭ蝙・FTD)"
        .AddItem "豺ｷ蜷亥梛"
        .AddItem "縺昴・莉・
    End With
    
    ' 蛯呵・Λ繝吶Ν
    Set lbl = pg.controls.Add("Forms.Label.1", "lblDementiaNote", True)
    With lbl
        .caption = "蛯呵・
        .Left = cmb.Left + cmb.Width + 12
        .Top = fraTop
        .Width = 40
        .Height = 18
    End With
    
    ' 蛯呵・ユ繧ｭ繧ｹ繝・
    Set txt = pg.controls.Add("Forms.TextBox.1", "txtDementiaNote", True)
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = fraTop - 2
        .Width = mp.Width - .Left - 12
        .Height = 54
        .multiline = True
        .IMEMode = fmIMEModeHiragana
    End With
End Sub



Public Sub BuildCog_BPSD()
    Dim f As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim topY As Single
    Dim lbl As MSForms.label
    Dim chk As MSForms.CheckBox
    Dim i As Long
    
    ' 隕ｪ繝輔Ξ繝ｼ繝
    On Error Resume Next
    Set f = GetCogRootFrame()
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "繝輔か繝ｼ繝縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ", vbExclamation
        Exit Sub
    End If
    
    ' 蟄舌・繝ｫ繝√・繝ｼ繧ｸ
    Set mp = SafeGetControl(f, "mpCogMental")
    If mp Is Nothing Then
        MsgBox "mpCogMental 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 隱咲衍讖溯・繝壹・繧ｸ
    Set pg = SafeGetPage(mp, "pgCognition")
    If pg Is Nothing Then
        MsgBox "pgCognition 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' --- 譌｢蟄錬PSD繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ蜑企勁 ---
    Dim c As MSForms.Control
    For i = pg.controls.count - 1 To 0 Step -1
        If TypeName(pg.controls(i)) = "CheckBox" _
           Or pg.controls(i).name Like "lblBPSD*" Then
            pg.controls.Remove pg.controls(i).name
        End If
    Next i
    
    ' --- 霑ｽ蜉菴咲ｽｮ・郁ｪ咲衍逞・・遞ｮ鬘槭ヶ繝ｭ繝・け縺ｮ荳具ｼ・---
    topY = 18 + 3 * 24 + 18 + 24   '6鬆・岼繝悶Ο繝・け + 菴咏區
    topY = topY + 24               '隱咲衍逞・ｨｮ鬘槭・陦・
    
    ' 隕句・縺・
    Set lbl = pg.controls.Add("Forms.Label.1", "lblBPSD_Title", True)
    With lbl
        .caption = "隱咲衍逞・・蜻ｨ霎ｺ逞・憾・・PSD・・
        .Left = 12
        .Top = topY
        .Width = 180
        .Height = 18
    End With
    
    topY = topY + 24
    
    ' BPSD鬆・岼
    Dim items
    items = Array("謚代≧縺､", "荳榊ｮ・, "辟ｦ辯･", "蟷ｻ隕・, "螯・Φ", _
                  "蠕伜ｾ・, "證ｴ險", "證ｴ蜉・, "荳咲ｩ・, "逹｡逵髫懷ｮｳ", "譏ｼ螟憺・ｻ｢")
    
    Dim col As Long, row As Long
    col = 0: row = 0
    
    For i = LBound(items) To UBound(items)
        Set chk = pg.controls.Add("Forms.CheckBox.1", "chkBPSD" & CStr(i), True)
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
    Dim pg As MSForms.page
    Dim lbl As MSForms.label
    Dim cmb As MSForms.ComboBox
    Dim txt As MSForms.TextBox
    Dim topY As Single
    
    ' 隕ｪ繝輔Ξ繝ｼ繝・・rame31・・
    On Error Resume Next
    Set f = GetCogRootFrame()
    On Error GoTo 0
    If f Is Nothing Then
        MsgBox "繝輔か繝ｼ繝縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ", vbExclamation
        Exit Sub
    End If
    
    ' 蟄舌・繝ｫ繝√・繝ｼ繧ｸ mpCogMental
    Set mp = SafeGetControl(f, "mpCogMental")
    If mp Is Nothing Then
        MsgBox "mpCogMental 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 邊ｾ逾樣擇繝壹・繧ｸ
    Set pg = SafeGetPage(mp, "pgMental")
    If pg Is Nothing Then
        MsgBox "pgMental 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    ' 譌｢蟄倥け繝ｪ繧｢・医ｄ繧顔峩縺礼畑・・
    Dim i As Long
    For i = pg.controls.count - 1 To 0 Step -1
        pg.controls.Remove pg.controls(i).name
    Next i
    
    ' 繝ｬ繧､繧｢繧ｦ繝・
    Dim rowGap As Single: rowGap = 26
    Dim lblW As Single: lblW = 90
    Dim cmbW As Single: cmbW = 150
    Dim left1 As Single: left1 = 12
    Dim left2 As Single: left2 = 260
    topY = 18
    
    ' --- 豌怜・ ---
    Set lbl = pg.controls.Add("Forms.Label.1", "lblMood", True)
    With lbl
        .caption = "豌怜・"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbMood", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "螳牙ｮ・
        .AddItem "繧・ｄ荳榊ｮ牙ｮ・
        .AddItem "荳榊ｮ牙ｮ・
    End With
    
    ' --- 諢乗ｬｲ ---
    Set lbl = pg.controls.Add("Forms.Label.1", "lblMotivation", True)
    With lbl
        .caption = "諢乗ｬｲ"
        .Left = left2
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbMotivation", True)
    With cmb
        .Left = left2 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "鬮倥＞"
        .AddItem "譎ｮ騾・
        .AddItem "菴弱＞"
        .AddItem "縺ｻ縺ｨ繧薙←縺ｪ縺・
    End With
    
    ' --- 荳榊ｮ・---
    topY = topY + rowGap
    
    Set lbl = pg.controls.Add("Forms.Label.1", "lblAnxiety", True)
    With lbl
        .caption = "荳榊ｮ・
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbAnxiety", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "縺ｪ縺・
        .AddItem "霆ｽ蠎ｦ"
        .AddItem "荳ｭ遲牙ｺｦ"
        .AddItem "蠑ｷ縺・
    End With
    
    ' --- 蟇ｾ莠ｺ ---
    Set lbl = pg.controls.Add("Forms.Label.1", "lblRelation", True)
    With lbl
        .caption = "蟇ｾ莠ｺ髢｢菫・
        .Left = left2
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbRelation", True)
    With cmb
        .Left = left2 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "濶ｯ螂ｽ"
        .AddItem "縺翫♀繧縺ｭ濶ｯ螂ｽ"
        .AddItem "繧・ｄ蝠城｡・
        .AddItem "蝠城｡後≠繧・
    End With
    
    ' --- 逹｡逵 ---
    topY = topY + rowGap
    
    Set lbl = pg.controls.Add("Forms.Label.1", "lblSleep", True)
    With lbl
        .caption = "逹｡逵"
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set cmb = pg.controls.Add("Forms.ComboBox.1", "cmbSleep", True)
    With cmb
        .Left = left1 + lblW + 6
        .Top = topY - 2
        .Width = cmbW
        .Style = fmStyleDropDownList
        .AddItem "濶ｯ螂ｽ"
        .AddItem "蜈･逵蝗ｰ髮｣"
        .AddItem "荳ｭ騾碑ｦ夐・"
        .AddItem "譌ｩ譛晁ｦ夐・"
        .AddItem "譌･荳ｭ蛯ｾ逵"
    End With
    
    ' --- 蛯呵・---
    topY = topY + rowGap + 8
    
    Set lbl = pg.controls.Add("Forms.Label.1", "lblMentalNote", True)
    With lbl
        .caption = "蛯呵・
        .Left = left1
        .Top = topY
        .Width = lblW
    End With
    
    Set txt = pg.controls.Add("Forms.TextBox.1", "txtMentalNote", True)
    With txt
        .Left = left1 + lblW + 6
        .Top = topY - 2
         .Width = mp.Width - .Left - 12
        .Height = 50
        .IMEMode = fmIMEModeHiragana
        .multiline = True
        .EnterKeyBehavior = True
    End With
End Sub





Private Sub BuildDailyLogTab()
    Dim mp As Object
    Dim pg As MSForms.page
    Dim exists As Boolean
    Dim fra As MSForms.Frame

    On Error GoTo EH

    '=== MultiPage1 繧貞叙蠕・===
    Set mp = EvalCtl("MultiPage1")

    '=== 縺吶〒縺ｫ縲梧律縲・・險倬鹸縲阪ち繝悶′縺ゅｋ縺狗｢ｺ隱搾ｼ亥・遲臥畑・・==
    Set pg = SafeGetPage(mp, "譌･縲・・險倬鹸")
    exists = Not (pg Is Nothing)

    '=== 縺ｪ縺代ｌ縺ｰ譁ｰ縺励＞繝壹・繧ｸ繧定ｿｽ蜉 ===
    If Not exists Then
        Set pg = mp.Pages.Add
        pg.caption = "譌･縲・・險倬鹸"
    End If

    '=== 繝輔Ξ繝ｼ繝縺檎┌縺代ｌ縺ｰ1蛟九□縺台ｽ懊ｋ ===
    On Error Resume Next
    Set fra = SafeGetControl(pg, "fraDailyLog")
    On Error GoTo EH

    If fra Is Nothing Then
        Set fra = pg.controls.Add("Forms.Frame.1", "fraDailyLog")
        fra.caption = "譌･縲・・險倬鹸"
        fra.Left = 6
    fra.Top = 6
    fra.Width = mp.Width - 24      ' 竊・MultiPage 縺ｮ蟷・°繧牙ｷｦ蜿ｳ12pt縺壹▽菴咏區
    fra.Height = mp.Height - 30    ' 竊・荳贋ｸ九・繧ｿ繝厄ｼ倶ｽ咏區縺ｶ繧薙ｒ蟾ｮ縺怜ｼ輔＞縺ｦ譫縺・▲縺ｱ縺・
    End If

       BuildDailyLogLayout
       BuildDailyLog_StaffAndNote

    Exit Sub

EH:


   


End Sub



Private Sub BuildDailyLogLayout()
    On Error GoTo EH

    Dim f As Object
    Dim lbl As Object
    Dim txt As Object
    Dim colGap As Single
    Dim rowGap As Single
    Dim leftMargin As Single
    Dim topStart As Single
    Dim colW As Single
    Dim boxH As Single
    Dim rightLeft As Single
    Dim topLabelW As Single
    Dim topInputW As Single
    Dim secondRowTop As Single

    Set f = GetDailyLogFrame()
    If f Is Nothing Then GoTo ExitHere

    leftMargin = 12
    colGap = 12
    rowGap = 0
    topStart = 48
    colW = (f.Width - leftMargin * 2 - colGap) / 2
    If colW < 120 Then colW = 120
    rightLeft = leftMargin + colW + colGap
    topLabelW = 42
    topInputW = 88
    
    '=== 險倬鹸譌･繝ｩ繝吶Ν ===
    On Error Resume Next
    Set lbl = f.controls("lblDailyDate")
    On Error GoTo EH

    If lbl Is Nothing Then Set lbl = f.controls.Add("Forms.Label.1", "lblDailyDate")
    With lbl
        .caption = "險倬鹸譌･"
        .Left = leftMargin
        .Top = 18
        .Width = topLabelW
        .Height = 18
    End With

    '=== 險倬鹸譌･繝・く繧ｹ繝・===
    On Error Resume Next
    Set txt = SafeGetControl(f, "txtDailyDate")
    On Error GoTo EH
    If txt Is Nothing Then Set txt = f.controls.Add("Forms.TextBox.1", "txtDailyDate")
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = lbl.Top - 2
        .Width = topInputW
        .Height = 18
    End With

    '=== 險倬鹸閠・Λ繝吶Ν ===
    On Error Resume Next

    Set lbl = f.controls("lblDailyStaff")
    On Error GoTo EH
    If lbl Is Nothing Then Set lbl = f.controls.Add("Forms.Label.1", "lblDailyStaff")
    With lbl
        .caption = "險倬鹸閠・
        .Left = txt.Left + txt.Width + 24
        .Top = 18
        .Width = 40
        .Height = 18
    End With

    '=== 險倬鹸閠・ユ繧ｭ繧ｹ繝・===
    On Error Resume Next
    Set txt = SafeGetControl(f, "txtDailyStaff")
    On Error GoTo EH
    If txt Is Nothing Then Set txt = f.controls.Add("Forms.TextBox.1", "txtDailyStaff")
    With txt
        .Left = lbl.Left + lbl.Width + 6
        .Top = lbl.Top - 2
        .Width = 100
        .Height = 18
    End With

    boxH = 95
    secondRowTop = topStart + 18 + boxH + rowGap - 6

    CreateDailyField f, "lblDailyTraining", "txtDailyTraining", "縲仙ｮ滓命蜀・ｮｹ縲・, leftMargin, topStart, colW, boxH
    CreateDailyField f, "lblDailyReaction", "txtDailyReaction", "縲仙茜逕ｨ閠・・蜿榊ｿ懊・, rightLeft, topStart, colW, boxH
    CreateDailyField f, "lblDailyAbnormal", "txtDailyAbnormal", "縲千焚蟶ｸ謇隕九・, leftMargin, secondRowTop, colW, boxH
    CreateDailyField f, "lblDailyPlan", "txtDailyPlan", "縲蝉ｻ雁ｾ後・譁ｹ驥昴・, rightLeft, secondRowTop, colW, boxH


    '=== 險倬鹸蜀・ｮｹ繝・く繧ｹ繝茨ｼ医・繝ｫ繝√Λ繧､繝ｳ・・===
    On Error Resume Next
    f.controls.Remove "lblDailyNote"
    f.controls.Remove "txtDailyNote"
    On Error GoTo EH
 
ExitHere:
    Exit Sub

EH:

    Resume ExitHere
End Sub


Private Sub BuildDailyLog_StaffAndNote()
    BuildDailyLogLayout
End Sub

Private Sub CreateDailyField(ByVal f As Object, ByVal lblName As String, ByVal txtName As String, ByVal caption As String, _
                            ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim lbl As Object
    Dim txt As Object

    On Error Resume Next
    Set lbl = f.controls(lblName)
    On Error GoTo 0
    If lbl Is Nothing Then Set lbl = f.controls.Add("Forms.Label.1", lblName)

    With lbl
        .caption = caption
        .Left = x
        .Top = y
        .Width = w
        .Height = 16
    End With

    On Error Resume Next
    Set txt = f.controls(txtName)
    On Error GoTo 0
    If txt Is Nothing Then Set txt = f.controls.Add("Forms.TextBox.1", txtName)

    With txt
        .Left = x
        .Top = y + 18
        .Width = w
        .Height = h
        .multiline = True
        .EnterKeyBehavior = True
        .ScrollBars = 2   ' fmScrollBarsVertical
    End With

End Sub

Public Sub BuildDailyLog_HistoryList(owner As Object)
    Dim f As Object
    Dim txtLeft As Object
    Dim txtRight As Object
    Dim lst As MSForms.ListBox
    Dim topPos As Single
    Dim margin As Single
    Dim fieldsBottom As Single

    margin = 12

    Set f = GetDailyLogFrame()
    If f Is Nothing Then Set f = SafeGetControl(owner, "fraDailyLog")
    If f Is Nothing Then Exit Sub

    Set txtLeft = SafeGetControl(f, "txtDailyAbnormal")
    Set txtRight = SafeGetControl(f, "txtDailyPlan")
    If txtLeft Is Nothing Or txtRight Is Nothing Then
        BuildDailyLogLayout
        Set txtLeft = SafeGetControl(f, "txtDailyAbnormal")
        Set txtRight = SafeGetControl(f, "txtDailyPlan")
    End If
    If txtLeft Is Nothing Or txtRight Is Nothing Then Exit Sub

    fieldsBottom = txtLeft.Top + txtLeft.Height
    If txtRight.Top + txtRight.Height > fieldsBottom Then
        fieldsBottom = txtRight.Top + txtRight.Height
    End If
    
    On Error Resume Next
    f.controls.Remove "lstDailyLogList"
    f.controls.Remove "lblDailyHistory"
    f.controls.Remove "lblDailyMonitoringCreate"
    On Error GoTo 0

'--- 螻･豁ｴ繝ｩ繝吶Ν菴懈・ ---
Dim lbl As MSForms.label
Dim lblMonitoring As MSForms.label

Set lblMonitoring = f.controls.Add("Forms.Label.1", "lblDailyMonitoringCreate", True)
With lblMonitoring
    .caption = "繝｢繝九ち繝ｪ繝ｳ繧ｰ譛ｬ譁・
    .Left = margin
    .Top = fieldsBottom - 4
    .Width = 200
    .Height = 18
    .Font.Bold = True
End With

    ' ListBox 霑ｽ蜉
    Set lst = f.controls.Add("Forms.ListBox.1", "lstDailyLogList", True)

  
    With lst
        .Left = margin
        .Top = topPos
        .Width = f.Width - margin * 2
        .Height = f.Height - .Top - 8
        .ColumnCount = 3          ' 險倬鹸蟷ｴ譛・/ 蜷榊燕 / 險倬鹸蜀・ｮｹ
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

    ' fraDailyLog 縺ｨ 險倬鹸閠・ユ繧ｭ繧ｹ繝医ｒ蜿門ｾ・
    Set f = GetDailyLogFrame()
    If f Is Nothing Then Set f = SafeGetControl(owner, "fraDailyLog")
    If f Is Nothing Then Exit Sub
    Set txtStaff = SafeGetControl(f, "txtDailyStaff")
    If txtStaff Is Nothing Then
        BuildDailyLogLayout
        Set txtStaff = SafeGetControl(f, "txtDailyStaff")
    End If
    If txtStaff Is Nothing Then Exit Sub
    BuildDailyLog_HistoryList owner   ' 笘・％繧後ｒ霑ｽ蜉・・istBox繧貞ｿ・★菴懊ｋ・・

    ' 譌｢縺ｫ繝懊ち繝ｳ縺後≠繧後・蜑企勁縺励※菴懊ｊ逶ｴ縺暦ｼ亥・遲会ｼ・
    On Error Resume Next
    f.controls.Remove "cmdDailyExtract"
    On Error GoTo 0

    ' 謚ｽ蜃ｺ繝懊ち繝ｳ霑ｽ蜉
    Set cmd = f.controls.Add("Forms.CommandButton.1", "cmdDailyExtract", True)

    With cmd
        .caption = "繝｢繝九ち繝ｪ繝ｳ繧ｰ菴懈・"
        .Width = 120
        .Height = 24
        .Top = txtStaff.Top
        .Left = f.Width - margin - .Width

    End With
    
     Set mDailyExtract = cmd


End Sub

Public Sub BuildDailyLog_SaveButton(owner As Object)
    Dim f As Object
    Dim txtStaff As Object
    Dim cmd As MSForms.CommandButton
    Dim cmdExtract As Object
    Dim margin As Single
    Dim rightGap As Single
    Dim extractW As Single

    margin = 12
    rightGap = 8
    extractW = 120

    ' fraDailyLog 縺ｨ 險倬鹸閠・ユ繧ｭ繧ｹ繝医ｒ蜿門ｾ・
    Set f = GetDailyLogFrame()
    If f Is Nothing Then Set f = SafeGetControl(owner, "fraDailyLog")
    If f Is Nothing Then Exit Sub
    Set txtStaff = SafeGetControl(f, "txtDailyStaff")
    If txtStaff Is Nothing Then
        BuildDailyLogLayout
        Set txtStaff = SafeGetControl(f, "txtDailyStaff")
    End If
    If txtStaff Is Nothing Then Exit Sub

    ' 譌｢縺ｫ繝懊ち繝ｳ縺後≠繧後・蜑企勁縺励※菴懊ｊ逶ｴ縺暦ｼ亥・遲会ｼ・
    On Error Resume Next
    f.controls.Remove "cmdDailySave"
    On Error GoTo 0

    ' 菫晏ｭ倥・繧ｿ繝ｳ霑ｽ蜉
    Set cmd = f.controls.Add("Forms.CommandButton.1", "cmdDailySave", True)

    With cmd
        .caption = "譌･縲・・險倬鹸繧剃ｿ晏ｭ・
        .Width = 110
        .Height = 24
        .Top = txtStaff.Top
        Set cmdExtract = SafeGetControl(f, "cmdDailyExtract")
        If cmdExtract Is Nothing Then
            .Left = f.Width - margin - extractW - rightGap - .Width
        Else
            .Left = cmdExtract.Left - rightGap - .Width
        End If
    End With
  
    Set mDailySave = cmd


End Sub



Private Sub mDailyExtract_Click()
    ' 竭 譚先侭譁・ｒ菴懊ｋ
    Call Me.BuildMonthlyDraft_FromDailyLog
    
    
    Dim box As Object
            Set box = DailyLogCtl("txtMonthlyMonitoringDraft")
    If box Is Nothing Then Exit Sub


        If InStr(1, box.value, "・医％縺ｮ譛医・險倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ・・, vbTextCompare) > 0 Then
            
            box.value = "縲先怦谺｡繝｢繝九ち繝ｪ繝ｳ繧ｰ荳区嶌縺阪・ & vbCrLf & _
            "蟇ｾ雎｡・・ & Me.controls("frHeader").controls("txtHdrName").value & vbCrLf & _
                 "譛滄俣・・ & Format$(DateSerial(Year(CDate(DailyLogCtl("txtDailyDate").value)), _
                                      Month(CDate(DailyLogCtl("txtDailyDate").value)), 1), "yyyy/mm/dd") & _
            " - " & _
            Format$(DateSerial(Year(CDate(DailyLogCtl("txtDailyDate").value)), _
                                Month(CDate(DailyLogCtl("txtDailyDate").value)) + 1, 0), "yyyy/mm/dd") & vbCrLf & vbCrLf & _
            "笆 縺薙・譛医↓險倬鹸縺輔ｌ縺溽音險倅ｺ矩・ & vbCrLf & _
            "縺薙・譛医・迚ｹ險倅ｺ矩・→縺ｪ繧玖ｨ倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ縺ｧ縺励◆縲・ & vbCrLf & _
            "菴楢ｪｿ髱｢縺ｫ螟ｧ縺阪↑螟牙虚縺ｯ縺ｪ縺上∵律縲・・繝ｪ繝上ン繝ｪ縺ｫ繧ょｮ牙ｮ壹＠縺ｦ蜿悶ｊ邨・∪繧後※縺・∪縺励◆縲・ & vbCrLf & _
            "莉雁ｾ後ｂ迴ｾ蝨ｨ縺ｮ迥ｶ諷九ｒ邯ｭ謖√〒縺阪ｋ繧医≧縲∝ｼ輔″邯壹″邨碁℃繧定ｦｳ蟇溘＠縺ｦ縺・″縺ｾ縺吶・

                      Call ExportMonitoring_ToMonthlyWorkbook( _
          CDate(DailyLogCtl("txtDailyDate").value), _
                Me.controls("frHeader").controls("txtHdrName").value, _
                box.value)

           Exit Sub

        End If

    
    

    ' 竭｡ AI縺ｧ荳区嶌縺阪↓螟画鋤
' AI縺ｧ荳区嶌縺阪↓螟画鋤
If Trim$(DailyLogCtl("txtMonthlyMonitoringDraft").value) = "" Then
    MsgBox "繝｢繝九ち繝ｪ繝ｳ繧ｰ譛ｬ譁・′遨ｺ縺ｧ縺吶ょ・縺ｫ蜀・ｮｹ繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・, vbExclamation
    Exit Sub
End If



DailyLogCtl("txtMonthlyMonitoringDraft").value = _
    OpenAI_BuildDraft( _
            "縲仙・蜉帙ヵ繧ｩ繝ｼ繝槭ャ繝亥宍螳医・ & vbCrLf & _
"莉･荳九・隕句・縺励ｒ縲∬｡ｨ險倥・鬆・ｺ上・險伜捷・遺蔓・峨ｒ荳蛻・､峨∴縺壹↓蠢・★蜃ｺ蜉帙☆繧九％縺ｨ縲・ & vbCrLf & _
"隕句・縺励・霑ｽ蜉繝ｻ蜑企勁繝ｻ險縺・鋤縺育ｦ∵ｭ｢縲り｣・｣ｾ・遺・/縲舌・逡ｪ蜿ｷ莉倥￠・臥ｦ∵ｭ｢縲・ & vbCrLf & _
"蠢・★縺薙・鬆・ｺ擾ｼ・ & vbCrLf & _
"笆 縺薙・譛医↓險倬鹸縺輔ｌ縺溽音險倅ｺ矩・ & vbCrLf & _
"笆 繧ｳ繝｡繝ｳ繝医・閠・ｯ・ & vbCrLf & vbCrLf & _
"繝ｻ譛ｬ譁・ｼ育ｵ碁℃繝ｻ譎らｳｻ蛻暦ｼ峨↓縺ｯ縲∽ｺ句ｮ溘・縺ｿ繧定ｨ倩ｼ峨☆繧九りｨ倬鹸縺ｫ譖ｸ縺九ｌ縺ｦ縺・↑縺・ｺ句ｮ溘ｄ謗ｨ貂ｬ縺ｯ縲∵悽譁・↓縺ｯ蜷ｫ繧√↑縺・ゅ後さ繝｡繝ｳ繝医・閠・ｯ溘肴ｬ・↓髯舌ｊ縲∬ｨ倬鹸蜀・ｮｹ繧定ｸ上∪縺医◆莉雁ｾ後・隕ｳ蟇溯ｦ也せ繧・蕗諢冗せ繧定ｨ倩ｼ峨＠縺ｦ繧医＞縲ゅ◎縺ｮ髫帙・縲∵妙螳壹ｒ驕ｿ縺代√娯雷笳九・蜿ｯ閭ｽ諤ｧ縺後≠繧九阪娯雷笳九↓逡呎э縺励※邨碁℃繧堤｢ｺ隱阪☆繧九阪↑縺ｩ縺ｮ陦ｨ迴ｾ縺ｫ髯仙ｮ壹☆繧九ょ現蟄ｦ逧・愛譁ｭ縲∵隼蝟・・謔ｪ蛹悶・譁ｭ螳壹∝屏譫憺未菫ゅ・譁ｭ螳壹・陦後ｏ縺ｪ縺・よ枚菴薙・縲後〒縺吶・縺ｾ縺呵ｪｿ縲阪→縺励∫樟蝣ｴ險倬鹸縺ｨ縺励※閾ｪ辟ｶ縺ｧ隱ｭ縺ｿ繧・☆縺・沐繧峨°縺輔ｒ謖√◆縺帙ｋ縲・, _
            DailyLogCtl("txtMonthlyMonitoringDraft").value _
        )


            Call ExportMonitoring_ToMonthlyWorkbook( _
        CDate(DailyLogCtl("txtDailyDate").value), _
        Me.controls("frHeader").controls("txtHdrName").value, _
         DailyLogCtl("txtMonthlyMonitoringDraft").value)

        
End Sub

Private Sub mDailySave_Click()
    mDailyLogManual = True
    Call SaveDailyLog_Append(Me)
    mDailyLogManual = False
    MsgBox "譌･縲・・險倬鹸繧剃ｿ晏ｭ倥＠縺ｾ縺励◆縲・, vbInformation
End Sub





' 隧穂ｾ｡繝輔か繝ｼ繝荳矩Κ縺ｫ縲後す繝ｼ繝医∈菫晏ｭ倥阪・繧ｿ繝ｳ繧・縺､驟咲ｽｮ縺吶ｋ・・蝗槫ｮ溯｡檎畑・・
Public Sub PlaceGlobalSaveButton_Once()

    Dim btnClose As MSForms.CommandButton
    Dim btnSave As MSForms.CommandButton
    Dim c As MSForms.Control

    ' 縲碁哩縺倥ｋ縲阪・繧ｿ繝ｳ繧偵く繝｣繝励す繝ｧ繝ｳ縺ｧ迚ｹ螳・
    For Each c In Me.controls
        If TypeOf c Is MSForms.CommandButton Then
            If c.caption = "髢峨§繧・ Then
                Set btnClose = c
                Exit For
            End If
        End If
    Next c

    If btnClose Is Nothing Then
        MsgBox "髢峨§繧九・繧ｿ繝ｳ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If

    ' 譌｢縺ｫ繧ｰ繝ｭ繝ｼ繝舌Ν菫晏ｭ倥・繧ｿ繝ｳ縺後≠繧九°遒ｺ隱・
    On Error Resume Next
    Set btnSave = Me.controls("cmdSaveGlobal")
    On Error GoTo 0

    ' 縺ｪ縺代ｌ縺ｰ譁ｰ隕丈ｽ懈・
    If btnSave Is Nothing Then
        Set btnSave = Me.controls.Add("Forms.CommandButton.1", "cmdSaveGlobal")
        btnSave.caption = "繧ｷ繝ｼ繝医∈菫晏ｭ・
    End If

    ' 髢峨§繧九・繧ｿ繝ｳ縺ｨ鬮倥＆繝ｻ邵ｦ菴咲ｽｮ繧偵◎繧阪∴縺ｦ縲∝ｷｦ髫｣縺ｫ驟咲ｽｮ
    With btnSave
        .Height = btnClose.Height
        .Top = btnClose.Top
        .Width = btnClose.Width + 50
        .Left = btnClose.Left - .Width - 12
    End With

    ' ---- 繧ｯ繝ｪ繧｢繝懊ち繝ｳ・・mdClearGlobal・峨ｒ驟咲ｽｮ ----
Dim btnClear As MSForms.CommandButton

' 譌｢縺ｫ蟄伜惠縺吶ｋ縺狗｢ｺ隱・
On Error Resume Next
Set btnClear = Me.controls("cmdClearGlobal")
On Error GoTo 0

' 縺ｪ縺代ｌ縺ｰ譁ｰ隕丈ｽ懈・
If btnClear Is Nothing Then
    Set btnClear = Me.controls.Add("Forms.CommandButton.1", "cmdClearGlobal")
    btnClear.caption = "繧ｯ繝ｪ繧｢"
End If

' 菫晏ｭ倥・繧ｿ繝ｳ縺ｮ蜿ｳ髫｣縺ｫ驟咲ｽｮ
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

    
    
    
    ' ---- 繝懊ち繝ｳ謨ｴ蛻暦ｼ井ｿ晏ｭ・竊・繧ｯ繝ｪ繧｢ 竊・髢峨§繧具ｼ・----

' 髢峨§繧九・繧ｿ繝ｳ繧貞渕貅悶→縺励※荳逡ｪ蜿ｳ遶ｯ縺ｫ蝗ｺ螳・
btnClose.Left = btnClose.Left

' 菫晏ｭ倥・繧ｿ繝ｳ繧帝哩縺倥ｋ縺ｮ蟾ｦ縺ｫ
btnSave.Left = btnClose.Left - btnSave.Width - 12

' 繧ｯ繝ｪ繧｢繝懊ち繝ｳ・・mdClearGlobal・峨ｒ菫晏ｭ倥・蟾ｦ縺ｫ
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

    For Each c In Me.controls
        ' 繝・く繧ｹ繝医・繝・け繧ｹ縺ｯ遨ｺ縺ｫ
        If TypeOf c Is MSForms.TextBox Then
            ' 隧穂ｾ｡譌･(txtEDate)縺ｨ譌･縲・ｨ倬鹸縺ｮ譌･莉・txtDailyDate)縺ｯ繧ｯ繝ｪ繧｢縺励↑縺・
            If c.name <> "txtEDate" And c.name <> "txtDailyDate" Then
                c.value = ""
            End If
        End If

        ' 繧ｳ繝ｳ繝懊・繝・け繧ｹ縺ｯ驕ｸ謚櫁ｧ｣髯､
        If TypeOf c Is MSForms.ComboBox Then
            c.value = ""
        End If

        ' 繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ縺ｯ繧ｪ繝・
        If TypeOf c Is MSForms.CheckBox Then
            c.value = False
        End If

        ' 繝ｪ繧ｹ繝医・繝・け繧ｹ縺ｯ驕ｸ謚槭□縺題ｧ｣髯､・磯・岼縺ｯ谿九☆・・
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
    
    ' 蟇ｾ雎｡縺ｮ荳隕ｧListBox繧貞叙蠕・
    Set lb = Me.controls("lstDailyLogList")
    
    ' 蜈ｨ陦後・蜈ｨ蛻励ｒ繧ｿ繝門玄蛻・ｊ・区隼陦後〒騾｣邨・
    For r = 0 To lb.ListCount - 1
        For c = 0 To lb.ColumnCount - 1
            If c > 0 Then buf = buf & vbTab
            buf = buf & CStr(lb.List(r, c))
        Next c
        buf = buf & vbCrLf
    Next r
    
    ' 繧ｯ繝ｪ繝・・繝懊・繝峨∈繧ｳ繝斐・
    Dim dobj As New MSForms.DataObject
    dobj.SetText buf
    dobj.PutInClipboard
    
    MsgBox "縺薙・譛医・險倬鹸荳隕ｧ繧偵け繝ｪ繝・・繝懊・繝峨↓繧ｳ繝斐・縺励∪縺励◆縲・ & vbCrLf & _
           "繝｡繝｢蟶ｳ繧Цord縺ｫ Ctrl+V 縺ｧ雋ｼ繧贋ｻ倥￠縺ｧ縺阪∪縺吶・, vbInformation
End Sub



Public Sub HookDailyLogList(lb As MSForms.ListBox)
    ' 譌･縲・・險倬鹸荳隕ｧ ListBox 逕ｨ縺ｮ繧､繝吶Φ繝医ヵ繝・け
    If mDailyList Is Nothing Then
        Set mDailyList = New clsDailyLogList
    End If
    Set mDailyList.lb = lb
End Sub



'=== 譌･縲・・險倬鹸繝輔Ξ繝ｼ繝蜿門ｾ励・繝ｫ繝代・・亥・騾壼喧逕ｨ・・===
Private Function GetDailyLogFrame() As MSForms.Frame
    Dim mp As Object
    Dim pg As Object
    Dim f As Object

    On Error Resume Next

   Set mp = EvalCtl("MultiPage1")
    If mp Is Nothing Then
        Exit Function
    End If

    ' 縲梧律縲・・險倬鹸縲阪・繝ｼ繧ｸ繧呈爾縺・
    Set pg = SafeGetPage(mp, "譌･縲・・險倬鹸")
    If pg Is Nothing Then
        Exit Function
    End If

    ' 繝輔Ξ繝ｼ繝 fraDailyLog 繧貞叙蠕・
    Set f = SafeGetControl(pg, "fraDailyLog")
    If f Is Nothing Then
        Exit Function
    End If

    Set GetDailyLogFrame = f
End Function



Private Function GetMainMultiPage() As MSForms.MultiPage
    Dim c As Control
    
    On Error Resume Next
    Set GetMainMultiPage = Me.controls("MultiPage1")
    On Error GoTo 0
    If Not GetMainMultiPage Is Nothing Then Exit Function

    For Each c In Me.controls
        If TypeName(c) = "MultiPage" Then
            Set GetMainMultiPage = c
            Exit Function
        End If
    Next
End Function



'=== 豁ｩ陦瑚ｩ穂ｾ｡繝輔Ξ繝ｼ繝蜿門ｾ励・繝ｫ繝代・・・rame6 蝗ｺ螳夲ｼ・===
Private Function GetWalkFrame() As MSForms.Frame
    Set GetWalkFrame = GetWalkRootFrame()
End Function

Public Function GetWalkRootFrame() As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim c As Control
    Dim best As MSForms.Frame
    Dim bestArea As Double
    Dim area As Double

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Function

    For Each pg In mp.Pages
        For Each c In pg.controls
            If TypeName(c) = "Frame" Then
                If InStr(1, CStr(c.caption), "", vbTextCompare) > 0 _
                   Or InStr(1, CStr(c.name), "walk", vbTextCompare) > 0 Then
                    Set GetWalkRootFrame = c
                    Exit Function
                End If
                area = CDbl(c.Width) * CDbl(c.Height)
                If best Is Nothing Or area > bestArea Then
                    If InStr(1, CStr(pg.caption), "", vbTextCompare) > 0 Then
                        Set best = c
                        bestArea = area
                    End If
                End If
            End If
        Next
    Next

    Set GetWalkRootFrame = best
End Function

Public Function GetCogRootFrame() As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim host As MSForms.Frame
    Dim c As Control
    Dim firstFrame As MSForms.Frame
    Dim frame30 As MSForms.Frame

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Function

    Set host = EvalCtl("Frame7")
    If host Is Nothing Then Exit Function

    For Each c In host.controls
        If TypeName(c) = "Frame" Then
            If firstFrame Is Nothing Then Set firstFrame = c
            If StrComp(CStr(c.name), "Frame30", vbTextCompare) = 0 Then
                Set frame30 = c
                Exit For
            End If
        End If
    Next
    

    If Not frame30 Is Nothing Then
        Set GetCogRootFrame = frame30
    Else
        Set GetCogRootFrame = firstFrame
    End If
End Function

Public Function GetCogTabs() As MSForms.MultiPage
    Dim f As MSForms.Frame
    Dim c As Control

    Set f = GetCogRootFrame()
    If f Is Nothing Then Exit Function

    For Each c In f.controls
        If TypeName(c) = "MultiPage" Then
            If InStr(1, CStr(c.name), "Cog", vbTextCompare) > 0 Or InStr(1, CStr(c.name), "Mental", vbTextCompare) > 0 Then
                Set GetCogTabs = c
                Exit Function
            End If
        End If
    Next
    For Each c In f.controls
        If TypeName(c) = "MultiPage" Then
            Set GetCogTabs = c
            Exit Function
        End If
    Next
End Function




Private Sub FixWalkRootFrameHeight()
    Dim f As MSForms.Frame
    Dim c As Control
    Dim maxBottom As Single

    Set f = GetWalkFrame()
    If f Is Nothing Then Exit Sub

    ' 蟄舌さ繝ｳ繝医Ο繝ｼ繝ｫ縺ｮ荳逡ｪ荳九・菴咲ｽｮ繧定ｪｿ縺ｹ繧・
    For Each c In f.controls
        If c.Top + c.Height > maxBottom Then
            maxBottom = c.Top + c.Height
        End If
    Next c

    ' 蠢・ｦ√↑繧蛾ｫ倥＆繧剃ｼｸ縺ｰ縺・
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

    ' 蟋ｿ蜍｢隧穂ｾ｡繧ｿ繝・
    FitFrameHeightToChildren SafeGetControl(Me, "Frame2")

    ' 霄ｫ菴捺ｩ溯・隧穂ｾ｡繧ｿ繝厄ｼ郁ｦｪ縺縺題ｪｿ謨ｴ縲ょｭ色rame12縺ｫ縺ｯ隗ｦ繧峨↑縺・ｼ・
    FitFrameHeightToChildren SafeGetControl(Me, "Frame12")
    FitFrameHeightToChildren SafeGetControl(Me, "Frame3")
    ' FitFrameHeightToChildren Me.Controls("Frame14")

    ' 豁ｩ陦瑚ｩ穂ｾ｡繧ｿ繝厄ｼ亥､ｧ譫・・
    FitFrameHeightToChildren SafeGetControl(Me, "Frame6")

    ' 隱咲衍繝ｻ邊ｾ逾槭ち繝厄ｼ郁ｦｪFrame7縺縺代ｒ隱ｿ謨ｴ・・
    FitFrameHeightToChildren SafeGetControl(Me, "Frame7")
    
    On Error GoTo 0
End Sub










Private Sub GetPageUsableArea( _
    ByVal pageIndex As Long, _
    ByRef x As Single, _
    ByRef y As Single, _
    ByRef w As Single, _
    ByRef h As Single)

    Dim mp As MSForms.MultiPage

    x = 0: y = 0: w = 0: h = 0   ' 繝・ヵ繧ｩ繝ｫ繝医け繝ｪ繧｢

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Sub

    If pageIndex < 0 Then Exit Sub
    If pageIndex > mp.Pages.count - 1 Then Exit Sub

    ' 莉翫・ MultiPage 蜈ｨ菴薙ｒ縲後・繝ｼ繧ｸ縺ｮ蛻ｩ逕ｨ蜿ｯ閭ｽ鬆伜沺縲阪→縺励※霑斐☆
    ' ・井ｽ咏區繧・ち繝門・縺ｮ繝槭う繝翫せ縺ｯ縲∝ｾ後〒 AlignRootFrame 蛛ｴ縺ｧ隱ｿ謨ｴ縺吶ｋ・・
    x = 0
    y = 0
    w = mp.Width
    h = mp.Height
End Sub



Private Sub AlignRootFrameToPage(ByVal pageIndex As Long, root As MSForms.Frame)
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim pageLeft As Single, pageTop As Single
    Dim pageWidth As Single, pageHeight As Single

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Sub
    If root Is Nothing Then Exit Sub
    If pageIndex < 0 Or pageIndex > mp.Pages.count - 1 Then Exit Sub

    Set pg = mp.Pages(pageIndex)

    '=== 繝壹・繧ｸ縺ｮ繧ｯ繝ｩ繧､繧｢繝ｳ繝磯伜沺・医ち繝悶ｒ髯､縺・◆荳ｭ霄ｫ驛ｨ蛻・ｼ峨ｒ邂怜・ ===
    ' 縺薙％縺ｯ PREVIEW 縺ｧ蜃ｺ縺ｦ縺・◆蛟､縺ｨ蜷後§繝ｭ繧ｸ繝・け縺ｫ謠・∴繧句燕謠舌〒縲・
    ' 繧ｷ繝ｳ繝励Ν縺ｫ MultiPage 縺ｮ蜀・・繧剃ｽｿ縺・
    pageLeft = 0
    pageTop = 0
    pageWidth = mp.Width          ' 繧ｿ繝門ｷｦ蜿ｳ縺ｮ菴咏區縺ｯ縺ｻ縺ｼ 0 謇ｱ縺・
    pageHeight = mp.Height - 40   ' 荳句・繝懊ち繝ｳ縺ｶ繧灘ｰ代＠縺縺第而縺医ａ

    '=== 繝ｫ繝ｼ繝医ヵ繝ｬ繝ｼ繝繧偵・繝ｼ繧ｸ荳譚ｯ縺ｫ繝輔ぅ繝・ヨ ===
    With root
        .Left = pageLeft
        .Top = pageTop
        .Width = pageWidth
        .Height = pageHeight
    End With
End Sub




Private Sub PreviewOnePage(ByVal idx As Long, ByVal mp As MSForms.MultiPage)
    Dim pg As MSForms.page
    Dim root As MSForms.Frame
    Dim x As Single, y As Single, w As Single, h As Single

    Set pg = mp.Pages(idx)
    Set root = GetPageRootFrame(idx)

    Debug.Print "--- Page", idx, "[" & pg.caption & "] ---"

    If root Is Nothing Then
        Debug.Print "  RootFrame: <NOT FOUND>"
        Exit Sub
    End If

    ' 迴ｾ蝨ｨ蛟､
    Debug.Print "  Current:", _
                "L=" & root.Left, _
                "T=" & root.Top, _
                "W=" & root.Width, _
                "H=" & root.Height

    ' AlignRootFrameToPage 縺御ｽｿ縺・・繝ｼ繧ｸ鬆伜沺
    GetPageUsableArea idx, x, y, w, h
    Debug.Print "  PageArea:", _
                "X=" & x, "Y=" & y, _
                "W=" & w, "H=" & h

    ' 繧ゅ＠ AlignRootFrameToPage 繧貞他繧薙□繧峨％縺・↑繧具ｼ遺ｻ螳滄圀縺ｫ縺ｯ譖ｸ縺肴鋤縺医↑縺・ｼ・
    Debug.Print "  WouldAlignTo:", _
                "L=" & (x), _
                "T=" & (y), _
                "W=" & (w), _
                "H=" & (h)
End Sub


Private Function GetPageRootFrame(ByVal pageIndex As Long) As MSForms.Frame
    Dim mp As MSForms.MultiPage
    Dim pg As MSForms.page
    Dim c As Control
    Dim f As MSForms.Frame
    Dim best As MSForms.Frame
    Dim bestArea As Single
    Dim area As Single

    Set mp = GetMainMultiPage()
    If mp Is Nothing Then Exit Function

    If pageIndex < 0 Or pageIndex > mp.Pages.count - 1 Then Exit Function

    Set pg = mp.Pages(pageIndex)
    If InStr(1, CStr(pg.caption), "Fm", vbTextCompare) > 0 Then
        Set GetPageRootFrame = GetCogRootFrame()
        If Not GetPageRootFrame Is Nothing Then Exit Function
    End If

    ' 縺昴・繝壹・繧ｸ蜀・・縲御ｸ逡ｪ螟ｧ縺阪↑ Frame = 繝ｫ繝ｼ繝医ヵ繝ｬ繝ｼ繝縲阪→縺ｿ縺ｪ縺・
    For Each c In pg.controls
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

    For i = 0 To mp.Pages.count - 1
        Set root = GetPageRootFrame(i)
        If Not root Is Nothing Then
            AlignRootFrameToPage i, root
        End If
    Next i
End Sub


 
Public Sub TidyBaseLayout_Once()
    If mBaseLayoutDone Then Exit Sub
    mBaseLayoutDone = True

    '笘・蝓ｺ譛ｬ繝ｬ繧､繧｢繧ｦ繝茨ｼ医・繝ｼ繧ｸ蜈ｱ騾夲ｼ峨・縺薙％縺縺代〒繧・ｋ
    Apply_AlignRoot_All


End Sub




Public Sub SetFormHeightSafe(ByVal newH As Single)
    Me.Height = newH
    DoEvents
    Dim mp As Object
    Set mp = EvalCtl("MultiPage1")
    If Not mp Is Nothing Then mp.Height = Me.InsideHeight - 12
    If mBaseLayoutDone Then
        Apply_AlignRoot_All
    End If
End Sub




Public Sub AdjustBottomButtons()

    Dim yBtn As Single

    ' 繝懊ち繝ｳ縺後∪縺辟｡縺・ち繧､繝溘Φ繧ｰ縺ｧ縺ｯ菴輔ｂ縺励↑縺・
    If Not ControlExists(Me, "btnCloseCtl") Then Exit Sub
    If Not ControlExists(Me, "cmdSaveGlobal") Then Exit Sub
    If Not ControlExists(Me, "cmdClearGlobal") Then Exit Sub

    yBtn = Me.InsideHeight - Me.controls("btnCloseCtl").Height - 12

    Me.controls("btnCloseCtl").Top = yBtn
    Me.controls("cmdSaveGlobal").Top = yBtn
    Me.controls("cmdClearGlobal").Top = yBtn
    
    
     ' 笘・％縺難ｼ亥燕髱｢縺ｸ・・
    Me.controls("btnCloseCtl").ZOrder 0
    Me.controls("cmdSaveGlobal").ZOrder 0
    Me.controls("cmdClearGlobal").ZOrder 0
    
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


minH = 620   ' 竊・隧穂ｾ｡繝輔か繝ｼ繝縺ｨ縺励※譛菴朱剞谺ｲ縺励＞鬮倥＆・郁ｪｿ謨ｴ蜿ｯ・・

maxH = Application.UsableHeight - (Me.Height - Me.InsideHeight) - 6

If maxH < minH Then
    ' 逕ｻ髱｢縺悟ｰ上＆縺吶℃繧句ｴ蜷医・縲∵怙菴弱し繧､繧ｺ繧貞━蜈・
    Me.Height = minH
Else
    Me.Height = maxH
    If Me.Height < minH Then Me.Height = minH
End If



    Dim frHeader As MSForms.Frame
    Dim frViewport As MSForms.Frame
    Dim mp As MSForms.MultiPage

    '--- Header・域桃菴懊ヰ繝ｼ・壼崋螳壹・髱槭せ繧ｯ繝ｭ繝ｼ繝ｫ・・--
    On Error Resume Next
    Set frHeader = Me.controls("frHeader")
    On Error GoTo 0
    If frHeader Is Nothing Then
        Set frHeader = Me.controls.Add("Forms.Frame.1", "frHeader", True)
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

    '--- 譌｢蟄倥・繝｡繧､繝ｳ MultiPage・域爾縺吶□縺代ゆｽ懊ｉ縺ｪ縺・ｼ・--
    Set mp = FindMainMultiPage()

    If Not mp Is Nothing Then
        mp.Left = PAD_SIDE
        mp.Top = frHeader.Top + frHeader.Height + GAP_V
        mp.Width = Me.InsideWidth - PAD_SIDE * 2
       mp.Height = Application.Max(120, (maxH - (Me.Height - Me.InsideHeight)) - mp.Top - PAD_SIDE)

        ' 鬮倥＆縺ｯ蠕梧ｮｵ縺ｮViewport縺ｧ豎ｺ繧√ｋ
    End If

    '--- Viewport・郁ｩ穂ｾ｡繧ｳ繝ｳ繝・Φ繝・ｰら畑・壼ｿ・ｦ√↑繧峨せ繧ｯ繝ｭ繝ｼ繝ｫ・・--
    On Error Resume Next
    Set frViewport = Me.controls("frViewport")
    On Error GoTo 0
    If frViewport Is Nothing Then
        Set frViewport = Me.controls.Add("Forms.Frame.1", "frViewport", True)
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

10  stepN = "GetHeader": Dim f As MSForms.Frame: Set f = Me.controls("frHeader")

20  stepN = "GetButtons"
    Dim bClear As MSForms.Control, bSave As MSForms.Control, bClose As MSForms.Control
    Set bClear = Me.controls("cmdClearGlobal")
    Set bSave = Me.controls("cmdSaveGlobal")
    Set bClose = Me.controls("btnCloseCtl")

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
   Dim mp1 As Object
   Dim pg1 As Object
   Dim fr32 As Object
   Dim btnLoadPrev As Object


    Static done As Boolean
    If done Then Exit Sub
    done = True

    Dim f As MSForms.Frame
    Set f = Me.controls("frHeader")

    ' 譌｢蟄倥・繧ｿ繝ｳ・亥・逅・・譛ｬ菴難ｼ・
    Dim bClear As MSForms.Control, bSave As MSForms.Control, bClose As MSForms.Control
    Set bClear = Me.controls("cmdClearGlobal")
    Set bSave = Me.controls("cmdSaveGlobal")
    Set bClose = Me.controls("btnCloseCtl")

    ' 譌｢蟄倥・隕九∴縺ｪ縺上☆繧具ｼ井ｽ咲ｽｮ縺ｯ隗ｦ繧峨↑縺・ｼ・
    bClear.Visible = False
    bSave.Visible = False
    bClose.Visible = False

    ' 繝倥ャ繝繝ｼ逕ｨ縺ｮ譁ｰ繝懊ち繝ｳ繧剃ｽ懊ｋ・亥錐蜑榊崋螳夲ｼ・
    Dim hClear As MSForms.CommandButton
    Dim hSave  As MSForms.CommandButton
    Dim hClose As MSForms.CommandButton

     On Error Resume Next
    Set hClear = f.controls("cmdClearHeader")
    Set hSave = f.controls("cmdSaveHeader")
    Set hClose = f.controls("cmdCloseHeader")
    On Error GoTo 0

    If hClear Is Nothing Then Set hClear = f.controls.Add("Forms.CommandButton.1", "cmdClearHeader", True)
    If hSave Is Nothing Then Set hSave = f.controls.Add("Forms.CommandButton.1", "cmdSaveHeader", True)
    If hClose Is Nothing Then Set hClose = f.controls.Add("Forms.CommandButton.1", "cmdCloseHeader", True)

    ' 隕九◆逶ｮ縺ｯ譌｢蟄倥ｒ雕剰･ｲ
    hClear.caption = bClear.caption: hClear.Width = bClear.Width: hClear.Height = bClear.Height
    hSave.caption = bSave.caption:   hSave.Width = bSave.Width:   hSave.Height = bSave.Height
    hClose.caption = bClose.caption: hClose.Width = bClose.Width: hClose.Height = bClose.Height

    ' 蜿ｳ蟇・○驟咲ｽｮ
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
' LoadPrev・亥燕蝗槭・蛟､繧定ｪｭ縺ｿ霎ｼ繧・峨・繝・ム繝ｼ繝懊ち繝ｳ + Hook
'==============================
Dim hLoadPrev As MSForms.CommandButton


' 譌｢縺ｫ縺ゅｌ縺ｰ縺昴ｌ繧呈雫繧・茨ｼ昴う繝ｳ繧ｹ繧ｿ繝ｳ繧ｹ繧貞｢励ｄ縺輔↑縺・ｼ・
On Error Resume Next
Set hLoadPrev = f.controls("cmdHdrLoadPrev")
On Error GoTo 0

' 辟｡縺代ｌ縺ｰ菴懊ｋ
If hLoadPrev Is Nothing Then
    Set hLoadPrev = f.controls.Add("Forms.CommandButton.1", "cmdHdrLoadPrev", True)
End If

hLoadPrev.caption = "蜑榊屓縺ｮ蛟､繧定ｪｭ縺ｿ霎ｼ繧"
hLoadPrev.Width = 180
hLoadPrev.Height = 24
hLoadPrev.Top = hClose.Top

' 菴咲ｽｮ・嗾xtHdrKana 縺ｮ蜿ｳ・・xtHdrKana 縺檎┌縺・ｴ蜷医・蜿ｳ遶ｯ縺ｮ蟾ｦ縺ｫ鄂ｮ縺擾ｼ・
On Error Resume Next
Dim tbKana As MSForms.Control
Set tbKana = f.controls("txtHdrKana")
On Error GoTo 0

If Not tbKana Is Nothing Then
    hLoadPrev.Left = tbKana.Left + tbKana.Width + 12
Else
    hLoadPrev.Left = hClear.Left - 12 - hLoadPrev.Width
End If

' Hook・医け繝ｪ繝・け縺ｧ譌｢蟄倥・ btnLoadPrevCtl_Click 縺ｸ豬√☆・・
Set mHdrLoadPrevHook = New clsHdrBtnHook
Set mHdrLoadPrevHook.btn = hLoadPrev
mHdrLoadPrevHook.tag = "LoadPrev"
Set mHdrLoadPrevHook.owner = Me
RearrangeHeaderTopAreaLayout

' 譌ｧ繝懊ち繝ｳ縺ｯ髱櫁｡ｨ遉ｺ
On Error Resume Next
Set mp1 = Me.controls("MultiPage1")
If Not mp1 Is Nothing Then
    Set pg1 = mp1.Pages(0)
    If Not pg1 Is Nothing Then
        If Not btnLoadPrev Is Nothing Then btnLoadPrev.Visible = False

        Dim legacyLoadPrev As MSForms.Control
        Set legacyLoadPrev = EvalCtl("btnLoadPrevCtl", "Page1")
        If Not legacyLoadPrev Is Nothing Then
            legacyLoadPrev.Visible = False
            legacyLoadPrev.Enabled = False
            legacyLoadPrev.Left = -1000
            legacyLoadPrev.Top = -1000
        End If
    End If
End If
On Error GoTo 0

    
End Sub

Private Sub RearrangeHeaderTopAreaLayout()
    Dim f As Object
    Dim btnArchive As Object
    Dim btnClear As Object, btnSave As Object, btnClose As Object, btnLoadPrev As Object
    Dim lblPID As Object, lblName As Object, lblKana As Object
    Dim txtPID As Object, txtName As Object, txtKana As Object
    Dim leftX As Single, midLeft As Single, midRight As Single
    Dim btnGap As Single, rowTop1 As Single, rowTop2 As Single, rowGap As Single
    Dim idW As Single, nameW As Single

    Set f = SafeGetControl(Me, "frHeader")
    If f Is Nothing Then Exit Sub

    Set btnArchive = SafeGetControl(f, "cmdArchiveDelete")
    Set btnClear = SafeGetControl(f, "cmdClearHeader")
    Set btnSave = SafeGetControl(f, "cmdSaveHeader")
    Set btnClose = SafeGetControl(f, "cmdCloseHeader")
    Set btnLoadPrev = SafeGetControl(f, "cmdHdrLoadPrev")

    Set txtPID = SafeGetControl(f, "txtHdrPID")
    Set txtName = SafeGetControl(f, "txtHdrName")
    Set txtKana = SafeGetControl(f, "txtHdrKana")
    Set lblPID = SafeGetControl(f, "lblHdrPID")
    Set lblName = SafeGetControl(f, "lblHdrName")
    Set lblKana = SafeGetControl(f, "lblHdrKana")

    rowTop1 = 6
    rowGap = 4

    If Not btnArchive Is Nothing Then
        btnArchive.Left = 8
        btnArchive.Top = (f.Height - btnArchive.Height) / 2
        leftX = btnArchive.Left + btnArchive.Width + 14
    Else
        leftX = 8
    End If

    btnGap = 6

    If Not btnClose Is Nothing Then
        btnClose.Top = rowTop1
        btnClose.Left = f.Width - 8 - btnClose.Width
    End If

    If Not btnSave Is Nothing Then
        btnSave.Top = rowTop1
        If Not btnClose Is Nothing Then
            btnSave.Left = btnClose.Left - btnGap - btnSave.Width
        End If
    End If

    If Not btnClear Is Nothing Then
        btnClear.Top = rowTop1
        If Not btnSave Is Nothing Then
            btnClear.Left = btnSave.Left - btnGap - btnClear.Width
        End If
    End If

    If Not btnLoadPrev Is Nothing And Not btnClear Is Nothing And Not btnClose Is Nothing Then
        btnLoadPrev.Top = rowTop1 + btnClose.Height + rowGap
        btnLoadPrev.Left = btnClear.Left
        btnLoadPrev.Width = (btnClose.Left + btnClose.Width) - btnClear.Left
    End If

    midLeft = leftX
    midRight = f.Width - 8
    If Not btnClear Is Nothing Then midRight = btnClear.Left - 12
    If midRight <= midLeft Then Exit Sub

    If txtPID Is Nothing Or txtName Is Nothing Or txtKana Is Nothing Then Exit Sub

    rowTop2 = rowTop1 + txtPID.Height + rowGap

    If Not lblPID Is Nothing Then
        lblPID.AutoSize = True
        lblPID.Left = midLeft
        lblPID.Top = rowTop1 + 2
    End If

    txtPID.Left = midLeft + IIf(lblPID Is Nothing, 0, lblPID.Width + 4)
    txtPID.Top = rowTop1
    idW = 72
    txtPID.Width = idW

    If Not lblName Is Nothing Then
        lblName.AutoSize = True
        lblName.Left = txtPID.Left + txtPID.Width + 10
        lblName.Top = rowTop1 + 2
    End If

    txtName.Left = IIf(lblName Is Nothing, txtPID.Left + txtPID.Width + 10, lblName.Left + lblName.Width + 4)
    txtName.Top = rowTop1
    nameW = midRight - txtName.Left
    If nameW < 140 Then nameW = 140
    txtName.Width = nameW

    If Not lblKana Is Nothing Then
        lblKana.AutoSize = True
        lblKana.Left = IIf(lblName Is Nothing, txtName.Left - 32, lblName.Left)
        lblKana.Top = rowTop2 + 2
    End If

    txtKana.Left = txtName.Left
    txtKana.Top = rowTop2
    txtKana.Width = txtName.Width
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

    Tighten_DailyLog_Boxes_ForLayout
End Sub

Private Sub Tighten_DailyLog_Boxes_ForLayout()
    Dim f As Object
    Dim txtTraining As Object
    Dim txtReaction As Object
    Dim txtAbnormal As Object
    Dim txtPlan As Object
    Dim lst As Object
    Dim lbl As Object
    Dim fieldsBottom As Single

    Set f = GetDailyLogFrame()
    If f Is Nothing Then Exit Sub

    Set txtTraining = SafeGetControl(f, "txtDailyTraining")
    Set txtReaction = SafeGetControl(f, "txtDailyReaction")
    Set txtAbnormal = SafeGetControl(f, "txtDailyAbnormal")
    Set txtPlan = SafeGetControl(f, "txtDailyPlan")
    Set lst = SafeGetControl(f, "lstDailyLogList")
    Set lbl = SafeGetControl(f, "lblDailyHistory")

    If txtTraining Is Nothing Or txtReaction Is Nothing Or txtAbnormal Is Nothing Or txtPlan Is Nothing Then Exit Sub

    txtTraining.Height = 50
    txtReaction.Height = 50
    txtAbnormal.Height = 50
    txtPlan.Height = 50

    fieldsBottom = txtAbnormal.Top + txtAbnormal.Height
    If txtPlan.Top + txtPlan.Height > fieldsBottom Then
        fieldsBottom = txtPlan.Top + txtPlan.Height
    End If

    If Not lbl Is Nothing Then lbl.Top = fieldsBottom + 15
    If Not lst Is Nothing Then
        If Not lbl Is Nothing Then lst.Top = lbl.Top + lbl.Height + 4
        lst.Height = f.Height - lst.Top - 8
    End If
End Sub



Private Sub ApplyScroll_MP1_Page3_7_Once()
    If mScrollOnce_347 Then Exit Sub
    mScrollOnce_347 = True

   

    Dim mp As Object
    Dim f As Object
    Set mp = EvalCtl("MultiPage1")

    'Page3: Frame3・・crollHeight = 578.35 + 24 = 602.35・・
    Set f = EvalCtl("Frame3")
    If Not f Is Nothing Then
        f.ScrollBars = fmScrollBarsVertical
        f.ScrollHeight = 900
    End If

    'Page7: Frame7・亥ｿ・ｦ∵凾縺ｮ縺ｿ繝舌・陦ｨ遉ｺ・・
Set f = EvalCtl("Frame7")
If Not f Is Nothing Then
    f.ScrollHeight = 584.35
    If f.ScrollHeight > f.Height Then
        f.ScrollBars = fmScrollBarsVertical
    Else
        f.ScrollBars = fmScrollBarsNone
    End If
End If



'Page2: Frame2・亥ｧｿ蜍｢隧穂ｾ｡縺ｮ荳玖ｦ句・繧悟ｯｾ遲厄ｼ・
Set f = EvalCtl("Frame2")
If Not f Is Nothing Then
    If Not mp Is Nothing Then f.Height = mp.Height
    f.ScrollBars = fmScrollBarsVertical
    f.ScrollHeight = 488   ' 464 + 24
End If




'Page1: Frame1・亥ｰ冗判髱｢縺ｧ荳九′隕句・繧後ｋ蟇ｾ遲厄ｼ・
Set f = EvalCtl("Frame1")
If Not f Is Nothing Then
    If Not mp Is Nothing Then f.Height = mp.Height
    f.ScrollBars = fmScrollBarsVertical
    f.ScrollHeight = 420   'Sli??j
End If




End Sub

Private Sub RequestQuitExcelAskAndCloseForm()
    mQuitMode = qmAsk
    Unload Me
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

   Application.OnKey "^+D"

    

    '蜿ｳ荳翫・蟆上＆縺・暦ｼ壹ヵ繧ｩ繝ｼ繝縺縺鷹哩縺倥ｋ・・xcel縺ｯ髢峨§縺ｪ縺・ｼ・
    If CloseMode = vbFormControlMenu Then
        mQuitMode = qmNone
        Exit Sub
    End If

    '髢峨§繧九・繧ｿ繝ｳ邨檎罰縺ｮ縺ｿ・壻ｿ晏ｭ倡｢ｺ隱・竊・Excel邨ゆｺ・
    If mQuitMode = qmAsk Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("菫晏ｭ倥＠縺ｾ縺吶°・・, vbYesNoCancel + vbQuestion, "邨ゆｺ・｢ｺ隱・)

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
            '菫晏ｭ倥○縺夂ｵゆｺ・ｼ哘xcel縺ｮ菫晏ｭ倡｢ｺ隱阪ｒ蜃ｺ縺輔↑縺・
            ThisWorkbook.Saved = True
        End If

        Application.Quit
        On Error GoTo 0
    End If

End Sub




Private Sub frHeader_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
    Set f = Me.controls("txtTUG").parent
    If f Is Nothing Then Exit Sub

    Dim btn As MSForms.CommandButton
    On Error Resume Next
    Set btn = f.controls("cmdPrintTestEval")
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = f.controls.Add("Forms.CommandButton.1", "cmdPrintTestEval", True)
        With btn
       .caption = "繧ｰ繝ｩ繝募魂蛻ｷ"
        btn.Width = 120
        btn.Height = 28
        btn.Left = f.InsideWidth - btn.Width - 28.35
        btn.Top = 12

        End With
    End If
    
btn.Left = f.InsideWidth - btn.Width - 28.35

    ' 竊・笘・％縺凪・・医％縺ｮ2陦後□縺題ｿｽ蜉・・
    Set mPrintBtnHook = New clsPrintBtnHook
    Set mPrintBtnHook.btn = btn
    
    
End Sub


Public Sub BuildMonthlyDraft_FromDailyLog()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wbOpenedHere As Boolean
    Dim txtDailyDate As Object
    Dim nm As String
    Dim v As Variant
    Dim dFrom As Date, dTo As Date
    Dim lastRow As Long, r As Long
    Dim s As String
    Dim hit As Long
    Dim d As Date, staff As String, note As String

    Set txtDailyDate = DailyLogCtl("txtDailyDate")
    If txtDailyDate Is Nothing Then Exit Sub

    ' 蟇ｾ雎｡譛茨ｼ晁ｨ倬鹸譌･・・xtDailyDate・峨・譛・
    v = txtDailyDate.value
    If Not IsDate(v) Then
        MsgBox "險倬鹸譌･縺ｮ谺・↓豁｣縺励＞譌･莉倥ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, vbExclamation
        Exit Sub
    End If
    dFrom = DateSerial(Year(CDate(v)), Month(CDate(v)), 1)
    dTo = DateSerial(Year(CDate(v)), Month(CDate(v)) + 1, 0)

    ' 蟇ｾ雎｡閠・ｼ昴ヵ繧ｩ繝ｼ繝豌丞錐・・ailyLog縺ｮB蛻励→荳閾ｴ・・
    nm = Trim$(Me.controls("frHeader").controls("txtHdrName").value)
    If nm = "" Then
        MsgBox "豌丞錐繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・, vbExclamation
        Exit Sub
    End If
    
    Dim pid As String, cntSameName As Long
    

Set ws = GetDailyLogSheetByDate(dFrom, False, wb, wbOpenedHere)
If ws Is Nothing Then
    s = "縲舌Δ繝九ち繝ｪ繝ｳ繧ｰ縲・ & vbCrLf _
      & "蛻ｩ逕ｨ閠・錐・・ & nm & vbCrLf _
      & "譛滄俣・・ & Format$(dFrom, "yyyy/mm/dd") & " ・・" & Format$(dTo, "yyyy/mm/dd") & vbCrLf & vbCrLf _
      & "隧ｲ蠖捺悄髢薙・險倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・ & vbCrLf
    GoTo WriteOut
End If

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    pid = Trim$(Me.controls("frHeader").controls("txtHdrPID").value)
    cntSameName = Application.WorksheetFunction.CountIf(ws.Range("C:C"), nm)

    s = "縲舌Δ繝九ち繝ｪ繝ｳ繧ｰ縲・ & vbCrLf _
      & "蟇ｾ雎｡・・ & nm & vbCrLf _
      & "譛滄俣・・ & Format$(dFrom, "yyyy/mm/dd") & " - " & Format$(dTo, "yyyy/mm/dd") & vbCrLf & vbCrLf _
      & "笆 縺薙・譛医↓險倬鹸縺輔ｌ縺溽音險倅ｺ矩・ｼ域凾邉ｻ蛻暦ｼ・ & vbCrLf

    hit = 0
    For r = 2 To lastRow
        If Trim$(ws.Cells(r, 3).value) = nm And (cntSameName = 1 Or Trim$(ws.Cells(r, 2).value) = pid) Then
            If IsDate(ws.Cells(r, 4).value) Then
                d = CDate(ws.Cells(r, 4).value)
                If d >= dFrom And d <= dTo Then
                    note = ExtractAbnormalFindingsFromNote(CStr(ws.Cells(r, 5).value))
                    If Len(Trim$(note)) > 0 Then
                        staff = CStr(ws.Cells(r, 6).value)
                        s = s & "繝ｻ" & Format$(d, "m/d") & "・・ & staff & "・・" & note & vbCrLf
                        hit = hit + 1
                    End If
                End If
            End If
        End If
    Next r

    If hit = 0 Then
        s = s & "繝ｻ・医％縺ｮ譛医・險倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ・・ & vbCrLf
    End If

    
    ' 蜃ｺ蜉帛・・郁ｵｷ蜍墓凾縺ｫ遒ｺ菫晄ｸ医∩縺縺悟ｿｵ縺ｮ縺溘ａ・・
WriteOut:
    Call Ensure_MonthlyDraftBox_UnderFraDailyLog
    DailyLogCtl("txtMonthlyMonitoringDraft").value = s

    If wbOpenedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Function ExtractAbnormalFindingsFromNote(ByVal note As String) As String
    
    
    Dim normalized As String
    Dim startPos As Long
    Dim endPos As Long
    Dim result As String

    normalized = Replace(note, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)


    startPos = InStr(1, normalized, "逡ｰ蟶ｸ謇隕・, vbTextCompare)
    If startPos = 0 Then Exit Function

    startPos = startPos + Len("逡ｰ蟶ｸ謇隕・)
    Do While startPos <= Len(normalized)
      Select Case Mid$(normalized, startPos, 1)
          Case "・・, ":", "縲・, "]", "・・, ")", " ", "縲", vbTab, vbLf
              startPos = startPos + 1
          Case Else
              Exit Do
       End Select
    Loop

    endPos = InStr(startPos, normalized, "莉雁ｾ後・譁ｹ驥・, vbTextCompare)
    If endPos > 0 Then
        result = Mid$(normalized, startPos, endPos - startPos)
    Else
        result = Mid$(normalized, startPos)
    End If

    result = Trim$(result)
    result = Replace(result, vbLf, vbCrLf)

    ExtractAbnormalFindingsFromNote = result
End Function



Public Sub EnsureNameSuggestList()
    Dim host As Object
    Dim tb As MSForms.TextBox
    Dim lb As MSForms.ListBox

    Set host = Me
    Set tb = host.controls("txtHdrName")


    On Error Resume Next
    Set lb = Me.controls("lstNameSuggest")
    On Error GoTo 0


    If lb Is Nothing Then
        Set lb = host.controls.Add("Forms.ListBox.1", "lstNameSuggest", True)
    End If

    With lb
        .Left = Me.controls("frHeader").Left + tb.Left
        .Top = Me.controls("frHeader").Top + tb.Top + tb.Height + 4
        .Width = 200     ' 讓ｪ繧偵さ繝ｳ繝代け繝医↓
        .Height = 60     ' 3莉ｶ縺上ｉ縺・ｦ九∴繧矩ｫ倥＆

        .Visible = False
    End With
    
      Set mNameSuggestSink = New cNameSuggestSink
      mNameSuggestSink.Hook Me.controls("lstNameSuggest")
    
End Sub



Private Sub txtHdrName_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' 繝倥ャ繝蜈･蜉・竊・陬丞哨縺ｸ蜷梧悄・域里蟄倥Ο繧ｸ繝・け鬧・虚縺ｮ縺溘ａ・・
    Me.controls("txtName").text = Me.controls("frHeader").controls("txtHdrName").text

    ' 蛟呵｣廝OX遒ｺ菫・竊・譖ｴ譁ｰ
    EnsureNameSuggestList
    Me.UpdateNameSuggest
End Sub




Private Sub lstNameSuggest_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim lb As MSForms.ListBox
    Set lb = Me.controls("lstNameSuggest")

    If lb.ListIndex < 0 Then Exit Sub

    ' 繝倥ャ繝縺ｮ豌丞錐縺縺大渚譏
    Me.controls("frHeader").controls("txtHdrName").text = lb.List(lb.ListIndex, 0)

    ' 陬丞哨蜷梧悄・域里蟄倥Ο繧ｸ繝・け逕ｨ・・
    Me.controls("txtName").text = Me.controls("frHeader").controls("txtHdrName").text

    lb.Visible = False
End Sub



'====================================================
' 蜑榊屓縺ｮ蛟､繧定ｪｭ縺ｿ霎ｼ繧繝懊ち繝ｳ・壻ｽ懈・・・・鄂ｮ・・n Error縺ｪ縺暦ｼ・
'====================================================
Private Function TryGetCtl(ByVal container As Object, ByVal ctlName As String, ByRef outCtl As Object) As Boolean
    Dim c As Object
    For Each c In container.controls
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

    Set hdr = SafeGetControl(f, "frHeader")
    If hdr Is Nothing Then Set hdr = EvalCtl("frHeader")
    If hdr Is Nothing Then Exit Sub

    Set kana = SafeGetControl(hdr, "txtHdrKana")
    If kana Is Nothing Then Exit Sub

    Set btn = SafeGetControl(hdr, BTN_NAME)
    If btn Is Nothing Then
    
    
        ' 辟｡縺代ｌ縺ｰ菴懊ｋ・・rHeader驟堺ｸ具ｼ・
        Set btn = hdr.controls.Add("Forms.CommandButton.1", BTN_NAME, True)
        btn.caption = "蜑榊屓縺ｮ蛟､繧定ｪｭ縺ｿ霎ｼ繧"
        btn.Accelerator = "L"
        btn.Width = 180
        btn.Height = 24
    End If

    ' 隕九◆逶ｮ蜷医ｏ縺帷畑縺ｮ蜿ら・繝懊ち繝ｳ・医≠繧後・・・
    Set refBtn = SafeGetControl(hdr, "cmdSaveHeader")
    If refBtn Is Nothing Then Set refBtn = SafeGetControl(hdr, "cmdClearHeader")
    If refBtn Is Nothing Then Set refBtn = SafeGetControl(hdr, "cmdCloseHeader")

    If Not refBtn Is Nothing Then
        btn.Font.name = refBtn.Font.name
        btn.Font.Size = refBtn.Font.Size
       
    End If

    RearrangeHeaderTopAreaLayout
End Sub







' ====== 繧ｬ繝ｼ繝会ｼ育┌髯舌Ν繝ｼ繝鈴亟豁｢・・=====


' ====== BasicInfo・・rame32・峨さ繝ｳ繝医Ο繝ｼ繝ｫ螳滉ｽ灘叙蠕・======
Private Function BIObj(ByVal ctrlName As String) As Object

    Set BIObj = SafeGetControl(Me, ctrlName)
End Function

Private Sub EnsureBasicInfoEnterFixedRouteReady()
    Const MAX_RETRY As Long = 10
    Dim i As Long

    For i = 1 To MAX_RETRY
        If BindBasicInfoEnterFixedRouteTargets() Then Exit Sub
        DoEvents
    Next i
End Sub

Private Function BindBasicInfoEnterFixedRouteTargets() As Boolean
    Set mBIEnter_txtLiving = ResolveBasicInfoText("txtLiving")
    Set mBIEnter_txtEvaluator = ResolveBasicInfoText("txtEvaluator")
    Set mBIEnter_txtEvaluatorJob = ResolveBasicInfoText("txtEvaluatorJob")
    Set mBIEnter_txtOnset = ResolveBasicInfoText("txtOnset")
    Set mBIEnter_txtDx = ResolveBasicInfoText("txtDx")
    Set mBIEnter_txtAdmDate = ResolveBasicInfoText("txtAdmDate")
    Set mBIEnter_txtDisDate = ResolveBasicInfoText("txtDisDate")
    Set mBIEnter_txtTxCourse = ResolveBasicInfoText("txtTxCourse")

    BindBasicInfoEnterFixedRouteTargets = _
        Not (mBIEnter_txtLiving Is Nothing) And _
        Not (mBIEnter_txtEvaluator Is Nothing) And _
        Not (mBIEnter_txtEvaluatorJob Is Nothing) And _
        Not (mBIEnter_txtOnset Is Nothing) And _
        Not (mBIEnter_txtDx Is Nothing) And _
        Not (mBIEnter_txtAdmDate Is Nothing) And _
        Not (mBIEnter_txtDisDate Is Nothing) And _
        Not (mBIEnter_txtTxCourse Is Nothing)
End Function


Private Function ResolveBasicInfoText(ByVal ctrlName As String) As MSForms.TextBox
    Dim c As Object

    Set c = SafeGetControl(Me, ctrlName)
    If c Is Nothing Then Exit Function
    If TypeName(c) <> "TextBox" Then Exit Function

    Set ResolveBasicInfoText = c
End Function
Private Sub mBIEnter_txtLiving_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' txtLiving is multiline: keep Enter as newline input.
End Sub

Private Sub mBIEnter_txtEvaluator_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtEvaluatorJob
End Sub

Private Sub mBIEnter_txtEvaluatorJob_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtOnset
End Sub

Private Sub mBIEnter_txtOnset_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtDx
End Sub


Private Sub mBIEnter_txtDx_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtAdmDate
End Sub

Private Sub mBIEnter_txtAdmDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtDisDate
End Sub

Private Sub mBIEnter_txtDisDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode, mBIEnter_txtTxCourse
End Sub

Private Sub mBIEnter_txtTxCourse_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleBasicInfoEnterRoute KeyCode
End Sub

Private Sub HandleBasicInfoEnterRoute(ByRef KeyCode As MSForms.ReturnInteger, Optional ByVal nextTarget As MSForms.TextBox = Nothing)
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    If nextTarget Is Nothing Then Exit Sub
    nextTarget.SetFocus
End Sub

Private Function ReadText(ByVal o As Object) As String
    ' Value蜆ｪ蜈医√ム繝｡縺ｪ繧欝ext
    On Error Resume Next
    ReadText = CStr(CallByName(o, "Value", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        ReadText = CStr(CallByName(o, "Text", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WriteText(ByVal o As Object, ByVal s As String)
    ' Value蜆ｪ蜈医√ム繝｡縺ｪ繧欝ext
    On Error Resume Next
    CallByName o, "Value", VbLet, s
    If Err.Number <> 0 Then
        Err.Clear
        CallByName o, "Text", VbLet, s
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ====== 蟷ｴ鮨｢譖ｴ譁ｰ譛ｬ菴・======
Private Sub UpdateAgeFromBirth()
    If mAgeBusy Then Exit Sub
    mAgeBusy = True

    Dim sBirth As String, sEDate As String
    Dim dtBirth As Date, dtEval As Date
    Dim age As Long

    sBirth = Trim$(ReadText(BIObj("txtBirth")))
    If Len(sBirth) = 0 Then
        WriteText BIObj("txtAge"), ""
        mAgeBusy = False
        Exit Sub
    End If

    If Not TryParseBirthDate_ShowaOrAD(sBirth, dtBirth) Then
        ' 蜈･蜉幃比ｸｭ/荳肴ｭ｣ 竊・蟷ｴ鮨｢縺ｯ遨ｺ谺・ｼ医お繝ｩ繝ｼ縺ｯ蜃ｺ縺輔↑縺・ｼ・
        WriteText BIObj("txtAge"), ""
        mAgeBusy = False
        Exit Sub
    End If

    sEDate = Trim$(ReadText(BIObj("txtEDate")))
    If IsDate(sEDate) Then
        dtEval = CDate(sEDate)
    Else
        dtEval = Date
    End If

    age = CalcAge(dtBirth, dtEval)
    WriteText BIObj("txtAge"), CStr(age)

    mAgeBusy = False
End Sub

Private Function CalcAge(ByVal birth As Date, ByVal asOfDate As Date) As Long
    Dim y As Long
    y = Year(asOfDate) - Year(birth)
    If Format$(asOfDate, "mmdd") < Format$(birth, "mmdd") Then y = y - 1
    CalcAge = y
End Function

' 譏ｭ蜥後・隘ｿ證ｦ縺ｮ荳｡蟇ｾ蠢懶ｼ域亊蜥後□縺代〒OK縺ｪ繧峨％縺薙□縺代〒螳檎ｵ撰ｼ・
Private Function TryParseBirthDate_ShowaOrAD(ByVal raw As String, ByRef outDate As Date) As Boolean
    Dim s As String, era As String
    Dim a As Variant
    Dim y As Long, m As Long, d As Long
    Dim adY As Long

    s = Trim$(raw)
    If Len(s) = 0 Then Exit Function

    ' 縺悶▲縺上ｊ豁｣隕丞喧・域枚蟄怜喧縺代＠縺ｪ縺・ｈ縺・∫ｽｮ謠帛ｯｾ雎｡繧呈・遉ｺ・・
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    s = Replace$(s, "・・, "/")
    s = Replace$(s, "譏ｭ蜥・, "S")
    s = Replace$(s, "蟷ｴ", "/")
    s = Replace$(s, "譛・, "/")
    s = Replace$(s, "譌･", "")
    s = Replace$(s, ".", "/")
    s = Replace$(s, "-", "/")
    s = UCase$(s)

    Do While InStr(s, "//") > 0
        s = Replace$(s, "//", "/")
    Loop

    ' 蜈磯ｭ縺郡縺ｪ繧画亊蜥・
    If Left$(s, 1) = "S" Then
        era = "S"
        s = Trim$(Mid$(s, 2))
    Else
        era = ""
    End If

    ' 隘ｿ證ｦ縺ｨ縺励※隱ｭ繧√ｋ縺ｪ繧峨◎繧後〒OK・井ｺ呈鋤・・
    If era = "" Then
        If IsDate(s) Then
            outDate = CDate(s)
            TryParseBirthDate_ShowaOrAD = True
        End If
        Exit Function
    End If

    a = Split(s, "/")
    If UBound(a) < 2 Then Exit Function

    y = val(a(0)): m = val(a(1)): d = val(a(2))
    If y <= 0 Or m <= 0 Or d <= 0 Then Exit Function

    ' 譏ｭ蜥・=1926 竊・1925 + 譏ｭ蜥悟ｹｴ
    adY = 1925 + y

    On Error GoTo EH
    outDate = DateSerial(adY, m, d)
    TryParseBirthDate_ShowaOrAD = True
    Exit Function
EH:
End Function

' ====== 逋ｺ轣ｫ縺梧ｪ縺励＞迺ｰ蠅・〒繧ょｿ・★譖ｴ譁ｰ縺輔ｌ繧九ヨ繝ｪ繧ｬ ======
Private Sub MultiPage1_Change()
    ' 繧ｿ繝也ｧｻ蜍輔・繧ｿ繧､繝溘Φ繧ｰ縺ｧ蠢・★蜷梧悄
    On Error Resume Next
    UpdateAgeFromBirth
    On Error GoTo 0
    UpdateAgeFromBirth
End Sub


Public Sub Align_BIHomeEnv_Once()

    Dim fr As Object
    
    Set fr = Me.controls("mpADL").Object.Pages(0).controls("frBIHomeEnv")
    
    If Not fr Is Nothing Then
        fr.Left = fr.parent.InsideWidth - fr.Width - 12
        fr.Top = 12
        fr.ZOrder 0
    End If

End Sub
