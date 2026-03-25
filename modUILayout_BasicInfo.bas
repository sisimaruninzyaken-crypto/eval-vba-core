Attribute VB_Name = "modUILayout_BasicInfo"





Public Sub TidyBasicInfo_TwoColumns()

    Dim f32 As Object
    Dim w As Double, h As Double
    Dim xL As Double, xR As Double, wCol As Double
    Dim xLbl As Double, xCtl As Double, wLbl As Double, wCtl As Double, wCtlShort As Double
    Dim wLblR As Double, wCtlR As Double
    Dim rowH As Double, gapY As Double, multiH As Double, needsH As Double
    Dim socialH As Double
    Dim yL As Double, yR As Double
    Dim i As Long
    Dim aCapL As Variant, aCtlL As Variant
    Dim aCapR As Variant, aCtlR As Variant

    Dim c As Object
    Dim txtED As Object
    Dim t As Object
    Dim xCtlR As Double
    Dim xRightCtl As Double
    Dim wNeeds As Double
    Dim wRightMulti As Double

    Set c = frmEval.EvalCtl("txtAge", "Page1")
    If c Is Nothing Then Set c = frmEval.EvalCtl("txtEDate", "Page1")
    If c Is Nothing Then Exit Sub
    Set f32 = c.parent

    ' 縲悟､画峩轤ｹ縺ｮ縺ｿ菫晏ｭ倥阪メ繧ｧ繝・け髱櫁｡ｨ遉ｺ・域悽菴薙・繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ・・
    Set c = frmEval.EvalCtl("chkDeltaOnly", "Page1")
    If Not c Is Nothing Then
        c.Visible = False
        c.Height = 0
    End If

    ' 譌ｧ Label### 繧貞・縺ｦ髫縺呻ｼ・rame32逶ｴ荳九・縺ｿ・・
    For Each c In f32.controls
        If TypeName(c) = "Label" Then
            If Left$(c.name, 5) = "Label" Then
                c.Visible = False
            End If
        End If
    Next c

    ' 蜿ｳ繧ｫ繝ｩ繝逕ｨ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ繧堤｢ｺ菫晢ｼ育┌縺代ｌ縺ｰ霑ｽ蜉・・
    Set txtED = frmEval.EvalCtl("txtEDate", "Page1")
    
    Set t = frmEval.EvalCtl("txtEvaluatorJob", "Page1")
    If t Is Nothing Then Set t = f32.controls.Add("Forms.TextBox.1", "txtEvaluatorJob", True)
    t.tag = "BI.EvaluatorJob"
    
    Set t = frmEval.EvalCtl("txtAdmDate", "Page1")
    If t Is Nothing Then Set t = f32.controls.Add("Forms.TextBox.1", "txtAdmDate", True)

    Set t = frmEval.EvalCtl("txtDisDate", "Page1")
    If t Is Nothing Then Set t = f32.controls.Add("Forms.TextBox.1", "txtDisDate", True)

    Set t = frmEval.EvalCtl("txtTxCourse", "Page1")
    If t Is Nothing Then Set t = f32.controls.Add("Forms.TextBox.1", "txtTxCourse", True)
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True

    Set t = frmEval.EvalCtl("txtComplications", "Page1")
    If t Is Nothing Then Set t = f32.controls.Add("Forms.TextBox.1", "txtComplications", True)
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True

    w = f32.InsideWidth
    h = f32.InsideHeight

    xL = 12
    wCol = (w - 36) / 2
    xR = xL + wCol + 12

    wLbl = 140
    wCtl = wCol - wLbl - 8
    wLblR = 60
    ' 蟾ｦ繧ｫ繝ｩ繝騾壼ｸｸ蜈･蜉帙・遏ｭ繧・ｼ・eeds縺ｯ wCtl 縺ｮ縺ｾ縺ｾ・・
        wCtlShort = wCtl
        If wCtlShort > 260 Then wCtlShort = 260
    xLbl = 0
    xCtl = wLbl + 8

    rowH = 16
    gapY = 6
    multiH = 42
    socialH = 50
    needsH = 58

    ' 蜿ｳ繧ｫ繝ｩ繝縺ｮ蜈･蜉帑ｽ咲ｽｮ縺ｯ譌｢蟄・txtEDate 縺ｫ蜷医ｏ縺帙ｋ・医≠繧後・・・
     xCtlR = 60 + 8
     xRightCtl = xR + xCtlR
     wNeeds = wCol - xCtl
     wRightMulti = wCol - xCtlR

    ' 髢句ｧ倶ｽ咲ｽｮ・亥ｷｦ蜿ｳ繧ｫ繝ｩ繝繧剃ｸ閾ｴ・・
    yR = 6
    yL = yR

    ' Left: 蛟倶ｺｺ諠・ｱ・・鬆・岼・・
aCapL = Array( _
    "蟷ｴ鮨｢", _
    "逕溷ｹｴ譛域律", _
    "諤ｧ蛻･", _
    "隕∽ｻ玖ｭｷ", _
    "鬮倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", _
    "隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", _
    "遉ｾ莨壼盾蜉迥ｶ豕・ _
)
    aCtlL = Array("txtAge", "txtBirth", "cboSex", "cboCare", "cboElder", "cboDementia", "txtLiving")

    For i = 0 To UBound(aCtlL)
        Call EnsureLabel(f32, "lblBI_L_" & CStr(i + 1), CStr(aCapL(i)), xL + xLbl, yL, wLbl, rowH)
        If CStr(aCtlL(i)) = "txtLiving" Then
            Call PlaceCtl(f32, CStr(aCtlL(i)), xL + xCtl, yL - 1, wCtl, socialH)
            yL = yL + socialH + gapY
        Else
            Call PlaceCtl(f32, CStr(aCtlL(i)), xL + xCtl, yL - 1, wCtlShort, rowH + 2)
            yL = yL + rowH + gapY
        End If
    Next i

    ' Left: Needs・域悽莠ｺ/螳ｶ譌擾ｼ・
    yL = yL + 10
    Call EnsureLabel(f32, "lblBI_NeedsPt", "譛ｬ莠ｺNeeds", xL + xLbl, yL, wLbl, rowH)
    Set t = frmEval.EvalCtl("txtNeedsPt", "Page1")
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True
    Call PlaceCtl(f32, "txtNeedsPt", xL + xCtl, yL - 1, wNeeds, needsH)

    yL = yL + needsH + gapY
    Call EnsureLabel(f32, "lblBI_NeedsFam", "螳ｶ譌蒐eeds", xL + xLbl, yL, wLbl, rowH)
    Set t = frmEval.EvalCtl("txtNeedsFam", "Page1")
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True
    Call PlaceCtl(f32, "txtNeedsFam", xL + xCtl, yL - 1, wNeeds, needsH)
    yL = yL + needsH + gapY
   


   

    aCapR = Array("隧穂ｾ｡譌･", "隧穂ｾ｡閠・, "隧穂ｾ｡閠・・遞ｮ")
    aCtlR = Array("txtEDate", "txtEvaluator", "txtEvaluatorJob")
    
    For i = 0 To UBound(aCtlR)
        Call EnsureLabel(f32, "lblBI_R_E_" & CStr(i + 1), CStr(aCapR(i)), xR + xLbl, yR, wLblR, rowH)
        Call PlaceCtl(f32, CStr(aCtlR(i)), xRightCtl, yR - 1, wCtl, rowH + 2)
        yR = yR + rowH + gapY
    Next i
    
    Set c = frmEval.EvalCtl("txtEvaluatorJob", "Page1")
    If Not c Is Nothing Then c.IMEMode = fmIMEModeHiragana

    yR = yR + 4

    ' Right: 蛹ｻ逋よュ蝣ｱ
    Call EnsureLabel(f32, "lblBI_R_Header_Med", "縲仙現逋よュ蝣ｱ縲・, xR + xLbl, yR, wLblR, rowH)
    Set c = frmEval.EvalCtl("lblBI_R_Header_Med", "Page1")
    If Not c Is Nothing Then c.Visible = False
    yR = yR + gapY

    If Not ControlExists(f32, "txtAdmDate") Then f32.controls.Add "Forms.TextBox.1", "txtAdmDate"
    If Not ControlExists(f32, "txtDisDate") Then f32.controls.Add "Forms.TextBox.1", "txtDisDate"


    ' 鬆・ｺ擾ｼ夂匱逞・律竊剃ｸｻ險ｺ譁ｭ竊貞・髯｢譌･竊帝髯｢譌･
    aCapR = Array("逋ｺ逞・律", "荳ｻ險ｺ譁ｭ", "蜈･髯｢譌･", "騾髯｢譌･")
    aCtlR = Array("txtOnset", "txtDx", "txtAdmDate", "txtDisDate")

   

    For i = 0 To UBound(aCtlR)
        Call EnsureLabel(f32, "lblBI_R_M_" & CStr(i + 1), CStr(aCapR(i)), xR + xLbl, yR, wLblR, rowH)
        Call PlaceCtl(f32, CStr(aCtlR(i)), xRightCtl, yR - 1, wCtl, rowH + 2)
        yR = yR + rowH + gapY
    Next i

    ' 豐ｻ逋らｵ碁℃・郁､・焚陦鯉ｼ・
    Call EnsureLabel(f32, "lblBI_R_M_5", "豐ｻ逋らｵ碁℃", xR + xLbl, yR, wLblR, rowH)
    Call PlaceCtl(f32, "txtTxCourse", xRightCtl, yR - 1, wRightMulti, multiH)
    Set c = frmEval.EvalCtl("txtTxCourse", "Page1")
    If Not c Is Nothing Then c.IMEMode = fmIMEModeHiragana
    yR = yR + multiH + gapY

    ' 蜷井ｽｵ逞・ｼ郁､・焚陦鯉ｼ・
    Call EnsureLabel(f32, "lblBI_R_M_6", "蜷井ｽｵ逞・, xR + xLbl, yR, wLblR, rowH)
    Call PlaceCtl(f32, "txtComplications", xRightCtl, yR - 1, wRightMulti, multiH)
    Set c = frmEval.EvalCtl("txtComplications", "Page1")
    If Not c Is Nothing Then c.IMEMode = fmIMEModeHiragana
    yR = yR + multiH + 8

    ' 蜿ｳ荳具ｼ壹Μ繧ｹ繧ｯ鄒､・域怙荳区ｮｵ縺ｸ・・
    Dim riskFrame As Object
    Set riskFrame = FindBasicInfoRiskFrame(f32)
    
    Set t = frmEval.EvalCtl("txtComplications", "Page1")
    If Not t Is Nothing And Not riskFrame Is Nothing Then
        Call PlaceCtl(f32, riskFrame.name, t.Left, t.Top + 55, t.Width, h - (t.Top + 55) - 12)
    End If
    If Not riskFrame Is Nothing Then Call ArrangeRiskChecks_TwoCols(riskFrame)

End Sub

Private Function FindBasicInfoRiskFrame(ByVal parent As Object) As Object
    Dim c As Object
    Dim child As Object

    For Each c In parent.controls
        If TypeName(c) = "Frame" Then
            For Each child In c.controls
                If TypeName(child) = "CheckBox" Then
                    If CStr(child.tag) = "RiskGroup" Then
                        Set FindBasicInfoRiskFrame = c
                        Exit Function
                    End If
                End If
            Next child
        End If
    Next c
End Function


' ===== helpers =====
Private Sub PlaceCtl(ByVal parent As Object, ByVal nm As String, ByVal L As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
    Dim c As Object
    On Error Resume Next
    Set c = parent.controls(nm)
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    c.Left = L
    c.Top = t
    c.Width = w
    c.Height = h
End Sub

Private Sub ArrangeRiskChecks_TwoCols(ByVal riskFrame As Object)

    Dim names() As String
    Dim tops() As Double
    Dim lefts() As Double
    Dim cnt As Long
    Dim i As Long, j As Long

    Dim c As Object
    Dim tmpS As String
    Dim tmpD As Double

    Dim padL As Double
    Dim padT As Double
    Dim rowH As Double
    Dim gapY As Double
    Dim colGap As Double
    Dim colW As Double
    Dim x1 As Double, x2 As Double
    Dim y As Double
    Dim idx As Long
    Dim half As Long

    padL = 12
    padT = 18
    rowH = 16
    gapY = 6
    colGap = 24

    ' Collect checkboxes
    cnt = 0
    For Each c In riskFrame.controls
        If TypeName(c) = "CheckBox" Then
            cnt = cnt + 1
            ReDim Preserve names(1 To cnt)
            ReDim Preserve tops(1 To cnt)
            ReDim Preserve lefts(1 To cnt)
            names(cnt) = c.name
            tops(cnt) = c.Top
            lefts(cnt) = c.Left
        End If
    Next c
    If cnt = 0 Then Exit Sub

    ' Sort by Top, then Left (simple bubble sort)
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If (tops(j) < tops(i)) Or ((tops(j) = tops(i)) And (lefts(j) < lefts(i))) Then
                tmpS = names(i): names(i) = names(j): names(j) = tmpS
                tmpD = tops(i): tops(i) = tops(j): tops(j) = tmpD
                tmpD = lefts(i): lefts(i) = lefts(j): lefts(j) = tmpD
            End If
        Next j
    Next i

    ' Layout
    colW = (riskFrame.InsideWidth - (padL * 2) - colGap) / 2
    If colW < 60 Then colW = 60

    x1 = padL
    x2 = padL + colW + colGap

    half = (cnt + 1) \ 2

    idx = 1
    y = padT
    For i = 1 To half
        Set c = riskFrame.controls(names(idx))
        c.Left = x1
        c.Top = y
        c.Width = colW
        idx = idx + 1
        y = y + rowH + gapY
        If idx > cnt Then Exit For
    Next i

    y = padT
    For i = 1 To (cnt - half)
        Set c = riskFrame.controls(names(idx))
        c.Left = x2
        c.Top = y
        c.Width = colW
        idx = idx + 1
        y = y + rowH + gapY
        If idx > cnt Then Exit For
    Next i

End Sub

Private Sub EnsureLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal L As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
    Dim lb As Object
    On Error Resume Next
    Set lb = parent.controls(nm)
    On Error GoTo 0

    If lb Is Nothing Then
        Set lb = parent.controls.Add("Forms.Label.1", nm)
    End If

    lb.Visible = True
    lb.caption = cap
    lb.Left = L
    lb.Top = t
    lb.Width = w
    lb.Height = h
End Sub






