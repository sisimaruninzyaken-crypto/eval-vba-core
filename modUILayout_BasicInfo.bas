Attribute VB_Name = "modUILayout_BasicInfo"





Public Sub TidyBasicInfo_TwoColumns()

    Dim uf As Object, mp As Object, pg As Object, f1 As Object, f32 As Object
    Dim W As Double, h As Double
    Dim xL As Double, xR As Double, wCol As Double
    Dim xLbl As Double, xCtl As Double, wLbl As Double, wCtl As Double
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
    Dim xRightCtl As Double
    Dim riskH As Double

    Set uf = frmEval
    Set mp = uf.Controls("MultiPage1")
    Set pg = mp.Pages("Page1")
    Set f1 = pg.Controls("Frame1")
    Set f32 = f1.Controls("Frame32")

    ' 「変更点のみ保存」チェック非表示（本体はチェックボックス）
    f32.Controls("chkDeltaOnly").Visible = False
    f32.Controls("chkDeltaOnly").Height = 0

    ' 旧 Label### を全て隠す（Frame32直下のみ）
    For Each c In f32.Controls
        If TypeName(c) = "Label" Then
            If Left$(c.name, 5) = "Label" Then
                c.Visible = False
            End If
        End If
    Next c

    ' 右カラム用コントロールを確保（無ければ追加）
    On Error Resume Next
    Set txtED = f32.Controls("txtEDate")
    On Error GoTo 0
    
    Set t = Nothing
    On Error Resume Next
    Set t = f32.Controls("txtEvaluatorJob")
    On Error GoTo 0
    If t Is Nothing Then Set t = f32.Controls.Add("Forms.TextBox.1", "txtEvaluatorJob", True)
    t.tag = "BI.EvaluatorJob"
    

    On Error Resume Next
    Set t = f32.Controls("txtAdmDate")
    On Error GoTo 0
    If t Is Nothing Then Set t = f32.Controls.Add("Forms.TextBox.1", "txtAdmDate", True)

    Set t = Nothing
    On Error Resume Next
    Set t = f32.Controls("txtDisDate")
    On Error GoTo 0
    If t Is Nothing Then Set t = f32.Controls.Add("Forms.TextBox.1", "txtDisDate", True)

    Set t = Nothing
    On Error Resume Next
    Set t = f32.Controls("txtTxCourse")
    On Error GoTo 0
    If t Is Nothing Then Set t = f32.Controls.Add("Forms.TextBox.1", "txtTxCourse", True)
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True

    Set t = Nothing
    On Error Resume Next
    Set t = f32.Controls("txtComplications")
    On Error GoTo 0
    If t Is Nothing Then Set t = f32.Controls.Add("Forms.TextBox.1", "txtComplications", True)
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True

    W = f32.InsideWidth
    h = f32.InsideHeight

    xL = 12
    wCol = (W - 36) / 2
    xR = xL + wCol + 12

    wLbl = 140
    wCtl = wCol - wLbl - 8
    wLblR = 60
    ' 左カラム通常入力は短め（Needsは wCtl のまま）
        wCtlShort = wCtl
        If wCtlShort > 260 Then wCtlShort = 260
    xLbl = 0
    xCtl = wLbl + 8

    rowH = 16
    gapY = 6
    multiH = 64
    socialH = 50
    needsH = 58

    ' 右カラムの入力位置は既存 txtEDate に合わせる（あれば）
   Dim xCtlR As Double
   xCtlR = 60 + 8          '右ラベル幅60 + 余白8（ここは55?70で微調整）
   xRightCtl = xR + xCtlR

    ' 開始位置（左右カラムを一致）
    yR = 6
    yL = yR

    ' Left: 個人情報（7項目）
aCapL = Array( _
    "年齢", _
    "生年月日", _
    "性別", _
    "要介護", _
    "高齢者の日常生活自立度", _
    "認知症高齢者の日常生活自立度", _
    "社会参加状況" _
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

    ' Left: Needs（本人/家族）
    yL = yL + 10
    Call EnsureLabel(f32, "lblBI_NeedsPt", "本人Needs", xL + xLbl, yL, wLbl, rowH)
    Set t = f32.Controls("txtNeedsPt")
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True
    Call PlaceCtl(f32, "txtNeedsPt", xL + xCtl, yL - 1, wCtl, needsH)

    yL = yL + needsH + gapY
    Call EnsureLabel(f32, "lblBI_NeedsFam", "家族Needs", xL + xLbl, yL, wLbl, rowH)
    Set t = f32.Controls("txtNeedsFam")
    t.multiline = True
    t.EnterKeyBehavior = True
    t.WordWrap = True
    Call PlaceCtl(f32, "txtNeedsFam", xL + xCtl, yL - 1, wCtl, needsH)
    yL = yL + needsH + gapY
   


   

    aCapR = Array("評価日", "評価者", "評価者職種")
    aCtlR = Array("txtEDate", "txtEvaluator", "txtEvaluatorJob")
    
    For i = 0 To UBound(aCtlR)
        Call EnsureLabel(f32, "lblBI_R_E_" & CStr(i + 1), CStr(aCapR(i)), xR + xLbl, yR, wLblR, rowH)
        Call PlaceCtl(f32, CStr(aCtlR(i)), xRightCtl, yR - 1, wCtl, rowH + 2)
        yR = yR + rowH + gapY
    Next i

    yR = yR + 4

    ' Right: 医療情報
    Call EnsureLabel(f32, "lblBI_R_Header_Med", "【医療情報】", xR + xLbl, yR, wLblR, rowH)
    f32.Controls("lblBI_R_Header_Med").Visible = False
    yR = yR + rowH + gapY

    ' 順序：発症日→主診断→入院日→退院日
    aCapR = Array("発症日", "主診断", "入院日", "退院日")
    aCtlR = Array("txtOnset", "txtDx", "txtAdmDate", "txtDisDate")

    For i = 0 To UBound(aCtlR)
        Call EnsureLabel(f32, "lblBI_R_M_" & CStr(i + 1), CStr(aCapR(i)), xR + xLbl, yR, wLblR, rowH)
        Call PlaceCtl(f32, CStr(aCtlR(i)), xRightCtl, yR - 1, wCtl, rowH + 2)
        yR = yR + rowH + gapY
    Next i

    ' 治療経過（複数行）
    Call EnsureLabel(f32, "lblBI_R_M_5", "治療経過", xR + xLbl, yR, wLblR, rowH)
    Call PlaceCtl(f32, "txtTxCourse", xRightCtl, yR - 1, wCtl, multiH)
    With f32.Controls("txtTxCourse")
       .IMEMode = fmIMEModeHiragana
    End With
    yR = yR + multiH + gapY

    ' 合併症（複数行）
    Call EnsureLabel(f32, "lblBI_R_M_6", "合併症", xR + xLbl, yR, wLblR, rowH)
    Call PlaceCtl(f32, "txtComplications", xRightCtl, yR - 1, wCtl, multiH)
    With f32.Controls("txtComplications")
     .IMEMode = fmIMEModeHiragana
    End With
    yR = yR + multiH + 8

    ' 右下：リスク群（最下段へ）
    riskH = h - yR - 12
    If riskH < 24 Then riskH = 24
    Call PlaceCtl(f32, "Frame33", xR + xLbl, yR, wCol - 6, riskH)
    Call ArrangeRiskChecks_TwoCols(f32.Controls("Frame33"))

End Sub

' ===== helpers =====
Private Sub PlaceCtl(ByVal parent As Object, ByVal nm As String, ByVal L As Double, ByVal t As Double, ByVal W As Double, ByVal h As Double)
    Dim c As Object
    On Error Resume Next
    Set c = parent.Controls(nm)
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    c.Left = L
    c.Top = t
    c.Width = W
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
    For Each c In riskFrame.Controls
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
        Set c = riskFrame.Controls(names(idx))
        c.Left = x1
        c.Top = y
        c.Width = colW
        idx = idx + 1
        y = y + rowH + gapY
        If idx > cnt Then Exit For
    Next i

    y = padT
    For i = 1 To (cnt - half)
        Set c = riskFrame.Controls(names(idx))
        c.Left = x2
        c.Top = y
        c.Width = colW
        idx = idx + 1
        y = y + rowH + gapY
        If idx > cnt Then Exit For
    Next i

End Sub

Private Sub EnsureLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal L As Double, ByVal t As Double, ByVal W As Double, ByVal h As Double)
    Dim lb As Object
    On Error Resume Next
    Set lb = parent.Controls(nm)
    On Error GoTo 0

    If lb Is Nothing Then
        Set lb = parent.Controls.Add("Forms.Label.1", nm)
    End If

    lb.Visible = True
    lb.caption = cap
    lb.Left = L
    lb.Top = t
    lb.Width = W
    lb.Height = h
End Sub






