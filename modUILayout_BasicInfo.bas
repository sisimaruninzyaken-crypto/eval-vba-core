Attribute VB_Name = "modUILayout_BasicInfo"

Public Sub TidyBasicInfo_TwoColumns()

    Dim uf As Object, mp As Object, pg As Object, f1 As Object, f32 As Object
    Dim W As Double, H As Double
    Dim xL As Double, xR As Double, wCol As Double
    Dim xLbl As Double, xCtl As Double, wLbl As Double, wCtl As Double
    Dim rowH As Double, gapY As Double
    Dim yL As Double, yR As Double
    Dim i As Long
    Dim aCapL As Variant, aCtlL As Variant
    Dim aCapR As Variant, aCtlR As Variant

    Set uf = frmEval
    Set mp = uf.Controls("MultiPage1")
    Set pg = mp.Pages("Page1")
    Set f1 = pg.Controls("Frame1")
    Set f32 = f1.Controls("Frame32")

    ' 「変更点のみ保存…」を消す（本体はチェックボックス）
    f32.Controls("chkDeltaOnly").Visible = False
    f32.Controls("chkDeltaOnly").Height = 0

    ' ---- 旧: 自動採番 Label### を全て隠す（Frame32内だけ）----
    Dim c As Object
    For Each c In f32.Controls
        If TypeName(c) = "Label" Then
            If Left$(c.name, 5) = "Label" Then
                c.Visible = False
            End If
        End If
    Next c

    W = f32.InsideWidth
    H = f32.InsideHeight

    xL = 12
    wCol = (W - 36) / 2
    xR = xL + wCol + 12

    wLbl = 90
    wCtl = wCol - wLbl - 18
    xLbl = 0
    xCtl = wLbl + 8

    rowH = 16
    gapY = 6

    ' ★開始位置（左右を完全に一致させる）
    yR = 6
    yL = yR

    ' 左：個人情報（7項目）
    aCapL = Array("年齢", "生年月日", "性別", "要介護度", "生活状況", "障害高齢者の日常生活自立度", "認知症高齢者の日常生活自立度")
    aCtlL = Array("txtAge", "txtBirth", "cboSex", "cboCare", "txtLiving", "cboElder", "cboDementia")

    For i = 0 To UBound(aCtlL)
        Call EnsureLabel(f32, "lblBI_L_" & CStr(i + 1), CStr(aCapL(i)), xL + xLbl, yL, wLbl, rowH)
        Call PlaceCtl(f32, CStr(aCtlL(i)), xL + xCtl, yL - 1, wCtl, rowH + 2)
        yL = yL + rowH + gapY
    Next i

    ' 左下：Needs（本人/家族 2行）
    yL = yL + 10

    Call EnsureLabel(f32, "lblBI_NeedsPt", "本人Needs", xL + xLbl, yL, wLbl, rowH)
    Call PlaceCtl(f32, "txtNeedsPt", xL + xCtl, yL - 1, wCtl, rowH + 2)

    yL = yL + rowH + gapY

    Call EnsureLabel(f32, "lblBI_NeedsFam", "家族Needs", xL + xLbl, yL, wLbl, rowH)
    Call PlaceCtl(f32, "txtNeedsFam", xL + xCtl, yL - 1, wCtl, rowH + 2)

    ' 右：医療情報（4項目）
    aCapR = Array("評価日", "評価者", "主診断", "発症日")
    aCtlR = Array("txtEDate", "txtEvaluator", "txtDx", "txtOnset")

    For i = 0 To UBound(aCtlR)
        Call EnsureLabel(f32, "lblBI_R_" & CStr(i + 1), CStr(aCapR(i)), xR + xLbl, yR, wLbl, rowH)
        Call PlaceCtl(f32, CStr(aCtlR(i)), xR + xCtl, yR - 1, wCtl, rowH + 2)
        yR = yR + rowH + gapY
    Next i

    yR = yR + 10

    ' 右下：リスク
    Call PlaceCtl(f32, "Frame33", xR + xLbl, yR, wCol - 6, H - yR - 12)

End Sub

' ===== helpers (このモジュール内) =====
Private Sub PlaceCtl(ByVal parent As Object, ByVal nm As String, ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)
    Dim c As Object
    On Error Resume Next
    Set c = parent.Controls(nm)
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    c.Left = L
    c.Top = T
    c.Width = W
    c.Height = H
End Sub

Private Sub EnsureLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)
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
    lb.Top = T
    lb.Width = W
    lb.Height = H
End Sub

