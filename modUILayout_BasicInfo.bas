Attribute VB_Name = "modUILayout_BasicInfo"
Public Sub TidyBasicInfo_TwoColumns()
   
    Dim uf As Object, mp As Object, pg As Object, f1 As Object, f32 As Object
    Dim W As Double, H As Double
    Dim xL As Double, xR As Double, wCol As Double
    Dim xLbl As Double, xCtl As Double, wLbl As Double, wCtl As Double
    Dim rowH As Double, gapY As Double
    Dim yL As Double, yR As Double
    Dim i As Long
    Dim aLbl As Variant, aCtl As Variant

    Set uf = frmEval
    Set mp = uf.Controls("MultiPage1")
    Set pg = mp.Pages("Page1")
    Set f1 = pg.Controls("Frame1")
    Set f32 = f1.Controls("Frame32")

    ' 「変更点のみ保存…」を消す（本体はチェックボックス）
    f32.Controls("chkDeltaOnly").Visible = False
    f32.Controls("chkDeltaOnly").Height = 0

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
    
    ' 左：個人情報（9項目）
    aLbl = Array("Label116", "Label118", "Label117", "Label123", "Label122", "Label124", "Label125")
    aCtl = Array("txtAge", "txtBirth", "cboSex", "cboCare", "txtLiving", "cboElder", "cboDementia")
    

    For i = 0 To UBound(aCtl)
        f32.Controls(CStr(aLbl(i))).Left = xL + xLbl
        f32.Controls(CStr(aLbl(i))).Top = yL
        f32.Controls(CStr(aLbl(i))).Width = wLbl
        f32.Controls(CStr(aLbl(i))).Height = rowH

        f32.Controls(CStr(aCtl(i))).Left = xL + xCtl
        f32.Controls(CStr(aCtl(i))).Top = yL - 1
        f32.Controls(CStr(aCtl(i))).Width = wCtl
        f32.Controls(CStr(aCtl(i))).Height = rowH + 2

        yL = yL + rowH + gapY
    Next i

    ' 左下：Needs（2行）
    yL = yL + 10

    f32.Controls("Label126").Left = xL + xLbl
    f32.Controls("Label126").Top = yL
    f32.Controls("Label126").Width = wLbl
    f32.Controls("Label126").Height = rowH

    f32.Controls("txtNeedsPt").Left = xL + xCtl
    f32.Controls("txtNeedsPt").Top = yL - 1
    f32.Controls("txtNeedsPt").Width = wCtl
    f32.Controls("txtNeedsPt").Height = rowH + 2

    yL = yL + rowH + gapY

    f32.Controls("Label127").Left = xL + xLbl
    f32.Controls("Label127").Top = yL
    f32.Controls("Label127").Width = wLbl
    f32.Controls("Label127").Height = rowH

    f32.Controls("txtNeedsFam").Left = xL + xCtl
    f32.Controls("txtNeedsFam").Top = yL - 1
    f32.Controls("txtNeedsFam").Width = wCtl
    f32.Controls("txtNeedsFam").Height = rowH + 2

    ' 右：医療情報
    f32.Controls("Label113").Left = xR + xLbl
    f32.Controls("Label113").Top = yR
    f32.Controls("Label113").Width = wLbl
    f32.Controls("Label113").Height = rowH

    f32.Controls("txtEDate").Left = xR + xCtl
    f32.Controls("txtEDate").Top = yR - 1
    f32.Controls("txtEDate").Width = wCtl
    f32.Controls("txtEDate").Height = rowH + 2

    yR = yR + rowH + gapY

    f32.Controls("Label114").Left = xR + xLbl
    f32.Controls("Label114").Top = yR
    f32.Controls("Label114").Width = wLbl
    f32.Controls("Label114").Height = rowH

    f32.Controls("txtEvaluator").Left = xR + xCtl
    f32.Controls("txtEvaluator").Top = yR - 1
    f32.Controls("txtEvaluator").Width = wCtl
    f32.Controls("txtEvaluator").Height = rowH + 2

    yR = yR + rowH + gapY

    f32.Controls("Label120").Left = xR + xLbl
    f32.Controls("Label120").Top = yR
    f32.Controls("Label120").Width = wLbl
    f32.Controls("Label120").Height = rowH

    f32.Controls("txtDx").Left = xR + xCtl
    f32.Controls("txtDx").Top = yR - 1
    f32.Controls("txtDx").Width = wCtl
    f32.Controls("txtDx").Height = rowH + 2

    yR = yR + rowH + gapY

    f32.Controls("Label121").Left = xR + xLbl
    f32.Controls("Label121").Top = yR
    f32.Controls("Label121").Width = wLbl
    f32.Controls("Label121").Height = rowH

    f32.Controls("txtOnset").Left = xR + xCtl
    f32.Controls("txtOnset").Top = yR - 1
    f32.Controls("txtOnset").Width = wCtl
    f32.Controls("txtOnset").Height = rowH + 2

    yR = yR + rowH + 10

    ' 右下：リスク
    f32.Controls("Frame33").Left = xR + xLbl
    f32.Controls("Frame33").Top = yR
    f32.Controls("Frame33").Width = wCol - 6
    f32.Controls("Frame33").Height = H - yR - 12
End Sub

