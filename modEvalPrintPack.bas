Attribute VB_Name = "modEvalPrintPack"
Option Explicit
Public Const FONT_BODY As String = "Yu Gothic UI"
Public Const FONT_SIZE_BODY As Double = 10.5
Public gForcedName As String
Public gForcedID As String

'====================================================
' A4荳｡髱｢繝代ャ繧ｯ逕滓・
' 陦ｨ・壽ｰ丞錐繝倥ャ繝 + TUG + 謠｡蜉・蜿ｳ蟾ｦ)
' 陬擾ｼ・0m豁ｩ陦・+ 5蝗樒ｫ九■荳翫′繧・+ 繧ｻ繝溘ち繝ｳ繝・Β
' 縺吶∋縺ｦ・壽怙螟ｧ8蝗・/ 蜷梧律驥崎､・・縺昴・譌･縺ｮ譛蠕後ｒ謗｡逕ｨ / 讓ｪ霆ｸ縺ｯ譌･莉倥Λ繝吶Ν縺ｮ縺ｿ
'====================================================
Public Sub Build_TestEval_PrintPack()



    Dim nm As String, idFilter As String
    Dim sh As Worksheet

If Len(gForcedName) > 0 Then
    nm = gForcedName: idFilter = gForcedID
End If




    nm = InputBox("豌丞錐・亥ｮ悟・荳閾ｴ・・)
    If Len(nm) = 0 Then Exit Sub
    idFilter = InputBox("ID縺ｧ邨槭ｋ蝣ｴ蜷医□縺大・蜉幢ｼ育ｩｺ谺・蜈ｨ莉ｶ・・)

    Set sh = ThisWorkbook.Worksheets("Viz_Print4")

    ' 繧ｷ繝ｼ繝亥・譛溷喧・域里蟄倥メ繝｣繝ｼ繝亥炎髯､・・
    ClearSheetAndCharts sh

    ' 繝壹・繧ｸ險ｭ螳夲ｼ・4邵ｦ繝ｻ2繝壹・繧ｸ・・
    SetupPrint3PagesA4 sh, nm

        

        
        
       ' =========================
' 1譫夂岼・夊ｩ穂ｾ｡邨先棡・亥・譫舌ユ繧ｭ繧ｹ繝茨ｼ・
' =========================
With sh.Range("A1")
    .value = "豌丞錐・・" & nm
    .Font.Size = 20
    .Font.Bold = True
End With



' =========================
' 謾ｹ繝壹・繧ｸ・・譫夂岼/3譫夂岼
' =========================
sh.ResetAllPageBreaks
sh.HPageBreaks.Add Before:=sh.rows(58)  ' 2譫夂岼髢句ｧ・
sh.HPageBreaks.Add Before:=sh.rows(117)  ' 3譫夂岼髢句ｧ・

' =========================
' 2譫夂岼・壹げ繝ｩ繝・縺､・・UG/謠｡蜉幢ｼ・
' =========================
AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "TUG謗ｨ遘ｻ・育ｧ抵ｼ・, "Test_TUG_sec", "遘・, _
    15, 850, 500, 220


AddGripChart_FromIO sh, nm, idFilter, _
    "謠｡蜉帶耳遘ｻ・亥承/蟾ｦ kg・・, "kg", _
    15, 1085, 500, 220

AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "10m豁ｩ陦梧耳遘ｻ・育ｧ抵ｼ・, "Test_10MWalk_sec", "遘・, _
    15, 1320, 500, 220
    
    
    ' =========================
' 3譫夂岼・壹げ繝ｩ繝・縺､・・0m/5STS/繧ｻ繝滂ｼ・
' =========================


AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "5蝗樒ｫ九■荳翫′繧頑耳遘ｻ・育ｧ抵ｼ・, "Test_5xSitStand_sec", "遘・, _
    15, 1600, 500, 220

AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "繧ｻ繝溘ち繝ｳ繝・Β謗ｨ遘ｻ・育ｧ抵ｼ・, "Test_SemiTandem_sec", "遘・, _
    15, 1835, 500, 220
 
        
        
        
        
     '蛻・梵繝・く繧ｹ繝茨ｼ医Ξ繧､繧｢繧ｦ繝育｢ｺ螳夲ｼ・
Write_AnalysisBoxes_ByRanges sh, nm, idFilter

'謾ｹ繝壹・繧ｸ縺ｯ譛蠕後↓1蝗槭□縺題ｨｭ螳・
sh.ResetAllPageBreaks
sh.HPageBreaks.Add Before:=sh.rows(58)   '2譫夂岼髢句ｧ・
sh.HPageBreaks.Add Before:=sh.rows(117)  '3譫夂岼髢句ｧ・


#If APP_DEBUG Then
    Debug.Print "HPageBreaks=" & sh.HPageBreaks.count
#End If


'繝励Ξ繝薙Η繝ｼ縺ｪ縺励〒蜊ｰ蛻ｷ
sh.PrintOut




End Sub

'====================================================
' 繝√Ε繝ｼ繝井ｽ懈・・亥腰荳邉ｻ蛻暦ｼ・
'====================================================
Private Sub AddSingleSeriesChart_FromIO( _
    ByVal sh As Worksheet, _
    ByVal nm As String, _
    ByVal idFilter As String, _
    ByVal chartTitle As String, _
    ByVal ioKey As String, _
    ByVal yUnit As String, _
    ByVal leftPt As Double, ByVal topPt As Double, ByVal widthPt As Double, ByVal heightPt As Double)

    Dim dates() As Date, vals() As Double, cnt As Long
    CollectSeries_FromIO nm, idFilter, ioKey, dates, vals, cnt

    If cnt = 0 Then
        ' 遨ｺ縺ｧ繧よ棧縺縺台ｽ懊ｉ縺壹せ繧ｭ繝・・・磯°逕ｨ荳翫ｏ縺九ｊ繧・☆縺・ｼ・
        Exit Sub
    End If

    Dim xLbl() As String, i As Long
    ReDim xLbl(1 To cnt)
    For i = 1 To cnt
        xLbl(i) = Format$(dates(i), "yyyy/mm/dd")
    Next i

    Dim co As ChartObject
    Set co = sh.ChartObjects.Add(Left:=leftPt, Top:=topPt, Width:=widthPt, Height:=heightPt)

    With co.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .chartTitle.text = chartTitle


        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = chartTitle
        .SeriesCollection(1).XValues = xLbl
        .SeriesCollection(1).values = vals

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "譌･莉・
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = yUnit
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop

        
    End With
End Sub

'====================================================
' 繝√Ε繝ｼ繝井ｽ懈・・域升蜉幢ｼ壼承蟾ｦ2邉ｻ蛻暦ｼ・
'====================================================
Private Sub AddGripChart_FromIO( _
    ByVal sh As Worksheet, _
    ByVal nm As String, _
    ByVal idFilter As String, _
    ByVal chartTitle As String, _
    ByVal yUnit As String, _
    ByVal leftPt As Double, ByVal topPt As Double, ByVal widthPt As Double, ByVal heightPt As Double)

    Dim dates() As Date, vR() As Double, vL() As Double, cnt As Long
    CollectGrip_FromIO nm, idFilter, dates, vR, vL, cnt

    If cnt = 0 Then Exit Sub

    Dim xLbl() As String, i As Long
    ReDim xLbl(1 To cnt)
    For i = 1 To cnt
        xLbl(i) = Format$(dates(i), "yyyy/mm/dd")
    Next i

    Dim co As ChartObject
    Set co = sh.ChartObjects.Add(Left:=leftPt, Top:=topPt, Width:=widthPt, Height:=heightPt)

    With co.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .chartTitle.text = chartTitle


        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "謠｡蜉・蜿ｳ(kg)"
        .SeriesCollection(1).XValues = xLbl
        .SeriesCollection(1).values = vR

        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "謠｡蜉・蟾ｦ(kg)"
        .SeriesCollection(2).XValues = xLbl
        .SeriesCollection(2).values = vL

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "譌･莉・
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = yUnit
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop

        
    End With
End Sub

'====================================================
' 繝・・繧ｿ蜿朱寔・亥腰荳邉ｻ蛻暦ｼ・
'====================================================
Private Sub CollectSeries_FromIO( _
    ByVal nm As String, _
    ByVal idFilter As String, _
    ByVal ioKey As String, _
    ByRef dates() As Date, _
    ByRef vals() As Double, _
    ByRef cnt As Long)

    Dim ws As Worksheet, lastR As Long, r As Long
    Dim idVal As String, ed As Variant, dt As Date
    Dim s As String, v As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastR = ws.Cells(ws.rows.count, 89).End(xlUp).row ' 89=豌丞錐

    cnt = 0
    For r = 2 To lastR
        If CStr(ws.Cells(r, 89).value) = nm Then
            idVal = CStr(ws.Cells(r, 97).value) ' 97=ID
            If Len(idFilter) = 0 Or idVal = idFilter Then

                ed = ws.Cells(r, 86).value ' 86=隧穂ｾ｡譌･・育｢ｺ螳夲ｼ・
                If Not IsDate(ed) Then GoTo ContinueNext
                dt = CDate(ed)

                s = CStr(ws.Cells(r, 1).Value2) ' 1=IO_TestEval
                v = GetIOVal_Pack(s, ioKey)

                If Len(v) = 0 Or v = "." Then GoTo ContinueNext
                v = Replace(v, ":", ".") ' 44:80 蟇ｾ遲・

                cnt = cnt + 1
                ReDim Preserve dates(1 To cnt)
                ReDim Preserve vals(1 To cnt)
                dates(cnt) = dt
                vals(cnt) = val(v)
            End If
        End If
ContinueNext:
    Next r

    If cnt = 0 Then Exit Sub

    SortByDate_1 dates, vals, cnt
    DedupByDate_1 dates, vals, cnt
    KeepLastN_1 dates, vals, cnt, 8
End Sub

'====================================================
' 繝・・繧ｿ蜿朱寔・域升蜉帛承蟾ｦ・・
'====================================================
Private Sub CollectGrip_FromIO( _
    ByVal nm As String, _
    ByVal idFilter As String, _
    ByRef dates() As Date, _
    ByRef vR() As Double, _
    ByRef vL() As Double, _
    ByRef cnt As Long)

    Dim ws As Worksheet, lastR As Long, r As Long
    Dim idVal As String, ed As Variant, dt As Date
    Dim s As String, sr As String, sl As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastR = ws.Cells(ws.rows.count, 89).End(xlUp).row

    cnt = 0
    For r = 2 To lastR
        If CStr(ws.Cells(r, 89).value) = nm Then
            idVal = CStr(ws.Cells(r, 97).value)
            If Len(idFilter) = 0 Or idVal = idFilter Then

                ed = ws.Cells(r, 86).value
                If Not IsDate(ed) Then GoTo ContinueNext
                dt = CDate(ed)

                s = CStr(ws.Cells(r, 1).Value2)
                sr = GetIOVal_Pack(s, "Test_Grip_R_kg")
                sl = GetIOVal_Pack(s, "Test_Grip_L_kg")

                If (Len(sr) = 0 Or sr = ".") And (Len(sl) = 0 Or sl = ".") Then GoTo ContinueNext

                sr = Replace(sr, ":", ".")
                sl = Replace(sl, ":", ".")

                cnt = cnt + 1
                ReDim Preserve dates(1 To cnt)
                ReDim Preserve vR(1 To cnt)
                ReDim Preserve vL(1 To cnt)
                dates(cnt) = dt
                vR(cnt) = IIf(Len(sr) = 0 Or sr = ".", 0, val(sr))
                vL(cnt) = IIf(Len(sl) = 0 Or sl = ".", 0, val(sl))
            End If
        End If
ContinueNext:
    Next r

    If cnt = 0 Then Exit Sub

    SortByDate_2 dates, vR, vL, cnt
    DedupByDate_2 dates, vR, vL, cnt
    KeepLastN_2 dates, vR, vL, cnt, 8
End Sub

'====================================================
' IO譁・ｭ怜・縺九ｉ key 縺ｮ蛟､繧呈栢縺擾ｼ亥玄蛻・ｊ | / 蠖｢蠑・key=value・・
'====================================================
Private Function GetIOVal_Pack(ByVal s As String, ByVal key As String) As String
    Dim parts() As String, i As Long, kv() As String
    If Len(s) = 0 Then Exit Function
    parts = Split(s, "|")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) >= 1 Then
            If Trim$(kv(0)) = key Then
                GetIOVal_Pack = Trim$(kv(1))
                Exit Function
            End If
        End If
    Next i
End Function

'====================================================
' 繧ｽ繝ｼ繝茨ｼ・酔譌･驥崎､・勁螟厄ｼ・怙螟ｧN莉ｶ
'====================================================
Private Sub SortByDate_1(ByRef d() As Date, ByRef v() As Double, ByVal cnt As Long)
    Dim i As Long, j As Long
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If d(i) > d(j) Then
                SwapDate_Pack d(i), d(j)
                SwapDbl_Pack v(i), v(j)
            End If
        Next j
    Next i
End Sub

Private Sub SortByDate_2(ByRef d() As Date, ByRef v1() As Double, ByRef v2() As Double, ByVal cnt As Long)
    Dim i As Long, j As Long
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If d(i) > d(j) Then
                SwapDate_Pack d(i), d(j)
                SwapDbl_Pack v1(i), v1(j)
                SwapDbl_Pack v2(i), v2(j)
            End If
        Next j
    Next i
End Sub

Private Sub DedupByDate_1(ByRef d() As Date, ByRef v() As Double, ByRef cnt As Long)
    Dim nd() As Date, nv() As Double, n As Long
    Dim i As Long, dayKey As Long, lastKey As Long

    n = 0
    For i = 1 To cnt
        dayKey = CLng(DateValue(d(i)))
        If n = 0 Or dayKey <> lastKey Then
            n = n + 1
            ReDim Preserve nd(1 To n)
            ReDim Preserve nv(1 To n)
            nd(n) = d(i)
            nv(n) = v(i)
            lastKey = dayKey
        Else
            nd(n) = d(i)
            nv(n) = v(i) ' 蜷梧律縺ｮ譛蠕後〒荳頑嶌縺・
        End If
    Next i

    d = nd: v = nv: cnt = n
End Sub

Private Sub DedupByDate_2(ByRef d() As Date, ByRef v1() As Double, ByRef v2() As Double, ByRef cnt As Long)
    Dim nd() As Date, n1() As Double, n2() As Double, n As Long
    Dim i As Long, dayKey As Long, lastKey As Long

    n = 0
    For i = 1 To cnt
        dayKey = CLng(DateValue(d(i)))
        If n = 0 Or dayKey <> lastKey Then
            n = n + 1
            ReDim Preserve nd(1 To n)
            ReDim Preserve n1(1 To n)
            ReDim Preserve n2(1 To n)
            nd(n) = d(i)
            n1(n) = v1(i)
            n2(n) = v2(i)
            lastKey = dayKey
        Else
            nd(n) = d(i)
            n1(n) = v1(i) ' 蜷梧律縺ｮ譛蠕後〒荳頑嶌縺・
            n2(n) = v2(i)
        End If
    Next i

    d = nd: v1 = n1: v2 = n2: cnt = n
End Sub

Private Sub KeepLastN_1(ByRef d() As Date, ByRef v() As Double, ByRef cnt As Long, ByVal n As Long)
    Dim i As Long, startIdx As Long, k As Long
    Dim nd() As Date, nv() As Double
    If cnt <= n Then Exit Sub
    startIdx = cnt - n + 1
    ReDim nd(1 To n)
    ReDim nv(1 To n)
    k = 0
    For i = startIdx To cnt
        k = k + 1
        nd(k) = d(i)
        nv(k) = v(i)
    Next i
    d = nd: v = nv: cnt = n
End Sub

Private Sub KeepLastN_2(ByRef d() As Date, ByRef v1() As Double, ByRef v2() As Double, ByRef cnt As Long, ByVal n As Long)
    Dim i As Long, startIdx As Long, k As Long
    Dim nd() As Date, n1() As Double, n2() As Double
    If cnt <= n Then Exit Sub
    startIdx = cnt - n + 1
    ReDim nd(1 To n)
    ReDim n1(1 To n)
    ReDim n2(1 To n)
    k = 0
    For i = startIdx To cnt
        k = k + 1
        nd(k) = d(i)
        n1(k) = v1(i)
        n2(k) = v2(i)
    Next i
    d = nd: v1 = n1: v2 = n2: cnt = n
End Sub

Private Sub SwapDate_Pack(ByRef a As Date, ByRef b As Date)
    Dim t As Date: t = a: a = b: b = t
End Sub

Private Sub SwapDbl_Pack(ByRef a As Double, ByRef b As Double)
    Dim t As Double: t = a: a = b: b = t
End Sub

'====================================================
' 繧ｷ繝ｼ繝亥・譛溷喧・・魂蛻ｷ險ｭ螳・
'====================================================
Private Sub ClearSheetAndCharts(ByVal sh As Worksheet)
    Dim co As ChartObject
    sh.Cells.Clear
    For Each co In sh.ChartObjects
        co.Delete
    Next co
End Sub

Private Sub SetupPrint3PagesA4(ByVal sh As Worksheet, ByVal nm As String)
    With sh.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait

        .leftMargin = Application.CentimetersToPoints(0.7)
        .RightMargin = Application.CentimetersToPoints(0.7)
        .TopMargin = Application.CentimetersToPoints(1.9)
        .BottomMargin = Application.CentimetersToPoints(1.9)

        .CenterHeader = nm

        .Zoom = 100
        .FitToPagesWide = False
        .FitToPagesTall = False
    End With
End Sub





Private Sub Write_AnalysisBoxes_ByRanges(ByVal sh As Worksheet, ByVal nm As String, ByVal idFilter As String)
    ' 谿矩ｪｸ謗・勁・医そ繝ｫ縺ｫ谿九▲縺溷・譫先枚蟄暦ｼ区立繝懊ャ繧ｯ繧ｹ・・
    On Error Resume Next
    sh.Range("B4:BC55").ClearContents
    sh.Shapes("SummaryBox").Delete
    sh.Shapes("InterpBox").Delete
    sh.Shapes("PlanBox").Delete
    On Error GoTo 0

    ' 3繝悶Ο繝・け・育ｯ・峇繧ｵ繧､繧ｺ縺ｫ霑ｽ蠕薙＠縺ｦ繝・く繧ｹ繝医・繝・け繧ｹ菴懈・・・
    PutBoxOnRange sh, "SummaryBox", sh.Range("B8:J20"), Build_Block_Summary(nm, idFilter)
    PutBoxOnRange sh, "InterpBox", sh.Range("B24:J37"), Build_Block_Interpretation(nm, idFilter)
    PutBoxOnRange sh, "PlanBox", sh.Range("B41:J54"), Build_Block_Plan(nm, idFilter)
End Sub

Private Sub PutBoxOnRange(ByVal sh As Worksheet, ByVal boxName As String, ByVal rg As Range, ByVal txt As String)
    Dim shp As Shape

    On Error Resume Next
    sh.Shapes(boxName).Delete
    On Error GoTo 0

    Set shp = sh.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=rg.Left, Top:=rg.Top, _
        Width:=rg.Width, Height:=rg.Height)

    With shp
        .name = boxName
        .Placement = xlFreeFloating

        With .TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 6
            .MarginBottom = 6
            .TextRange.text = txt
            .TextRange.Font.name = "Meiryo UI"
            .TextRange.Font.Size = 10.5
        End With
    End With
End Sub




Private Function Build_TestEval_AnalysisText_Pack(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim dr() As Date, vR() As Double, vL() As Double, cntG As Long

    Dim sTUG As String, s10m As String, s5 As String, sSemi As String, sGrip As String
    Dim body As String

    ' TUG・亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    sTUG = MetricLine_TimeSmallerBetter_Pack("TUG", "遘・, cnt, v)

    ' 10m・亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s10m = MetricLine_TimeSmallerBetter_Pack("10m豁ｩ陦・, "遘・, cnt, v)

    ' 5蝗樒ｫ九■荳翫′繧奇ｼ亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s5 = MetricLine_TimeSmallerBetter_Pack("5蝗樒ｫ九■荳翫′繧・, "遘・, cnt, v)

    ' 繧ｻ繝滂ｼ亥､ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    sSemi = MetricLine_LargerBetter_Pack("繧ｻ繝溘ち繝ｳ繝・Β", "遘・, cnt, v)

    ' 謠｡蜉幢ｼ亥､ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・∝承/蟾ｦ・・
    CollectGrip_FromIO nm, idFilter, dr, vR, vL, cntG
    sGrip = GripLine_Pack(cntG, vR, vL)

    body = ""
    body = body & "縲占ｩ穂ｾ｡邨先棡縺ｮ蛻・梵縲・ & vbCrLf
    body = body & "豌丞錐・・" & nm & vbCrLf & vbCrLf

    body = body & "笆隕∫せ・育峩霑代→蜑榊屓縺ｮ豈碑ｼ・ｼ・ & vbCrLf
    body = body & "繝ｻ" & sTUG & vbCrLf
    body = body & "繝ｻ" & s10m & vbCrLf
    body = body & "繝ｻ" & s5 & vbCrLf
    body = body & "繝ｻ" & sSemi & vbCrLf
    body = body & "繝ｻ" & sGrip & vbCrLf & vbCrLf

    body = body & "笆隗｣驥茨ｼ育ｰ｡譏難ｼ・ & vbCrLf
    body = body & "繝ｻ遘ｻ蜍戊・蜉幢ｼ啜UG/10m豁ｩ陦後′謾ｹ蝟・ｼ育洒邵ｮ・峨＠縺ｦ縺・ｌ縺ｰ縲∝虚菴憺幕蟋九・譁ｹ蜷題ｻ｢謠帙・豁ｩ陦碁溷ｺｦ縺ｮ謾ｹ蝟・′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    body = body & "繝ｻ荳玖い遲句鴨繝ｻ遶九■荳翫′繧奇ｼ・蝗樒ｫ九■荳翫′繧翫′遏ｭ邵ｮ縺吶ｌ縺ｰ縲∫ｫ九■荳翫′繧雁渚蠕ｩ縺ｮ蜉ｹ邇・髄荳翫′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    body = body & "繝ｻ繝舌Λ繝ｳ繧ｹ・壹そ繝溘ち繝ｳ繝・Β縺悟ｻｶ髟ｷ縺吶ｌ縺ｰ髱咏噪繝舌Λ繝ｳ繧ｹ縺ｮ蜷台ｸ翫′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    body = body & "繝ｻ荳願い遲句鴨・壽升蜉帙′蠅怜刈縺吶ｌ縺ｰ蜈ｨ霄ｫ遲句鴨繝ｻ豢ｻ蜍暮㍼縺ｮ謾ｹ蝟・欠讓吶・荳縺､縺ｫ縺ｪ繧翫∪縺吶・ & vbCrLf & vbCrLf

    body = body & "笆谺｡縺ｮ譁ｹ驥晢ｼ医◆縺溘″蜿ｰ・・ & vbCrLf
    body = body & "繝ｻ菴惹ｸ・蛛懈ｻ槭′縺ゅｋ鬆・岼縺ｯ縲・ｻ蠎ｦ・磯ｱ・峨・雋闕ｷ・亥屓謨ｰ/譎る俣・峨・繝輔か繝ｼ繝繧定ｦ狗峩縺励・鬆・岼縺壹▽謾ｹ蝟・ｭ悶ｒ蠖薙※縺ｾ縺吶・ & vbCrLf
    body = body & "繝ｻ霆｢蛟偵Μ繧ｹ繧ｯ縺檎桝繧上ｌ繧句ｴ蜷茨ｼ・UG謔ｪ蛹・繧ｻ繝滉ｽ惹ｸ具ｼ峨・縲∵婿蜷題ｻ｢謠帙・谿ｵ蟾ｮ繝ｻ迢ｭ謇豁ｩ陦後・邱ｴ鄙偵ｒ霑ｽ蜉縺励∪縺吶・ & vbCrLf

    Build_TestEval_AnalysisText_Pack = body
End Function

Private Function MetricLine_TimeSmallerBetter_Pack(ByVal label As String, ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then
        MetricLine_TimeSmallerBetter_Pack = label & "・壹ョ繝ｼ繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    latest = v(cnt)
    If cnt >= 2 Then
        prev = v(cnt - 1)
        diff = latest - prev
        MetricLine_TimeSmallerBetter_Pack = label & "・夂峩霑・" & latest & unit & "・亥燕蝗・" & prev & unit & "・・ & TrendWord_Time_Pack(diff)
    Else
        MetricLine_TimeSmallerBetter_Pack = label & "・夂峩霑・" & latest & unit & "・亥燕蝗槭↑縺暦ｼ・
    End If
End Function

Private Function MetricLine_LargerBetter_Pack(ByVal label As String, ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then
        MetricLine_LargerBetter_Pack = label & "・壹ョ繝ｼ繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    latest = v(cnt)
    If cnt >= 2 Then
        prev = v(cnt - 1)
        diff = latest - prev
        MetricLine_LargerBetter_Pack = label & "・夂峩霑・" & latest & unit & "・亥燕蝗・" & prev & unit & "・・ & TrendWord_Larger_Pack(diff)
    Else
        MetricLine_LargerBetter_Pack = label & "・夂峩霑・" & latest & unit & "・亥燕蝗槭↑縺暦ｼ・
    End If
End Function

Private Function GripLine_Pack(ByVal cnt As Long, ByRef vR() As Double, ByRef vL() As Double) As String
    If cnt <= 0 Then
        GripLine_Pack = "謠｡蜉幢ｼ亥承/蟾ｦ・会ｼ壹ョ繝ｼ繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim lr As Double, ll As Double, pr As Double, pl As Double
    lr = vR(cnt): ll = vL(cnt)

    If cnt >= 2 Then
        pr = vR(cnt - 1): pl = vL(cnt - 1)
        GripLine_Pack = "謠｡蜉幢ｼ亥承/蟾ｦ・会ｼ夂峩霑・" & lr & " / " & ll & "kg・亥燕蝗・" & pr & " / " & pl & "kg・・ & GripTrendWord_Pack(lr - pr, ll - pl)
    Else
        GripLine_Pack = "謠｡蜉幢ｼ亥承/蟾ｦ・会ｼ夂峩霑・" & lr & " / " & ll & "kg・亥燕蝗槭↑縺暦ｼ・
    End If
End Function

Private Function TrendWord_Time_Pack(ByVal diff As Double) As String
    If Abs(diff) < 0.000001 Then
        TrendWord_Time_Pack = "・亥､牙喧縺ｪ縺暦ｼ・
    ElseIf diff < 0 Then
        TrendWord_Time_Pack = "・域隼蝟・ｼ夂洒邵ｮ・・
    Else
        TrendWord_Time_Pack = "・井ｽ惹ｸ具ｼ壼ｻｶ髟ｷ・・
    End If
End Function

Private Function TrendWord_Larger_Pack(ByVal diff As Double) As String
    If Abs(diff) < 0.000001 Then
        TrendWord_Larger_Pack = "・亥､牙喧縺ｪ縺暦ｼ・
    ElseIf diff > 0 Then
        TrendWord_Larger_Pack = "・域隼蝟・ｼ壼髄荳奇ｼ・
    Else
        TrendWord_Larger_Pack = "・井ｽ惹ｸ具ｼ壻ｽ惹ｸ具ｼ・
    End If
End Function

Private Function GripTrendWord_Pack(ByVal diffR As Double, ByVal diffL As Double) As String
    Dim r As String, L As String
    If Abs(diffR) < 0.000001 Then
        r = "蜿ｳ=ﾂｱ0"
    ElseIf diffR > 0 Then
        r = "蜿ｳ=竊・
    Else
        r = "蜿ｳ=竊・
    End If

    If Abs(diffL) < 0.000001 Then
        L = "蟾ｦ=ﾂｱ0"
    ElseIf diffL > 0 Then
        L = "蟾ｦ=竊・
    Else
        L = "蟾ｦ=竊・
    End If

    GripTrendWord_Pack = "・・ & r & " / " & L & "・・
End Function





'====================================================
' 3繝悶Ο繝・け譖ｸ縺崎ｾｼ縺ｿ・・・曖C・・
'====================================================
Private Sub Write_TestEval_AnalysisText_3Blocks(ByVal sh As Worksheet, ByVal nm As String, ByVal idFilter As String)
    WriteBlock sh, "D6:BC10", Build_Block_Summary(nm, idFilter)
    WriteBlock sh, "D12:BC25", Build_Block_Interpretation(nm, idFilter)
    WriteBlock sh, "D27:BC40", Build_Block_Plan(nm, idFilter)
End Sub

Private Sub WriteBlock(ByVal sh As Worksheet, ByVal addr As String, ByVal txt As String)
    With sh.Range(addr)
        .ClearContents
        .Merge
        .value = txt
        .WrapText = True
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlLeft
        .Font.name = FONT_BODY
.Font.Size = FONT_SIZE_BODY

        .Font.Bold = False
        .IndentLevel = 1
    End With
End Sub

'====================================================
' 荳頑ｮｵ・夊ｦ∫ｴ・ｼ域焚蛟､・句燕蝗樊ｯ費ｼ・
'====================================================
Private Function Build_Block_Summary(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim dG() As Date, vR() As Double, vL() As Double, cntG As Long
    Dim s As String

    s = "縲占ｩ穂ｾ｡邨先棡縺ｮ隕∫ｴ・・ & vbCrLf
    s = s & "豌丞錐・・" & nm & vbCrLf
    s = s & "逶ｴ霑・蜑榊屓縺ｮ豈碑ｼ・ｼ亥酔譌･驥崎､・勁螟悶・譛螟ｧ8蝗橸ｼ・ & vbCrLf & vbCrLf

    ' TUG・亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & "繝ｻTUG・・ & SummaryLine_TimeBetterSmall(d, v, cnt, "遘・) & vbCrLf

    ' 10m・亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s = s & "繝ｻ10m豁ｩ陦鯉ｼ・ & SummaryLine_TimeBetterSmall(d, v, cnt, "遘・) & vbCrLf

    ' 5STS・亥ｰ上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & "繝ｻ5蝗樒ｫ九■荳翫′繧奇ｼ・ & SummaryLine_TimeBetterSmall(d, v, cnt, "遘・) & vbCrLf

    ' 繧ｻ繝滂ｼ亥､ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・ｼ・
    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    s = s & "繝ｻ繧ｻ繝溘ち繝ｳ繝・Β・・ & SummaryLine_BetterLarge(d, v, cnt, "遘・) & vbCrLf

    ' 謠｡蜉幢ｼ亥､ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・ｼ壼承/蟾ｦ・・
    CollectGrip_FromIO nm, idFilter, dG, vR, vL, cntG
    s = s & "繝ｻ謠｡蜉幢ｼ亥承/蟾ｦ・会ｼ・ & SummaryLine_Grip(dG, vR, vL, cntG) & vbCrLf

    Build_Block_Summary = s
End Function

Private Function SummaryLine_TimeBetterSmall(ByRef d() As Date, ByRef v() As Double, ByVal cnt As Long, ByVal unit As String) As String
    If cnt <= 0 Then
        SummaryLine_TimeBetterSmall = "繝・・繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    Dim d1 As String, d0 As String, mark As String

    latest = v(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_TimeBetterSmall = d1 & "  " & latest & unit & "・亥・蝗橸ｼ・
        Exit Function
    End If

    prev = v(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    diff = latest - prev

    If Abs(diff) < 0.000001 Then
        mark = "竊・
    ElseIf diff < 0 Then
        mark = "竊・  '遏ｭ邵ｮ=謾ｹ蝟・
    Else
        mark = "竊・  '蟒ｶ髟ｷ=謔ｪ蛹・
    End If

    SummaryLine_TimeBetterSmall = d1 & "  " & latest & unit & "  " & mark & _
                                 "・亥燕蝗・" & d0 & " " & prev & unit & " / 蟾ｮ " & Format$(diff, "+0.0;-0.0;0.0") & unit & "・・
End Function

Private Function SummaryLine_BetterLarge(ByRef d() As Date, ByRef v() As Double, ByVal cnt As Long, ByVal unit As String) As String
    If cnt <= 0 Then
        SummaryLine_BetterLarge = "繝・・繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    Dim d1 As String, d0 As String, mark As String

    latest = v(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_BetterLarge = d1 & "  " & latest & unit & "・亥・蝗橸ｼ・
        Exit Function
    End If

    prev = v(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    diff = latest - prev

    If Abs(diff) < 0.000001 Then
        mark = "竊・
    ElseIf diff > 0 Then
        mark = "竊・  '蟒ｶ髟ｷ=謾ｹ蝟・
    Else
        mark = "竊・  '遏ｭ邵ｮ=菴惹ｸ・
    End If

    SummaryLine_BetterLarge = d1 & "  " & latest & unit & "  " & mark & _
                             "・亥燕蝗・" & d0 & " " & prev & unit & " / 蟾ｮ " & Format$(diff, "+0.0;-0.0;0.0") & unit & "・・
End Function

Private Function SummaryLine_Grip(ByRef d() As Date, ByRef r() As Double, ByRef L() As Double, ByVal cnt As Long) As String
    If cnt <= 0 Then
        SummaryLine_Grip = "繝・・繧ｿ縺ｪ縺・
        Exit Function
    End If

    Dim d1 As String, d0 As String
    Dim lr As Double, ll As Double, pr As Double, pl As Double
    Dim dr As Double, dl As Double, markR As String, markL As String

    lr = r(cnt): ll = L(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_Grip = d1 & "  " & lr & " / " & ll & "kg・亥・蝗橸ｼ・
        Exit Function
    End If

    pr = r(cnt - 1): pl = L(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    dr = lr - pr: dl = ll - pl

If Abs(dr) < 0.000001 Then
    markR = "竊・
ElseIf dr > 0 Then
    markR = "竊・
Else
    markR = "竊・
End If

If Abs(dl) < 0.000001 Then
    markL = "竊・
ElseIf dl > 0 Then
    markL = "竊・
Else
    markL = "竊・
End If


    SummaryLine_Grip = d1 & "  " & lr & " / " & ll & "kg・・ & markR & "/" & markL & _
                       " 蜑榊屓 " & d0 & " " & pr & " / " & pl & "kg・・
End Function


'====================================================
' 荳ｭ谿ｵ・夊ｧ｣驥茨ｼ医げ繝ｩ繝輔・諢丞袖縺･縺托ｼ・
'====================================================
Private Function Build_Block_Interpretation(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim s As String

    s = "縲仙､牙喧縺ｮ隗｣驥医・ & vbCrLf

    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & Interpret_Time("TUG", cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s = s & Interpret_Time("10m豁ｩ陦・, cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & Interpret_Time("5蝗樒ｫ九■荳翫′繧・, cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    s = s & Interpret_Larger("繧ｻ繝溘ち繝ｳ繝・Β", cnt, v)

    Build_Block_Interpretation = s
End Function

'====================================================
' 荳区ｮｵ・壽婿驥晢ｼ域ｬ｡縺ｮ荳謇具ｼ・
'====================================================
Private Function Build_Block_Plan(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim s As String

    s = "縲蝉ｻ雁ｾ後・譁ｹ驥晢ｼ医◆縺溘″蜿ｰ・峨・ & vbCrLf

    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & Plan_Time(cnt, v, _
        "譁ｹ蜷題ｻ｢謠帙・谿ｵ蟾ｮ繝ｻ迢ｭ謇豁ｩ陦後ｒ蜷ｫ繧隱ｲ鬘後ｒ谿ｵ髫守噪縺ｫ霑ｽ蜉縺励∪縺吶・, _
        "蝓ｺ譛ｬ蜍穂ｽ懊・螳牙ｮ壼喧繧堤ｶ咏ｶ壹＠縲∬ｲ闕ｷ驥上・蠕ｮ隱ｿ謨ｴ繧定｡後＞縺ｾ縺吶・)

    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & Plan_Time(cnt, v, _
        "遶九■荳翫′繧雁渚蠕ｩ縺ｮ蝗樊焚繝ｻ繝・Φ繝昴ｒ隱ｿ謨ｴ縺励∽ｸ玖い遲句鴨蜷台ｸ翫ｒ蝗ｳ繧翫∪縺吶・, _
        "繝輔か繝ｼ繝遒ｺ隱阪ｒ蜆ｪ蜈医＠縲∝渚蠕ｩ蜉ｹ邇・・謾ｹ蝟・ｒ蝗ｳ繧翫∪縺吶・)

    Build_Block_Plan = s
End Function

'====================================================
' 遏ｭ譁・・繝ｫ繝代・
'====================================================
Private Function Line_Short_Time(ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then Line_Short_Time = "繝・・繧ｿ縺ｪ縺・: Exit Function
    If cnt >= 2 Then
        Line_Short_Time = v(cnt) & unit & "・亥燕蝗樊ｯ・" & Format$(v(cnt) - v(cnt - 1), "+0.0;-0.0;0.0") & unit & "・・
    Else
        Line_Short_Time = v(cnt) & unit & "・亥・蝗橸ｼ・
    End If
End Function

Private Function Line_Short_Larger(ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then Line_Short_Larger = "繝・・繧ｿ縺ｪ縺・: Exit Function
    If cnt >= 2 Then
        Line_Short_Larger = v(cnt) & unit & "・亥燕蝗樊ｯ・" & Format$(v(cnt) - v(cnt - 1), "+0.0;-0.0;0.0") & unit & "・・
    Else
        Line_Short_Larger = v(cnt) & unit & "・亥・蝗橸ｼ・
    End If
End Function

Private Function Line_Short_Grip(ByVal cnt As Long, ByRef r() As Double, ByRef L() As Double) As String
    If cnt <= 0 Then Line_Short_Grip = "繝・・繧ｿ縺ｪ縺・: Exit Function
    If cnt >= 2 Then
        Line_Short_Grip = r(cnt) & " / " & L(cnt) & "kg・亥燕蝗樊ｯ・" & _
            Format$(r(cnt) - r(cnt - 1), "+0.0;-0.0;0.0") & " / " & _
            Format$(L(cnt) - L(cnt - 1), "+0.0;-0.0;0.0") & "kg・・
    Else
        Line_Short_Grip = r(cnt) & " / " & L(cnt) & "kg・亥・蝗橸ｼ・
    End If
End Function

Private Function Interpret_Time(ByVal label As String, ByVal cnt As Long, ByRef v() As Double) As String
    ' 蟆上＆縺・⊇縺ｩ濶ｯ縺・ｼ・UG/10m/5STS・・
    If cnt < 2 Then
        Interpret_Time = "繝ｻ" & label & "・壼・蝗櫁ｩ穂ｾ｡縺ｮ縺溘ａ縲∽ｻ雁ｾ後・謗ｨ遘ｻ繧堤｢ｺ隱阪＠縺ｾ縺吶・ & vbCrLf
        Exit Function
    End If

    Dim d1 As Double, d3 As Double, rng As Double
    d1 = v(cnt) - v(cnt - 1) '蜑榊屓豈・
    d3 = Trend3_Pack(v, cnt, True) '逶ｴ霑・蝗槭・蛯ｾ蜷托ｼ・rue=蟆上＆縺・⊇縺ｩ濶ｯ縺・ｼ・
    rng = Range3_Pack(v, cnt)      '逶ｴ霑・蝗槭・螟牙虚蟷・

    If rng >= 5 Then
        Interpret_Time = "繝ｻ" & label & "・壽律縺ｫ繧医ｋ螟牙虚縺悟､ｧ縺阪￥縲∽ｽ楢ｪｿ繝ｻ迺ｰ蠅・・逍ｲ蜉ｴ縺ｮ蠖ｱ髻ｿ縺悟・縺ｦ縺・ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶・ & vbCrLf
    ElseIf d3 < 0 Then
        Interpret_Time = "繝ｻ" & label & "・夂峩霑・蝗槭〒謾ｹ蝟・だ蜷代′邯壹＞縺ｦ縺翫ｊ縲∝虚菴懷柑邇・・蜷台ｸ翫′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    ElseIf d3 > 0 Then
        Interpret_Time = "繝ｻ" & label & "・夂峩霑・蝗槭〒繧・ｄ菴惹ｸ句だ蜷代′縺ゅｊ縲∬ｲ闕ｷ驥上ｄ蝗槫ｾｩ迥ｶ豕√・蜀咲｢ｺ隱阪′蠢・ｦ√〒縺吶・ & vbCrLf
    ElseIf d1 < 0 Then
        Interpret_Time = "繝ｻ" & label & "・壼燕蝗槭ｈ繧顔洒邵ｮ縺励※縺翫ｊ縲∵隼蝟・′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    ElseIf d1 > 0 Then
        Interpret_Time = "繝ｻ" & label & "・壼燕蝗槭ｈ繧雁ｻｶ髟ｷ縺励※縺翫ｊ縲∫夢蜉ｴ繧・李縺ｿ遲峨・蠖ｱ髻ｿ縺瑚・∴繧峨ｌ縺ｾ縺吶・ & vbCrLf
    Else
        Interpret_Time = "繝ｻ" & label & "・壼､ｧ縺阪↑螟牙喧縺ｯ縺ｪ縺上∫ｶｭ謖∝だ蜷代〒縺吶・ & vbCrLf
    End If
End Function

Private Function Interpret_Larger(ByVal label As String, ByVal cnt As Long, ByRef v() As Double) As String
    ' 螟ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・ｼ医そ繝滂ｼ・
    If cnt < 2 Then
        Interpret_Larger = "繝ｻ" & label & "・壼・蝗櫁ｩ穂ｾ｡縺ｮ縺溘ａ縲∽ｻ雁ｾ後・謗ｨ遘ｻ繧堤｢ｺ隱阪＠縺ｾ縺吶・ & vbCrLf
        Exit Function
    End If

    Dim d1 As Double, d3 As Double, rng As Double
    d1 = v(cnt) - v(cnt - 1)
    d3 = Trend3_Pack(v, cnt, False) 'False=螟ｧ縺阪＞縺ｻ縺ｩ濶ｯ縺・
    rng = Range3_Pack(v, cnt)

    If rng >= 10 Then
        Interpret_Larger = "繝ｻ" & label & "・壽律縺ｫ繧医ｋ螟牙虚縺悟､ｧ縺阪￥縲∝ｧｿ蜍｢繝ｻ謾ｯ謖∝渕蠎暮擇繝ｻ豕ｨ諢城・蛻・・蠖ｱ髻ｿ縺悟・縺ｦ縺・ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶・ & vbCrLf
    ElseIf d3 > 0 Then
        Interpret_Larger = "繝ｻ" & label & "・夂峩霑・蝗槭〒蜷台ｸ雁だ蜷代′邯壹＞縺ｦ縺翫ｊ縲・撕逧・ヰ繝ｩ繝ｳ繧ｹ縺ｮ謾ｹ蝟・′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    ElseIf d3 < 0 Then
        Interpret_Larger = "繝ｻ" & label & "・夂峩霑・蝗槭〒菴惹ｸ句だ蜷代′縺ゅｊ縲√・繧峨▽縺崎ｦ∝屏・育夢蜉ｴ繝ｻ逍ｼ逞帙・豕ｨ諢擾ｼ峨ｒ蜀咲｢ｺ隱阪＠縺ｾ縺吶・ & vbCrLf
    ElseIf d1 > 0 Then
        Interpret_Larger = "繝ｻ" & label & "・壼燕蝗槭ｈ繧雁髄荳翫＠縺ｦ縺翫ｊ縲∵隼蝟・′遉ｺ蜚・＆繧後∪縺吶・ & vbCrLf
    ElseIf d1 < 0 Then
        Interpret_Larger = "繝ｻ" & label & "・壼燕蝗槭ｈ繧贋ｽ惹ｸ九＠縺ｦ縺翫ｊ縲∵髪謖∝渕蠎暮擇繧・ｧｿ蜍｢縺ｮ蜀咲｢ｺ隱阪′蠢・ｦ√〒縺吶・ & vbCrLf
    Else
        Interpret_Larger = "繝ｻ" & label & "・壼､ｧ縺阪↑螟牙喧縺ｯ縺ｪ縺上∫ｶｭ謖∝だ蜷代〒縺吶・ & vbCrLf
    End If
End Function

'==== 逶ｴ霑・蝗槭・縲悟だ蜷代搾ｼ壼ｹｳ蝮・→縺ｮ蟾ｮ・医＊縺｣縺上ｊ・・====
Private Function Trend3_Pack(ByRef v() As Double, ByVal cnt As Long, ByVal smallerBetter As Boolean) As Double
    Dim n As Long, a As Double, b As Double, c As Double
    If cnt >= 3 Then
        a = v(cnt - 2): b = v(cnt - 1): c = v(cnt)
        ' 蟆上＆縺・⊇縺ｩ濶ｯ縺・ｼ壻ｸ九′縺｣縺ｦ縺・ｌ縺ｰ繝槭う繝翫せ・域隼蝟・ｼ・
        If smallerBetter Then
            Trend3_Pack = (c - a)
        Else
            Trend3_Pack = (c - a)
        End If
    Else
        Trend3_Pack = 0
    End If
End Function

'==== 逶ｴ霑・蝗槭・螟牙虚蟷・ｼ・ax-min・・====
Private Function Range3_Pack(ByRef v() As Double, ByVal cnt As Long) As Double
    If cnt < 3 Then
        Range3_Pack = 0
        Exit Function
    End If
    Dim a As Double, b As Double, c As Double, mx As Double, mn As Double
    a = v(cnt - 2): b = v(cnt - 1): c = v(cnt)
    mx = a: If b > mx Then mx = b: If c > mx Then mx = c
    mn = a: If b < mn Then mn = b: If c < mn Then mn = c
    Range3_Pack = mx - mn
End Function


Private Function Plan_Time(ByVal cnt As Long, ByRef v() As Double, ByVal improveTxt As String, ByVal stableTxt As String) As String
    ' 蟆上＆縺・⊇縺ｩ濶ｯ縺・ｼ・UG/5STS縺ｪ縺ｩ・・
    If cnt <= 0 Then
        Plan_Time = "繝ｻ繝・・繧ｿ縺後↑縺・◆繧√∬ｩ穂ｾ｡繧堤ｶ咏ｶ壹＠縺ｦ蛯ｾ蜷代ｒ遒ｺ隱阪＠縺ｾ縺吶・ & vbCrLf
        Exit Function
    End If

    If cnt = 1 Then
        Plan_Time = "繝ｻ蛻晏屓隧穂ｾ｡縺ｮ縺溘ａ縲∝酔譚｡莉ｶ縺ｧ蜀崎ｩ穂ｾ｡縺怜渕貅門､繧貞崋繧√∪縺吶・ & vbCrLf
        Exit Function
    End If

    Dim diff As Double
    diff = v(cnt) - v(cnt - 1) '繝槭う繝翫せ=謾ｹ蝟・√・繝ｩ繧ｹ=謔ｪ蛹・

    ' 逶ｴ霑・蝗槭・螟牙虚蟷・′螟ｧ縺阪＞・晄律蜀・譌･髢薙・繝悶Ξ縺悟ｼｷ縺・
    Dim rng As Double
    rng = Range3_Pack(v, cnt)

    If rng >= 5 Then
        Plan_Time = "繝ｻ譌･縺ｫ繧医ｋ螟牙虚縺悟､ｧ縺阪＞縺溘ａ縲∬ｩ穂ｾ｡譚｡莉ｶ・磯擽繝ｻ陬懷勧蜈ｷ繝ｻ逍ｼ逞帙・逍ｲ蜉ｴ繝ｻ譎る俣蟶ｯ・峨ｒ謠・∴縺ｦ螳牙ｮ壼喧繧貞━蜈医＠縺ｾ縺吶・ & vbCrLf
    ElseIf diff < -0.000001 Then
        Plan_Time = "繝ｻ" & improveTxt & vbCrLf
        Plan_Time = Plan_Time & "繝ｻ謾ｹ蝟・′蜃ｺ縺ｦ縺・ｋ縺溘ａ縲∵ｮｵ髫守噪縺ｫ隱ｲ鬘碁屮蠎ｦ・域婿蜷題ｻ｢謠帙・谿ｵ蟾ｮ繝ｻ迢ｭ謇縺ｪ縺ｩ・峨ｒ荳翫￡縺ｦ豎主喧繧堤漁縺・∪縺吶・ & vbCrLf
    ElseIf diff > 0.000001 Then
        Plan_Time = "繝ｻ蜑榊屓繧医ｊ謔ｪ蛹悶＠縺ｦ縺・ｋ縺溘ａ縲∬ｲ闕ｷ驥擾ｼ亥屓謨ｰ/騾溷ｺｦ・峨→莨第・驥上ｒ隱ｿ謨ｴ縺励√ヵ繧ｩ繝ｼ繝縺ｮ蜀咲｢ｺ隱阪ｒ蜆ｪ蜈医＠縺ｾ縺吶・ & vbCrLf
        Plan_Time = Plan_Time & "繝ｻ" & stableTxt & vbCrLf
    Else
        Plan_Time = "繝ｻ邯ｭ謖∝だ蜷代・縺溘ａ縲∫樟陦後Γ繝九Η繝ｼ繧堤ｶ咏ｶ壹＠縺､縺､縲∝ｼｱ轤ｹ縺ｫ縺ｪ繧翫ｄ縺吶＞螻髱｢・磯幕蟋句虚菴懊・遶九■逶ｴ繧翫・邨ら乢逍ｲ蜉ｴ・峨ｒ驥咲せ隕ｳ蟇溘＠縺ｾ縺吶・ & vbCrLf
        Plan_Time = Plan_Time & "繝ｻ" & stableTxt & vbCrLf
    End If
End Function




