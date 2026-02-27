Attribute VB_Name = "modEvalPrintPack"
Option Explicit
Public Const FONT_BODY As String = "Yu Gothic UI"
Public Const FONT_SIZE_BODY As Double = 10.5
Public gForcedName As String
Public gForcedID As String

'====================================================
' A4両面パック生成
' 表：氏名ヘッダ + TUG + 握力(右左)
' 裏：10m歩行 + 5回立ち上がり + セミタンデム
' すべて：最大8回 / 同日重複はその日の最後を採用 / 横軸は日付ラベルのみ
'====================================================
Public Sub Build_TestEval_PrintPack()



    Dim nm As String, idFilter As String
    Dim sh As Worksheet

If Len(gForcedName) > 0 Then
    nm = gForcedName: idFilter = gForcedID
End If




    nm = InputBox("氏名（完全一致）")
    If Len(nm) = 0 Then Exit Sub
    idFilter = InputBox("IDで絞る場合だけ入力（空欄=全件）")

    Set sh = ThisWorkbook.Worksheets("Viz_Print4")

    ' シート初期化（既存チャート削除）
    ClearSheetAndCharts sh

    ' ページ設定（A4縦・2ページ）
    SetupPrint3PagesA4 sh, nm

        

        
        
       ' =========================
' 1枚目：評価結果（分析テキスト）
' =========================
With sh.Range("A1")
    .value = "氏名： " & nm
    .Font.Size = 20
    .Font.Bold = True
End With



' =========================
' 改ページ：2枚目/3枚目
' =========================
sh.ResetAllPageBreaks
sh.HPageBreaks.Add Before:=sh.rows(58)  ' 2枚目開始
sh.HPageBreaks.Add Before:=sh.rows(117)  ' 3枚目開始

' =========================
' 2枚目：グラフ2つ（TUG/握力）
' =========================
AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "TUG推移（秒）", "Test_TUG_sec", "秒", _
    15, 850, 500, 220


AddGripChart_FromIO sh, nm, idFilter, _
    "握力推移（右/左 kg）", "kg", _
    15, 1085, 500, 220

AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "10m歩行推移（秒）", "Test_10MWalk_sec", "秒", _
    15, 1320, 500, 220
    
    
    ' =========================
' 3枚目：グラフ3つ（10m/5STS/セミ）
' =========================


AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "5回立ち上がり推移（秒）", "Test_5xSitStand_sec", "秒", _
    15, 1600, 500, 220

AddSingleSeriesChart_FromIO sh, nm, idFilter, _
    "セミタンデム推移（秒）", "Test_SemiTandem_sec", "秒", _
    15, 1835, 500, 220
 
        
        
        
        
     '分析テキスト（レイアウト確定）
Write_AnalysisBoxes_ByRanges sh, nm, idFilter

'改ページは最後に1回だけ設定
sh.ResetAllPageBreaks
sh.HPageBreaks.Add Before:=sh.rows(58)   '2枚目開始
sh.HPageBreaks.Add Before:=sh.rows(117)  '3枚目開始


#If APP_DEBUG Then
    Debug.Print "HPageBreaks=" & sh.HPageBreaks.Count
#End If


'プレビューなしで印刷
sh.PrintOut




End Sub

'====================================================
' チャート作成（単一系列）
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
        ' 空でも枠だけ作らずスキップ（運用上わかりやすい）
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
        .chartTitle.Text = chartTitle


        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = chartTitle
        .SeriesCollection(1).XValues = xLbl
        .SeriesCollection(1).values = vals

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "日付"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yUnit
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop

        
    End With
End Sub

'====================================================
' チャート作成（握力：右左2系列）
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
        .chartTitle.Text = chartTitle


        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "握力 右(kg)"
        .SeriesCollection(1).XValues = xLbl
        .SeriesCollection(1).values = vR

        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "握力 左(kg)"
        .SeriesCollection(2).XValues = xLbl
        .SeriesCollection(2).values = vL

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "日付"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yUnit
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop

        
    End With
End Sub

'====================================================
' データ収集（単一系列）
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
    lastR = ws.Cells(ws.rows.Count, 89).End(xlUp).row ' 89=氏名

    cnt = 0
    For r = 2 To lastR
        If CStr(ws.Cells(r, 89).value) = nm Then
            idVal = CStr(ws.Cells(r, 97).value) ' 97=ID
            If Len(idFilter) = 0 Or idVal = idFilter Then

                ed = ws.Cells(r, 86).value ' 86=評価日（確定）
                If Not IsDate(ed) Then GoTo ContinueNext
                dt = CDate(ed)

                s = CStr(ws.Cells(r, 1).Value2) ' 1=IO_TestEval
                v = GetIOVal_Pack(s, ioKey)

                If Len(v) = 0 Or v = "." Then GoTo ContinueNext
                v = Replace(v, ":", ".") ' 44:80 対策

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
' データ収集（握力右左）
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
    lastR = ws.Cells(ws.rows.Count, 89).End(xlUp).row

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
' IO文字列から key の値を抜く（区切り | / 形式 key=value）
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
' ソート＆同日重複除外＆最大N件
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
            nv(n) = v(i) ' 同日の最後で上書き
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
            n1(n) = v1(i) ' 同日の最後で上書き
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
' シート初期化＆印刷設定
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

        .LeftMargin = Application.CentimetersToPoints(0.7)
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
    ' 残骸掃除（セルに残った分析文字＋旧ボックス）
    On Error Resume Next
    sh.Range("B4:BC55").ClearContents
    sh.Shapes("SummaryBox").Delete
    sh.Shapes("InterpBox").Delete
    sh.Shapes("PlanBox").Delete
    On Error GoTo 0

    ' 3ブロック（範囲サイズに追従してテキストボックス作成）
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
            .TextRange.Text = txt
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

    ' TUG（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    sTUG = MetricLine_TimeSmallerBetter_Pack("TUG", "秒", cnt, v)

    ' 10m（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s10m = MetricLine_TimeSmallerBetter_Pack("10m歩行", "秒", cnt, v)

    ' 5回立ち上がり（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s5 = MetricLine_TimeSmallerBetter_Pack("5回立ち上がり", "秒", cnt, v)

    ' セミ（大きいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    sSemi = MetricLine_LargerBetter_Pack("セミタンデム", "秒", cnt, v)

    ' 握力（大きいほど良い、右/左）
    CollectGrip_FromIO nm, idFilter, dr, vR, vL, cntG
    sGrip = GripLine_Pack(cntG, vR, vL)

    body = ""
    body = body & "【評価結果の分析】" & vbCrLf
    body = body & "氏名： " & nm & vbCrLf & vbCrLf

    body = body & "■要点（直近と前回の比較）" & vbCrLf
    body = body & "・" & sTUG & vbCrLf
    body = body & "・" & s10m & vbCrLf
    body = body & "・" & s5 & vbCrLf
    body = body & "・" & sSemi & vbCrLf
    body = body & "・" & sGrip & vbCrLf & vbCrLf

    body = body & "■解釈（簡易）" & vbCrLf
    body = body & "・移動能力：TUG/10m歩行が改善（短縮）していれば、動作開始・方向転換・歩行速度の改善が示唆されます。" & vbCrLf
    body = body & "・下肢筋力・立ち上がり：5回立ち上がりが短縮すれば、立ち上がり反復の効率向上が示唆されます。" & vbCrLf
    body = body & "・バランス：セミタンデムが延長すれば静的バランスの向上が示唆されます。" & vbCrLf
    body = body & "・上肢筋力：握力が増加すれば全身筋力・活動量の改善指標の一つになります。" & vbCrLf & vbCrLf

    body = body & "■次の方針（たたき台）" & vbCrLf
    body = body & "・低下/停滞がある項目は、頻度（週）・負荷（回数/時間）・フォームを見直し、1項目ずつ改善策を当てます。" & vbCrLf
    body = body & "・転倒リスクが疑われる場合（TUG悪化/セミ低下）は、方向転換・段差・狭所歩行の練習を追加します。" & vbCrLf

    Build_TestEval_AnalysisText_Pack = body
End Function

Private Function MetricLine_TimeSmallerBetter_Pack(ByVal label As String, ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then
        MetricLine_TimeSmallerBetter_Pack = label & "：データなし"
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    latest = v(cnt)
    If cnt >= 2 Then
        prev = v(cnt - 1)
        diff = latest - prev
        MetricLine_TimeSmallerBetter_Pack = label & "：直近 " & latest & unit & "（前回 " & prev & unit & "）" & TrendWord_Time_Pack(diff)
    Else
        MetricLine_TimeSmallerBetter_Pack = label & "：直近 " & latest & unit & "（前回なし）"
    End If
End Function

Private Function MetricLine_LargerBetter_Pack(ByVal label As String, ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then
        MetricLine_LargerBetter_Pack = label & "：データなし"
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    latest = v(cnt)
    If cnt >= 2 Then
        prev = v(cnt - 1)
        diff = latest - prev
        MetricLine_LargerBetter_Pack = label & "：直近 " & latest & unit & "（前回 " & prev & unit & "）" & TrendWord_Larger_Pack(diff)
    Else
        MetricLine_LargerBetter_Pack = label & "：直近 " & latest & unit & "（前回なし）"
    End If
End Function

Private Function GripLine_Pack(ByVal cnt As Long, ByRef vR() As Double, ByRef vL() As Double) As String
    If cnt <= 0 Then
        GripLine_Pack = "握力（右/左）：データなし"
        Exit Function
    End If

    Dim lr As Double, ll As Double, pr As Double, pl As Double
    lr = vR(cnt): ll = vL(cnt)

    If cnt >= 2 Then
        pr = vR(cnt - 1): pl = vL(cnt - 1)
        GripLine_Pack = "握力（右/左）：直近 " & lr & " / " & ll & "kg（前回 " & pr & " / " & pl & "kg）" & GripTrendWord_Pack(lr - pr, ll - pl)
    Else
        GripLine_Pack = "握力（右/左）：直近 " & lr & " / " & ll & "kg（前回なし）"
    End If
End Function

Private Function TrendWord_Time_Pack(ByVal diff As Double) As String
    If Abs(diff) < 0.000001 Then
        TrendWord_Time_Pack = "（変化なし）"
    ElseIf diff < 0 Then
        TrendWord_Time_Pack = "（改善：短縮）"
    Else
        TrendWord_Time_Pack = "（低下：延長）"
    End If
End Function

Private Function TrendWord_Larger_Pack(ByVal diff As Double) As String
    If Abs(diff) < 0.000001 Then
        TrendWord_Larger_Pack = "（変化なし）"
    ElseIf diff > 0 Then
        TrendWord_Larger_Pack = "（改善：向上）"
    Else
        TrendWord_Larger_Pack = "（低下：低下）"
    End If
End Function

Private Function GripTrendWord_Pack(ByVal diffR As Double, ByVal diffL As Double) As String
    Dim r As String, l As String
    If Abs(diffR) < 0.000001 Then
        r = "右=±0"
    ElseIf diffR > 0 Then
        r = "右=↑"
    Else
        r = "右=↓"
    End If

    If Abs(diffL) < 0.000001 Then
        l = "左=±0"
    ElseIf diffL > 0 Then
        l = "左=↑"
    Else
        l = "左=↓"
    End If

    GripTrendWord_Pack = "（" & r & " / " & l & "）"
End Function





'====================================================
' 3ブロック書き込み（D～BC）
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
' 上段：要約（数値＋前回比）
'====================================================
Private Function Build_Block_Summary(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim dG() As Date, vR() As Double, vL() As Double, cntG As Long
    Dim s As String

    s = "【評価結果の要約】" & vbCrLf
    s = s & "氏名： " & nm & vbCrLf
    s = s & "直近/前回の比較（同日重複除外・最大8回）" & vbCrLf & vbCrLf

    ' TUG（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & "・TUG：" & SummaryLine_TimeBetterSmall(d, v, cnt, "秒") & vbCrLf

    ' 10m（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s = s & "・10m歩行：" & SummaryLine_TimeBetterSmall(d, v, cnt, "秒") & vbCrLf

    ' 5STS（小さいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & "・5回立ち上がり：" & SummaryLine_TimeBetterSmall(d, v, cnt, "秒") & vbCrLf

    ' セミ（大きいほど良い）
    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    s = s & "・セミタンデム：" & SummaryLine_BetterLarge(d, v, cnt, "秒") & vbCrLf

    ' 握力（大きいほど良い：右/左）
    CollectGrip_FromIO nm, idFilter, dG, vR, vL, cntG
    s = s & "・握力（右/左）：" & SummaryLine_Grip(dG, vR, vL, cntG) & vbCrLf

    Build_Block_Summary = s
End Function

Private Function SummaryLine_TimeBetterSmall(ByRef d() As Date, ByRef v() As Double, ByVal cnt As Long, ByVal unit As String) As String
    If cnt <= 0 Then
        SummaryLine_TimeBetterSmall = "データなし"
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    Dim d1 As String, d0 As String, mark As String

    latest = v(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_TimeBetterSmall = d1 & "  " & latest & unit & "（初回）"
        Exit Function
    End If

    prev = v(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    diff = latest - prev

    If Abs(diff) < 0.000001 Then
        mark = "→"
    ElseIf diff < 0 Then
        mark = "↓"  '短縮=改善
    Else
        mark = "↑"  '延長=悪化
    End If

    SummaryLine_TimeBetterSmall = d1 & "  " & latest & unit & "  " & mark & _
                                 "（前回 " & d0 & " " & prev & unit & " / 差 " & Format$(diff, "+0.0;-0.0;0.0") & unit & "）"
End Function

Private Function SummaryLine_BetterLarge(ByRef d() As Date, ByRef v() As Double, ByVal cnt As Long, ByVal unit As String) As String
    If cnt <= 0 Then
        SummaryLine_BetterLarge = "データなし"
        Exit Function
    End If

    Dim latest As Double, prev As Double, diff As Double
    Dim d1 As String, d0 As String, mark As String

    latest = v(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_BetterLarge = d1 & "  " & latest & unit & "（初回）"
        Exit Function
    End If

    prev = v(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    diff = latest - prev

    If Abs(diff) < 0.000001 Then
        mark = "→"
    ElseIf diff > 0 Then
        mark = "↑"  '延長=改善
    Else
        mark = "↓"  '短縮=低下
    End If

    SummaryLine_BetterLarge = d1 & "  " & latest & unit & "  " & mark & _
                             "（前回 " & d0 & " " & prev & unit & " / 差 " & Format$(diff, "+0.0;-0.0;0.0") & unit & "）"
End Function

Private Function SummaryLine_Grip(ByRef d() As Date, ByRef r() As Double, ByRef l() As Double, ByVal cnt As Long) As String
    If cnt <= 0 Then
        SummaryLine_Grip = "データなし"
        Exit Function
    End If

    Dim d1 As String, d0 As String
    Dim lr As Double, ll As Double, pr As Double, pl As Double
    Dim dr As Double, dl As Double, markR As String, markL As String

    lr = r(cnt): ll = l(cnt)
    d1 = Format$(d(cnt), "yyyy/mm/dd")

    If cnt = 1 Then
        SummaryLine_Grip = d1 & "  " & lr & " / " & ll & "kg（初回）"
        Exit Function
    End If

    pr = r(cnt - 1): pl = l(cnt - 1)
    d0 = Format$(d(cnt - 1), "yyyy/mm/dd")
    dr = lr - pr: dl = ll - pl

If Abs(dr) < 0.000001 Then
    markR = "→"
ElseIf dr > 0 Then
    markR = "↑"
Else
    markR = "↓"
End If

If Abs(dl) < 0.000001 Then
    markL = "→"
ElseIf dl > 0 Then
    markL = "↑"
Else
    markL = "↓"
End If


    SummaryLine_Grip = d1 & "  " & lr & " / " & ll & "kg（" & markR & "/" & markL & _
                       " 前回 " & d0 & " " & pr & " / " & pl & "kg）"
End Function


'====================================================
' 中段：解釈（グラフの意味づけ）
'====================================================
Private Function Build_Block_Interpretation(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim s As String

    s = "【変化の解釈】" & vbCrLf

    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & Interpret_Time("TUG", cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_10MWalk_sec", d, v, cnt
    s = s & Interpret_Time("10m歩行", cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & Interpret_Time("5回立ち上がり", cnt, v)

    CollectSeries_FromIO nm, idFilter, "Test_SemiTandem_sec", d, v, cnt
    s = s & Interpret_Larger("セミタンデム", cnt, v)

    Build_Block_Interpretation = s
End Function

'====================================================
' 下段：方針（次の一手）
'====================================================
Private Function Build_Block_Plan(ByVal nm As String, ByVal idFilter As String) As String
    Dim d() As Date, v() As Double, cnt As Long
    Dim s As String

    s = "【今後の方針（たたき台）】" & vbCrLf

    CollectSeries_FromIO nm, idFilter, "Test_TUG_sec", d, v, cnt
    s = s & Plan_Time(cnt, v, _
        "方向転換・段差・狭所歩行を含む課題を段階的に追加します。", _
        "基本動作の安定化を継続し、負荷量の微調整を行います。")

    CollectSeries_FromIO nm, idFilter, "Test_5xSitStand_sec", d, v, cnt
    s = s & Plan_Time(cnt, v, _
        "立ち上がり反復の回数・テンポを調整し、下肢筋力向上を図ります。", _
        "フォーム確認を優先し、反復効率の改善を図ります。")

    Build_Block_Plan = s
End Function

'====================================================
' 短文ヘルパー
'====================================================
Private Function Line_Short_Time(ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then Line_Short_Time = "データなし": Exit Function
    If cnt >= 2 Then
        Line_Short_Time = v(cnt) & unit & "（前回比 " & Format$(v(cnt) - v(cnt - 1), "+0.0;-0.0;0.0") & unit & "）"
    Else
        Line_Short_Time = v(cnt) & unit & "（初回）"
    End If
End Function

Private Function Line_Short_Larger(ByVal unit As String, ByVal cnt As Long, ByRef v() As Double) As String
    If cnt <= 0 Then Line_Short_Larger = "データなし": Exit Function
    If cnt >= 2 Then
        Line_Short_Larger = v(cnt) & unit & "（前回比 " & Format$(v(cnt) - v(cnt - 1), "+0.0;-0.0;0.0") & unit & "）"
    Else
        Line_Short_Larger = v(cnt) & unit & "（初回）"
    End If
End Function

Private Function Line_Short_Grip(ByVal cnt As Long, ByRef r() As Double, ByRef l() As Double) As String
    If cnt <= 0 Then Line_Short_Grip = "データなし": Exit Function
    If cnt >= 2 Then
        Line_Short_Grip = r(cnt) & " / " & l(cnt) & "kg（前回比 " & _
            Format$(r(cnt) - r(cnt - 1), "+0.0;-0.0;0.0") & " / " & _
            Format$(l(cnt) - l(cnt - 1), "+0.0;-0.0;0.0") & "kg）"
    Else
        Line_Short_Grip = r(cnt) & " / " & l(cnt) & "kg（初回）"
    End If
End Function

Private Function Interpret_Time(ByVal label As String, ByVal cnt As Long, ByRef v() As Double) As String
    ' 小さいほど良い（TUG/10m/5STS）
    If cnt < 2 Then
        Interpret_Time = "・" & label & "：初回評価のため、今後の推移を確認します。" & vbCrLf
        Exit Function
    End If

    Dim d1 As Double, d3 As Double, rng As Double
    d1 = v(cnt) - v(cnt - 1) '前回比
    d3 = Trend3_Pack(v, cnt, True) '直近3回の傾向（True=小さいほど良い）
    rng = Range3_Pack(v, cnt)      '直近3回の変動幅

    If rng >= 5 Then
        Interpret_Time = "・" & label & "：日による変動が大きく、体調・環境・疲労の影響が出ている可能性があります。" & vbCrLf
    ElseIf d3 < 0 Then
        Interpret_Time = "・" & label & "：直近3回で改善傾向が続いており、動作効率の向上が示唆されます。" & vbCrLf
    ElseIf d3 > 0 Then
        Interpret_Time = "・" & label & "：直近3回でやや低下傾向があり、負荷量や回復状況の再確認が必要です。" & vbCrLf
    ElseIf d1 < 0 Then
        Interpret_Time = "・" & label & "：前回より短縮しており、改善が示唆されます。" & vbCrLf
    ElseIf d1 > 0 Then
        Interpret_Time = "・" & label & "：前回より延長しており、疲労や痛み等の影響が考えられます。" & vbCrLf
    Else
        Interpret_Time = "・" & label & "：大きな変化はなく、維持傾向です。" & vbCrLf
    End If
End Function

Private Function Interpret_Larger(ByVal label As String, ByVal cnt As Long, ByRef v() As Double) As String
    ' 大きいほど良い（セミ）
    If cnt < 2 Then
        Interpret_Larger = "・" & label & "：初回評価のため、今後の推移を確認します。" & vbCrLf
        Exit Function
    End If

    Dim d1 As Double, d3 As Double, rng As Double
    d1 = v(cnt) - v(cnt - 1)
    d3 = Trend3_Pack(v, cnt, False) 'False=大きいほど良い
    rng = Range3_Pack(v, cnt)

    If rng >= 10 Then
        Interpret_Larger = "・" & label & "：日による変動が大きく、姿勢・支持基底面・注意配分の影響が出ている可能性があります。" & vbCrLf
    ElseIf d3 > 0 Then
        Interpret_Larger = "・" & label & "：直近3回で向上傾向が続いており、静的バランスの改善が示唆されます。" & vbCrLf
    ElseIf d3 < 0 Then
        Interpret_Larger = "・" & label & "：直近3回で低下傾向があり、ふらつき要因（疲労・疼痛・注意）を再確認します。" & vbCrLf
    ElseIf d1 > 0 Then
        Interpret_Larger = "・" & label & "：前回より向上しており、改善が示唆されます。" & vbCrLf
    ElseIf d1 < 0 Then
        Interpret_Larger = "・" & label & "：前回より低下しており、支持基底面や姿勢の再確認が必要です。" & vbCrLf
    Else
        Interpret_Larger = "・" & label & "：大きな変化はなく、維持傾向です。" & vbCrLf
    End If
End Function

'==== 直近3回の「傾向」：平均との差（ざっくり） ====
Private Function Trend3_Pack(ByRef v() As Double, ByVal cnt As Long, ByVal smallerBetter As Boolean) As Double
    Dim n As Long, a As Double, b As Double, c As Double
    If cnt >= 3 Then
        a = v(cnt - 2): b = v(cnt - 1): c = v(cnt)
        ' 小さいほど良い：下がっていればマイナス（改善）
        If smallerBetter Then
            Trend3_Pack = (c - a)
        Else
            Trend3_Pack = (c - a)
        End If
    Else
        Trend3_Pack = 0
    End If
End Function

'==== 直近3回の変動幅（max-min） ====
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
    ' 小さいほど良い（TUG/5STSなど）
    If cnt <= 0 Then
        Plan_Time = "・データがないため、評価を継続して傾向を確認します。" & vbCrLf
        Exit Function
    End If

    If cnt = 1 Then
        Plan_Time = "・初回評価のため、同条件で再評価し基準値を固めます。" & vbCrLf
        Exit Function
    End If

    Dim diff As Double
    diff = v(cnt) - v(cnt - 1) 'マイナス=改善、プラス=悪化

    ' 直近3回の変動幅が大きい＝日内/日間のブレが強い
    Dim rng As Double
    rng = Range3_Pack(v, cnt)

    If rng >= 5 Then
        Plan_Time = "・日による変動が大きいため、評価条件（靴・補助具・疼痛・疲労・時間帯）を揃えて安定化を優先します。" & vbCrLf
    ElseIf diff < -0.000001 Then
        Plan_Time = "・" & improveTxt & vbCrLf
        Plan_Time = Plan_Time & "・改善が出ているため、段階的に課題難度（方向転換・段差・狭所など）を上げて汎化を狙います。" & vbCrLf
    ElseIf diff > 0.000001 Then
        Plan_Time = "・前回より悪化しているため、負荷量（回数/速度）と休息量を調整し、フォームの再確認を優先します。" & vbCrLf
        Plan_Time = Plan_Time & "・" & stableTxt & vbCrLf
    Else
        Plan_Time = "・維持傾向のため、現行メニューを継続しつつ、弱点になりやすい局面（開始動作・立ち直り・終盤疲労）を重点観察します。" & vbCrLf
        Plan_Time = Plan_Time & "・" & stableTxt & vbCrLf
    End If
End Function




