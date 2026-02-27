Attribute VB_Name = "modPainIO"
'=== [TEMP] Pain IO Load Parse Helpers (一時) ===========================
Option Private Module


Option Explicit

Private Const COL_IO As Long = 156  ' HEADER_IO 列（EvalData）
Public gPainLoadEnabled As Boolean   ' 既定=False（読込禁止）

Private Function NormalizePainSite(ByVal s As String) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim a() As String, i As Long, t As String

    If Len(Trim$(s)) = 0 Then
        NormalizePainSite = ""
        Exit Function
    End If

    a = Split(s, "/")
    For i = LBound(a) To UBound(a)
        t = Trim$(a(i))
        ' 「手」と「指」はまとめて「手/指」に統一
        If t = "手" Or t = "指" Then
            d("手/指") = 1
        Else
            d(t) = 1
        End If
    Next

    NormalizePainSite = Join(d.keys, "/")
End Function



' IO文字列から key に対応する値を返す（": " 区切り、"|" レコード区切り）
Public Function IO_GetVal(ByVal ioText As String, ByVal key As String) As String
    Dim recs() As String, i As Long, t As String, p As Long, k As String, v As String
    IO_GetVal = ""
    If Len(ioText) = 0 Or Len(key) = 0 Then Exit Function

    recs = Split(CStr(ioText), "|")
    For i = LBound(recs) To UBound(recs)
        t = Trim$(recs(i))
        If Len(t) = 0 Then GoTo NextI
       p = InStr(1, t, ":")
If p = 0 Then
    p = InStr(1, t, "=")   ' ★ここを追加：=区切りにも対応
End If

If p > 0 Then
    k = Trim$(Left$(t, p - 1))
    v = Trim$(Mid$(t, p + 1))
    If StrComp(k, key, vbBinaryCompare) = 0 Then
        IO_GetVal = v
        Exit Function
    End If
End If

NextI:
    Next i
End Function


'=== [TEMP] Pain IO Load (最小：コンボ＋VAS) ============================
' 直近最終行のIOを読込、疼痛タブの主要コンボとVASへ反映
Private Sub LoadPainFromSheet_MinCombos(ByVal owner As Object)

If Not gPainLoadEnabled Then
    
    Exit Sub
End If




    Dim ws As Worksheet, lr As Long, s As String, hubRow As Long
Dim pg As Object, t As String
Set ws = ThisWorkbook.Worksheets("EvalData")

hubRow = FindLatestRowByName(ws, Trim$(owner.txtName.Text))
lr = hubRow                        '（Debug.Printで lr を使うなら合わせておく）
s = ReadStr_Compat("IO_Pain", hubRow, ws)


    
    
    
     

    Set pg = owner.Controls("mpPhys").Pages(4)
    If pg Is Nothing Then Exit Sub
' ---- Combo: 発症/期間単位/日内変動 ----
t = IO_GetVal(s, "cmbPainOnset")
If Len(t) > 0 Then pg.Controls("cmbPainOnset").value = t

' ★持続期間（数字：txtPainDuration）を読み込み
t = IO_GetVal(s, "txtPainDuration")

Debug.Print "[txtDur-Test]", "hubRow=", hubRow, " val=", t



If Len(t) > 0 Then
    On Error Resume Next
    pg.Controls("txtPainDuration").Text = t
    On Error GoTo 0
End If

Debug.Print "[txtDur-AfterSet]", "[" & pg.Controls("txtPainDuration").Text & "]"







t = IO_GetVal(s, "cmbPainDurationUnit")
If Len(t) > 0 Then pg.Controls("cmbPainDurationUnit").value = t

t = IO_GetVal(s, "cmbPainDayPeriod")
If Len(t) > 0 Then pg.Controls("cmbPainDayPeriod").value = t


    ' 最新行（IO/NOTE基準）
lr = WorksheetFunction.Max(ws.Cells(ws.rows.Count, 156).End(xlUp).row, ws.Cells(ws.rows.Count, 157).End(xlUp).row)



' ---- VAS（IO優先／空ならNOTE数値を採用）----
Dim altVAS As String
s = ReadStr_Compat("IO_Pain", hubRow, ws)
t = IO_GetVal(s, "VAS")
altVAS = ""


' いったんクリア
On Error Resume Next
pg.Controls("fraVAS").Controls("txtVAS").Text = ""
pg.Controls("fraVAS").Controls("sldVAS").value = 0
On Error GoTo 0

If Len(t) = 0 And IsNumeric(altVAS) Then t = Trim$(altVAS)
If Len(t) > 0 Then
    On Error Resume Next
    pg.Controls("fraVAS").Controls("txtVAS").Text = t
    pg.Controls("fraVAS").Controls("sldVAS").value = CLng(t)
    On Error GoTo 0
End If


   

    
End Sub

' 文字列 "A/B/C" → Dictionary(Set) 化
Private Function MakeSetFromSlash(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim a() As String, i As Long, t As String
    If Len(Trim$(s)) > 0 Then
        a = Split(s, "/")
        For i = LBound(a) To UBound(a)
            t = Trim$(a(i))
            If Len(t) > 0 Then d(t) = 1
        Next
    End If
    Set MakeSetFromSlash = d
End Function

' ListBox の選択復元（項目文字列一致）
Private Sub RestoreListBoxSelections(lb As MSForms.ListBox, ByVal slash As String)
    Dim want As Object: Set want = MakeSetFromSlash(slash)
    Dim j As Long, txt As String
    If lb Is Nothing Then Exit Sub
    For j = 0 To lb.ListCount - 1
        txt = CStr(lb.List(j))
        lb.Selected(j) = (want.exists(txt))
    Next
End Sub

Private Sub RestorePainFactors(ByVal container As Object, ByVal slash As String)
    Dim want As Object: Set want = MakeSetFromSlash(slash)
    Dim c As Object, base As String
    ' いったん全解除
    For Each c In container.Controls
        If TypeName(c) = "CheckBox" Then c.value = False
    Next
    ' 該当のみ True
    For Each c In container.Controls
        If TypeName(c) = "CheckBox" Then
            base = c.name
            If LCase$(Left$(base, 3)) = "chk" Then base = Mid$(base, 4)
            If want.exists(base) Then c.value = True
        End If
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            RestorePainFactors c, slash
        End If
    Next
End Sub


' 直近最終行のIOを読込、ListBox と Factors を復元
Private Sub LoadPainFromSheet_MinLists(ByVal owner As Object)
    Dim ws As Worksheet, lr As Long, s As String
    Dim pg As Object, t As String
 Set ws = ThisWorkbook.Worksheets("EvalData")

Dim hubRow As Long
hubRow = FindLatestRowByName(ws, Trim$(owner.txtName.Text))

Debug.Print "[PAIN][MinLists] hubRow=", hubRow   ' ←ログは残してOK

s = ReadStr_Compat("IO_Pain", hubRow, ws)

On Error Resume Next
Set pg = owner.Controls("mpPhys").Pages(4)
On Error GoTo 0
If pg Is Nothing Then Exit Sub



    
    
    

    
    Debug.Print "[DBG-PainQual-RAW]  "; IO_GetVal(s, "PainQual")
Debug.Print "[DBG-PainSite-RAW]  "; IO_GetVal(s, "PainSite")
Debug.Print "[DBG-PainFactors-RAW]"; IO_GetVal(s, "PainFactors")

    
    


' ---- ListBox : PainQual / PainSite ----
t = IO_GetVal(s, "PainQual")
On Error Resume Next
RestoreListBoxSelections pg.Controls("lstPainQual"), t

    t = IO_GetVal(s, "PainSite")
    RestoreListBoxSelections pg.Controls("lstPainSite"), t
    On Error GoTo 0

    ' ---- PainFactors : fraPainFactors 配下の CheckBox (Name一致) ----
    t = IO_GetVal(s, "PainFactors")
    On Error Resume Next
    RestorePainFactors pg.Controls("fraPainFactors"), t
    On Error GoTo 0

    
End Sub
'=== [/TEMP] ============================================================



'=== [TEMP] Pain UI Selection Probe ====================================
Private Function FindByNameRecursive(container As Object, ByVal target As String) As Object
    Dim c As Object, r As Object
    For Each c In container.Controls
        If StrComp(CStr(c.name), target, vbBinaryCompare) = 0 Then Set FindByNameRecursive = c: Exit Function
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            Set r = FindByNameRecursive(c, target)
            If Not r Is Nothing Then Set FindByNameRecursive = r: Exit Function
        End If
    Next
End Function



'=== [TEMP] Pain IO Load (NOTE) ========================================
Private Function FindLargestTextBoxOnPage(pg As Object) As MSForms.TextBox
    Dim c As Object, area As Double, bestArea As Double
    For Each c In pg.Controls
        If TypeName(c) = "TextBox" Then
            area = c.Width * c.Height
            If area > bestArea Then
                bestArea = area
                Set FindLargestTextBoxOnPage = c
            End If
        End If
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            Dim r As MSForms.TextBox
            Set r = FindLargestTextBoxOnPage(c)
            If Not r Is Nothing Then
                area = r.Width * r.Height
                If area > bestArea Then
                    bestArea = area
                    Set FindLargestTextBoxOnPage = r
                End If
            End If
        End If
    Next
End Function



'=== [TEMP] NOTE TextBox Finder =======================================
' 優先1: 名称に "Memo" を含む TextBox
' 優先2: MultiLine=True の TextBox（VAS系を除外）
' 優先3: 上記が無ければ最大面積。ただし VAS配下は除外
Private Function FindNoteTextBox(pg As Object) As MSForms.TextBox
    Dim best As MSForms.TextBox
    Dim bestArea As Double
    Dim c As Object

    ' 再帰探索
    For Each c In pg.Controls
        If TypeName(c) = "TextBox" Then
            If InStr(1, c.name, "Memo", vbTextCompare) > 0 Then
                Set FindNoteTextBox = c
                Exit Function
            End If
        End If

        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            Dim r As MSForms.TextBox
            Set r = FindNoteTextBox(c)
            If Not r Is Nothing Then Set FindNoteTextBox = r: Exit Function
        End If
    Next

    ' MultiLine 優先（VAS配下は除外）
    For Each c In pg.Controls
        If TypeName(c) = "TextBox" Then
            If SafeIsMultiLine(c) And Not IsUnderVAS(c) Then
                Set FindNoteTextBox = c
                Exit Function
            End If
        End If
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            Dim r2 As MSForms.TextBox
            Set r2 = FindNoteTextBox(c)
            If Not r2 Is Nothing Then Set FindNoteTextBox = r2: Exit Function
        End If
    Next

    ' 最大面積（VAS配下除外）
    bestArea = -1
    For Each c In pg.Controls
        If TypeName(c) = "TextBox" Then
            If Not IsUnderVAS(c) Then
                If c.Width * c.Height > bestArea Then
                    bestArea = c.Width * c.Height
                    Set best = c
                End If
            End If
        End If
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            Dim r3 As MSForms.TextBox
            Set r3 = FindNoteTextBox(c)
            If Not r3 Is Nothing Then
                If r3.Width * r3.Height > bestArea Then
                    bestArea = r3.Width * r3.Height
                    Set best = r3
                End If
            End If
        End If
    Next
    If Not best Is Nothing Then Set FindNoteTextBox = best
End Function

Private Function SafeIsMultiLine(tb As Object) As Boolean
    On Error Resume Next
    SafeIsMultiLine = CBool(tb.multiline)
    On Error GoTo 0
End Function

' fraVAS 配下かどうかを厳密判定（オブジェクト参照で再帰）
Private Function IsUnderVAS(target As Object) As Boolean
    Dim pg As Object, vas As Object
    Set pg = frmEval.Controls("mpPhys").Pages(4)
    On Error Resume Next
    Set vas = pg.Controls("fraVAS")
    On Error GoTo 0
    If vas Is Nothing Then Exit Function
    IsUnderVAS = IsDescendantOf(vas, target)
End Function

Private Function IsDescendantOf(container As Object, target As Object) As Boolean
    Dim c As Object
    For Each c In container.Controls
        If c Is target Then IsDescendantOf = True: Exit Function
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            If IsDescendantOf(c, target) Then IsDescendantOf = True: Exit Function
        End If
    Next
End Function


'=== [TEMP] NOTE Loader (置換版) =======================================
Private Sub LoadPainFromSheet_Note(ByVal owner As Object)
    Const COL_NOTE As Long = 108  ' HEADER_NOTE 列
    Dim ws As Worksheet, lr As Long, noteText As String
    Dim pg As Object, tb As MSForms.TextBox

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    noteText = CStr(ws.Cells(lr, COL_NOTE).value)

    Set pg = owner.Controls("mpPhys").Pages(4)
    If pg Is Nothing Then Exit Sub

    Set tb = FindNoteTextBox(pg)
    If tb Is Nothing Or IsUnderVAS(tb) Or tb.name = "txtPainDuration" Then Exit Sub
    If Not tb Is Nothing Then tb.Text = noteText

   
End Sub
'======================================================================



'=== [TEMP] Pain IO Loader (Finalize版) ===============================
Private Sub LoadPainFromSheet(ByVal owner As Object)





    Dim ws As Worksheet, lr As Long, ioText As String, noteText As String
    Dim pg As Object
    Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim hubRow As Long
hubRow = FindLatestRowByName(ws, Trim$(owner.txtName.Text))


ioText = ReadStr_Compat("IO_Pain", hubRow, ws)
'noteText = CStr(ws.Cells(hubRow, 108).value)


    Set pg = owner.Controls("mpPhys").Pages(4)
    If pg Is Nothing Then Exit Sub

   LoadPainFromSheet_MinCombos owner
Debug.Print "[After-MinCombos]", pg.Controls("txtPainDuration").Text

LoadPainFromSheet_MinLists owner
Debug.Print "[After-MinLists]", pg.Controls("txtPainDuration").Text

'LoadPainFromSheet_Note owner
Debug.Print "[After-Note]", pg.Controls("txtPainDuration").Text

Debug.Print "[After-PainCore-End]", frmEval.Controls("mpPhys").Pages(4).Controls("txtPainDuration").Text


End Sub
'======================================================================






'=== [TEMP] Latest row helper (IO/NOTE基準) ============================
Private Function LatestRowIO(ByVal ws As Worksheet) As Long
    LatestRowIO = WorksheetFunction.Max(ws.Cells(ws.rows.Count, 156).End(xlUp).row, ws.Cells(ws.rows.Count, 157).End(xlUp).row)

End Function
'======================================================================


'=== [TEMP] VAS単体読込デバッグ =======================================
Public Sub Debug_LoadVAS_FromLatest(ByVal owner As Object)
    Dim ws As Worksheet, lr As Long, s As String, t As String, alt As String
    Dim pg As Object
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = WorksheetFunction.Max(ws.Cells(ws.rows.Count, 156).End(xlUp).row, ws.Cells(ws.rows.Count, 157).End(xlUp).row)
    s = CStr(ws.Cells(lr, 156).value)
    t = IO_GetVal(s, "VAS")
    alt = CStr(ws.Cells(lr, 157).value)

    Set pg = owner.Controls("mpPhys").Pages(4)

    Debug.Print "[VAS-DBG] lr=", lr, "| IO.VAS=", t, "| NOTE=", alt

    ' まずクリア
    On Error Resume Next
    pg.Controls("fraVAS").Controls("txtVAS").Text = ""
    pg.Controls("fraVAS").Controls("sldVAS").value = 0
    On Error GoTo 0

    ' IOにあればそれを、無ければNOTE数値を適用
    If Len(t) = 0 And IsNumeric(alt) Then t = Trim$(alt)
    If Len(t) > 0 Then
        On Error Resume Next
        pg.Controls("fraVAS").Controls("txtVAS").Text = t
        pg.Controls("fraVAS").Controls("sldVAS").value = CLng(t)
        On Error GoTo 0
    End If

    Debug.Print "[VAS-DBG-After]", pg.Controls("fraVAS").Controls("txtVAS").Text, pg.Controls("fraVAS").Controls("sldVAS").value
End Sub
'======================================================================

'=== [TEMP] Pain UI Clear (起動時は空で開始) ===========================
Public Sub ClearPainUI(ByVal owner As Object)
    Dim pg As Object, c As Object, lb As MSForms.ListBox
    Set pg = owner.Controls("mpPhys").Pages(4)
    If pg Is Nothing Then Exit Sub

    ' --- Combo / Text ---
    On Error Resume Next
    pg.Controls("cmbPainOnset").value = ""
    pg.Controls("cmbPainDurationUnit").value = ""
    pg.Controls("cmbPainDayPeriod").value = ""
    pg.Controls("txtPainDuration").Text = ""
    On Error GoTo 0

    ' --- VAS ---
    On Error Resume Next
    pg.Controls("fraVAS").Controls("txtVAS").Text = ""
    pg.Controls("fraVAS").Controls("sldVAS").value = 0
    On Error GoTo 0

    ' --- ListBox 全解除 ---
    On Error Resume Next
    Set lb = pg.Controls("lstPainQual")
    If Not lb Is Nothing Then
        Dim i As Long: For i = 0 To lb.ListCount - 1: lb.Selected(i) = False: Next
    End If
    Set lb = pg.Controls("lstPainSite")
    If Not lb Is Nothing Then
        For i = 0 To lb.ListCount - 1: lb.Selected(i) = False: Next
    End If
    On Error GoTo 0


    ' [Pain-UI] ensure no default selection (DO NOT REMOVE)
On Error Resume Next
pg.Controls("lstPainQual").ListIndex = -1: pg.Controls("lstPainSite").ListIndex = -1
On Error GoTo 0



    ' --- Factors 全チェック解除（再帰）---
    ClearChecksRecursive pg
End Sub

Private Sub ClearChecksRecursive(container As Object)
    Dim c As Object
    For Each c In container.Controls
        If TypeName(c) = "CheckBox" Then c.value = False
        If TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then ClearChecksRecursive c
    Next
End Sub
'======================================================================



'=== [TEMP] 手動：最新行を即読込 ======================================
Public Sub LoadLatestPainNow()
    

    gPainLoadEnabled = True
    LoadPainFromSheet frmEval
    gPainLoadEnabled = False

End Sub




'======================================================================

Sub ExportAllVBA()
    Dim p As String, vbComp As Object, ext As String
    p = ThisWorkbook.path & "\vba_export"  ' 出力先フォルダ
    On Error Resume Next
    MkDir p
    On Error GoTo 0
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas"   ' vbext_ct_StdModule
            Case 2: ext = ".cls"   ' vbext_ct_ClassModule
            Case 3: ext = ".frm"   ' vbext_ct_MSForm
            Case Else: ext = ".txt"
        End Select
        vbComp.Export p & "\" & vbComp.name & ext
    Next
    Debug.Print "[Export]", p
End Sub


'=== LoadLatestSensoryNow（2025-10-22統合版）===
Public Sub LoadLatestSensoryNow(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_Sensory", ws)
    If r <= 0 Then
        Debug.Print "[LoadSensory] header not found"
        Exit Sub
    End If

    '--- 旧読込ロジックは後方互換のためコメントアウト ---
    'Call ParseSensoryData(ws.Cells(r, HeaderCol("IO_Sensory", ws)).Value)
    '------------------------------------------------------------

    ' 新ロジック：直接APIで読み込み
    Dim raw As String
    raw = LoadLatestSensoryNow_Raw(ws)
    Debug.Print "[LoadSensory] R=" & r & " Len=" & Len(raw) & " | " & Left$(raw, 60)
End Sub




Public Sub Temp_FindPainLoader()
    Dim vbc As Object
    Dim cm As Object
    Dim i As Long, ln As String, proc As String

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Set cm = vbc.CodeModule
        For i = 1 To cm.CountOfLines
            ln = cm.lines(i, 1)
            If InStr(1, ln, "LoadPainFromSheet", vbTextCompare) > 0 _
               Or InStr(1, ln, "IO_Pain", vbTextCompare) > 0 Then

                proc = cm.ProcOfLine(i, 0)
                Debug.Print "[PAIN]", vbc.name, "line=", i, "proc=", proc, " | ", ln
            End If
        Next i
    Next vbc
End Sub


