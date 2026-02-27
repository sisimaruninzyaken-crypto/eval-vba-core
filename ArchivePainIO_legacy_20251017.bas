Attribute VB_Name = "ArchivePainIO_legacy_20251017"
'=== modPainIO ===
Option Private Module

Option Explicit



'―― 設定（雛形の4点：必要に応じて後で差し替え可能） ―――――――――――――――――――――――――――――――
Private Const PAGE_HINT As String = "疼痛"            ' 対象タブの見出しの一部（例：「疼痛」「痛み」など）
Private Const HEADER_IO As String = "IO_Pain"         ' 本体のシリアライズ列
Private Const HEADER_NOTE As String = ""              ' 備考列：現行仕様では廃止
     ' 備考列（不要なら空文字にする）
Private keys As Variant                                ' R/Lペア化の論理キー。必要に応じて固定化可
' 例：後で必要なら Array("VAS","PainQual","PainCourse","PainSite","PainFactors","PainDuration")

' 区切り（テンプレ既定）
Private Const SEP_REC As String = "|"  ' レコード区切り
Private Const SEP_KV  As String = ":"  ' キーと値
Private Const SEP_RL  As String = ","  ' R/L 連結




'―― パブリックAPI（保存） ―――――――――――――――――――――――――――――――――――――――――――――
Public Sub SavePainToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)

    Dim pg As Object
    On Error Resume Next
    Set pg = owner.Controls.Item("mpPhys").Pages(4)   ' 疼痛（部位／NRS）
    On Error GoTo 0
    If pg Is Nothing Then
    
    
    Exit Sub
End If

    

    Dim combos As Collection
    Set combos = New Collection
    CollectCombos pg, combos  ' Page直下＋Frame内のComboBoxを再帰収集

    If combos.Count = 0 Then
       

    End If

    ' Top/Left ソート（TOL=6の簡易閾値：Top優先→Left）
    Dim arr() As Variant: arr = ControlsToArray(combos)
    SortByTopLeft arr, 6

    ' R/L ペアリングしてシリアライズ
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As String, base As String, side As String, valText As String

    For i = LBound(arr) To UBound(arr)
        nm = arr(i)("Name")
        valText = ComboValueText(arr(i)("Ref"))
        
        

        
        base = BaseNameRL(nm, side) ' side="R" or "L" or ""
        If Not dict.exists(base) Then dict.Add base, CreateObject("Scripting.Dictionary")
        If Len(side) = 0 Then side = "V" ' 単独項目（R/Lでない）
        dict(base)(side) = valText
        ' 任意：既知のキー順を使う場合は KEYS を後段で利用
        
      


        
    Next

    Dim parts As Collection: Set parts = New Collection
    Dim k As Variant, rec As String, vR As String, vL As String, vV As String

    ' 既定：検出順（Top/Left）で出力。固定順にしたい場合は KEYS を使用
    For Each k In dict.keys
        vR = NzS(dict(k), "R")
        vL = NzS(dict(k), "L")
        vV = NzS(dict(k), "V")
        If dict(k).exists("R") Or dict(k).exists("L") Then
            rec = CStr(k) & SEP_KV & " R=" & vR & SEP_RL & "L=" & vL
        Else
            rec = CStr(k) & SEP_KV & " " & vV
        End If
        parts.Add rec
    Next
    
        ' === 持続期間（数字：txtPainDuration）を単独キーとして保存 ===
    Dim durText As String
    On Error Resume Next
    durText = CStr(pg.Controls("txtPainDuration").Text)
    On Error GoTo 0
    If Len(Trim$(durText)) > 0 Then
        parts.Add "txtPainDuration" & SEP_KV & " " & durText
    End If

    

' === ListBox（複数選択）の保存を追加 ===
Dim c As Object, sel As String, j As Long, base2 As String
For Each c In pg.Controls
    If TypeName(c) = "ListBox" Then
        sel = ""
        For j = 0 To c.ListCount - 1
            If c.Selected(j) Then
                If Len(sel) > 0 Then sel = sel & "/"
                sel = sel & CStr(c.List(j))
            End If
        Next
        ' 何も選ばれていない場合は空のまま（出力しない）
        If Len(sel) > 0 Then
            ' 例：lstPainQual → PainQual, lstPainSite → PainSite
            base2 = c.name
            If LCase$(Left$(base2, 3)) = "lst" Then base2 = Mid$(base2, 4)
            
            'If base2 = "PainSite" Then sel = NormalizePainSite(sel)
            parts.Add base2 & SEP_KV & " " & sel
        End If
    End If
Next
' === 追加ここまで ===

' === CheckBox（Trueのみ）をまとめて保存 ===
Dim factors As Collection: Set factors = New Collection
CollectChecksRecursive pg, factors
If factors.Count > 0 Then
    Dim uniq As Object: Set uniq = CreateObject("Scripting.Dictionary")
    Dim ii As Long: For ii = 1 To factors.Count: uniq(factors(ii)) = 1: Next
    parts.Add "PainFactors" & SEP_KV & " " & Join(uniq.keys, "/")
End If




' === 追加ここまで ===

' === VAS（0 でも保存） ===
Dim vasText As String
On Error Resume Next
vasText = CStr(pg.Controls("fraVAS").Controls("txtVAS").Text)   ' TextBox 優先
If Len(vasText) = 0 Then vasText = CStr(pg.Controls("fraVAS").Controls("sldVAS").value)  ' ScrollBar 代替
On Error GoTo 0

' 「0」も有効値として保存する
If (Len(vasText) > 0) Or vasText = "0" Then
    parts.Add "VAS" & SEP_KV & " " & vasText
End If
' === 追加ここまで ===


    Dim outText As String: outText = JoinCollection(parts, SEP_REC)
    Debug.Print "[IO-FINAL]", outText


    ' 備考（Page内の最大TextBoxテキスト）
    Dim noteText As String: noteText = LargestTextBoxValue(pg)


    Debug.Print "[IO-FINAL]", outText

    ' 書き込み
    ws.Cells(r, EnsureHeaderCol(ws, HEADER_IO)).value = outText
    If LenB(HEADER_NOTE) > 0 Then
        ws.Cells(r, EnsureHeaderCol(ws, HEADER_NOTE)).value = noteText
    End If

   
    
   


    
    
    
End Sub

'―― 補助：対象MultiPageとPage探索 ―――――――――――――――――――――――――――――――――
Private Function FindTargetMultiPage(ByVal owner As Object, ByVal hint As String, ByRef outPage As Object) As Object
    Dim ctl As Object, mp As Object, i As Long
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For i = 0 To mp.Pages.Count - 1
                If InStr(1, mp.Pages(i).caption, hint, vbTextCompare) > 0 Then
                    Set outPage = mp.Pages(i)
                    Set FindTargetMultiPage = mp
                    Exit Function
                End If
            Next
        End If
    Next
    Set outPage = Nothing
    Set FindTargetMultiPage = Nothing
End Function

'―― 補助：Page配下のComboBoxを再帰収集（Frame内含む） ―――――――――――――――――――――――――
Private Sub CollectCombos(ByVal container As Object, ByRef bag As Collection)
    Dim ctl As Object
    For Each ctl In container.Controls
        Select Case TypeName(ctl)
            Case "ComboBox": bag.Add ctl
            Case "Frame":    CollectCombos ctl, bag
        End Select
    Next
End Sub

'―― 補助：Combos→配列（Name/Top/Left/Ref） ―――――――――――――――――――――――――――――――
Private Function ControlsToArray(ByVal bag As Collection) As Variant
    Dim i As Long, o As Object
    Dim arr() As Variant
    ReDim arr(0 To bag.Count - 1)
    For i = 1 To bag.Count
        Set o = bag(i)
       Set arr(i - 1) = CreateMap4("Name", o.name, "Top", CLng(o.Top), "Left", CLng(o.Left), "Ref", o)

    Next
    ControlsToArray = arr
End Function

Private Sub SortByTopLeft(ByRef arr As Variant, ByVal tol As Long)
   Dim i As Long, j As Long, tmp As Object
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If (Abs(arr(i)("Top") - arr(j)("Top")) > tol And arr(i)("Top") > arr(j)("Top")) _
               Or (Abs(arr(i)("Top") - arr(j)("Top")) <= tol And arr(i)("Left") > arr(j)("Left")) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'―― 補助：R/L 基底名抽出 ―――――――――――――――――――――――――――――――――――――
Private Function BaseNameRL(ByVal name As String, ByRef side As String) As String
    Dim s As String: s = UCase$(name)
    If Right$(s, 2) = "_R" Then side = "R": BaseNameRL = Left$(name, Len(name) - 2): Exit Function
    If Right$(s, 2) = "_L" Then side = "L": BaseNameRL = Left$(name, Len(name) - 2): Exit Function
    side = "": BaseNameRL = name
End Function

'―― 補助：Comboの値取得（Style=2対策：Value→ListIndex→AddItem） ―――――――――――――――――
Private Function ComboValueText(ByVal cbo As Object) As String
    On Error Resume Next
    Dim t As String
    t = CStr(cbo.value)
    If LenB(t) = 0 Then
        If cbo.ListIndex >= 0 Then t = CStr(cbo.List(cbo.ListIndex))
    End If
    If LenB(t) = 0 Then
        ' 未登録値が入っているケースに備えてTextも見る
        t = CStr(cbo.Text)
    End If
    ComboValueText = t
End Function

'―― 補助：最大TextBoxの値（備考想定） ―――――――――――――――――――――――――――――
Private Function LargestTextBoxValue(ByVal container As Object) As String
    Dim ctl As Object, area As Double, maxArea As Double, best As Object
    For Each ctl In container.Controls
        If TypeName(ctl) = "TextBox" Then
            area = ctl.Width * ctl.Height
            If area > maxArea Then maxArea = area: Set best = ctl
        ElseIf TypeName(ctl) = "Frame" Then
            Dim s As String: s = LargestTextBoxValue(ctl)
            If LenB(s) > 0 And area = 0 Then LargestTextBoxValue = s ' ネスト側で決まったら採用
        End If
    Next
    If Not best Is Nothing Then LargestTextBoxValue = CStr(best.Text)
End Function

'―― 小物ユーティリティ ―――――――――――――――――――――――――――――――――――――
Private Function NzS(ByVal dict As Object, ByVal k As String) As String
    If dict.exists(k) Then NzS = CStr(dict(k)) Else NzS = ""
End Function

Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, s() As String
    ReDim s(1 To c.Count)
    For i = 1 To c.Count: s(i) = CStr(c(i)): Next
    JoinCollection = Join(s, sep)
End Function

Private Function CreateMap4(k1 As String, v1 As Variant, k2 As String, v2 As Variant, k3 As String, v3 As Variant, k4 As String, v4 As Variant) As Variant
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add k1, v1: d.Add k2, v2: d.Add k3, v3: d.Add k4, v4
    Set CreateMap4 = d
End Function






Private Sub CollectChecksRecursive(parent As Object, coll As Collection)
    Dim c As Object, nm2 As String
    For Each c In parent.Controls
        If TypeName(c) = "CheckBox" Then
            If c.value = True Then
                nm2 = c.name
                If LCase$(Left$(nm2, 3)) = "chk" Then nm2 = Mid$(nm2, 4)
                coll.Add nm2
            End If
        ElseIf TypeName(c) = "Frame" Or TypeName(c) = "Page" Then
            CollectChecksRecursive c, coll
        End If
    Next
End Sub



Public Sub DumpPainFrames_Once()
    Dim pg As Object, f As Object, c As Object
    Dim uf As Object: Set uf = frmEval
Set pg = uf.mpPhys.Pages(4)

    Debug.Print "[Page]", pg.caption
    For Each f In pg.Controls
        If TypeName(f) = "Frame" Then
            Debug.Print "[Frame]", f.name, "count", f.Controls.Count
            For Each c In f.Controls
                If TypeName(c) = "CheckBox" Then Debug.Print "  [Chk]", c.name, c.value
                If TypeName(c) = "Frame" Then Debug.Print "  [SubFrame]", c.name, "count", c.Controls.Count
            Next
        End If
    Next
End Sub



Public Sub SavePain_CheckOnce()
    Dim uf As Object: Set uf = frmEval
    ' チェック1つON（誘因：動作で増悪）
    uf.Controls("mpPhys").Pages(4).Controls("fraPainFactors").Controls("chkPainProv_Move").value = True
    ' 保存（ボタン相当）
    SaveEvaluation_Append_From uf
    ' 直近行のIO/NOTEを数値表示
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub



Public Sub SavePain_AppendTest_Once()
    Dim uf As Object: Set uf = frmEval
    On Error Resume Next
    uf.txtName.Text = "検証Append"
    uf.Controls("chkDiffOnly").value = False
    On Error GoTo 0

    ' 保存（全体保存ルート）
    SaveEvaluation_Append_From uf

    ' 直近行のIO/NOTEを可視化
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub


Public Sub Test_SaveAtRow_Once()
    Dim ws As Worksheet, rr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    rr = 108   ' ← 任意の検証行

    'SaveAllSectionsToSheet ws, rr, frmEval

    Debug.Print "[WroteRow]", rr
   'Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 107).value), 180)
   'Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 108).value), 120)

End Sub



Public Sub SavePain_FillAndAppend_Once()
    Dim uf As Object: Set uf = frmEval
    On Error Resume Next
    uf.txtName.Text = "検証Append3"
    uf.Controls("chkDiffOnly").value = False
    With uf.Controls("mpPhys").Pages(4)
        .Controls("cmbPainOnset").ListIndex = 0
        .Controls("cmbPainDurationUnit").ListIndex = 0
        .Controls("cmbPainDayPeriod").ListIndex = 2
        Dim lb As MSForms.ListBox
        Set lb = .Controls("lstPainQual"): If lb.ListCount > 0 Then lb.Selected(0) = True
        Set lb = .Controls("lstPainSite"): If lb.ListCount > 0 Then lb.Selected(0) = True
        .Controls("fraPainFactors").Controls("chkPainProv_Move").value = True
    End With
    On Error GoTo 0

    SaveEvaluation_Append_From uf

    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub





Private Function GetCtlVal(o As Object) As String
    On Error Resume Next
    GetCtlVal = "" & o.value
    If Len(GetCtlVal) = 0 Then GetCtlVal = "" & o.Text
    On Error GoTo 0
End Function





'=== [TEMP] Pain IO One-Shot ===========================================
Public Sub SaveAndReloadLatest(ByVal owner As Object)
    SaveEvaluation_Append_From owner
    gPainLoadEnabled = True
   LoadLatestPainNow
    gPainLoadEnabled = False

End Sub

'======================================================================

'=== [TEMP] 保存系ハンドラ特定スキャナ ================================
Public Sub Scan_SaveHandlers()
    Dim vbComp As Object, cm As Object
    Dim lineCount As Long, i As Long
    Dim pat1 As String, pat2 As String, pat3 As String
    pat1 = "保存"         ' メッセージ文言
    pat2 = "Save"         ' サブルーチン名キーワード
    pat3 = "Application.Caller" ' シートボタン想定

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set cm = vbComp.CodeModule
        lineCount = cm.CountOfLines
        For i = 1 To lineCount
            Dim txt As String: txt = cm.lines(i, 1)
            If (InStr(txt, pat1) > 0) Or (InStr(txt, pat2) > 0) Or (InStr(txt, pat3) > 0) Then
                Debug.Print "[HIT]", vbComp.name, "L", i, "|"; Left$(Trim$(txt), 180)
            End If
        Next i
    Next vbComp
    Debug.Print "[Scan] done."
End Sub
'=====================================================================




