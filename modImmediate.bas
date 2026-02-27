Attribute VB_Name = "modImmediate"


'--- 補助（このモジュール内だけで完結） ---
Private Function NzLng(v As Variant) As Long
    If IsNumeric(v) Then NzLng = CLng(v) Else NzLng = 0
End Function

Private Function NormName(ByVal s As String) As String
    s = Replace(s, vbCrLf, "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    On Error Resume Next
    s = StrConv(s, vbNarrow) ' 全角→半角（可能なら）
    On Error GoTo 0
    NormName = LCase$(s)
End Function

Private Sub DumpLook_Simple(ByVal look As Object, ByVal ws As Worksheet)
    Dim k As Variant, c As Variant
    For Each k In Array("Basic.Name", "Basic.ID", "Basic.Age", "Basic.EvalDate")
        If look.exists(k) Then
            c = NzLng(look(k))
            If c > 0 Then
                Debug.Print "[LOOK]", k, "Col", c, "hdr=""" & CStr(ws.Cells(1, c).value) & """"
            Else
                Debug.Print "[LOOK][MISS]", k
            End If
        Else
            Debug.Print "[LOOK][NONE]", k
        End If
    Next
End Sub

Private Sub DumpCandidate_Simple(ByVal ws As Worksheet, ByVal look As Object, ByVal rowNum As Long, ByVal idx As Long)
    Dim nm$, pid$, age$, dt$
    If NzLng(look("Basic.Name")) > 0 Then nm = CStr(ws.Cells(rowNum, look("Basic.Name")).value)
    If NzLng(look("Basic.ID")) > 0 Then pid = CStr(ws.Cells(rowNum, look("Basic.ID")).value)
    If NzLng(look("Basic.Age")) > 0 Then age = CStr(ws.Cells(rowNum, look("Basic.Age")).value)
    If NzLng(look("Basic.EvalDate")) > 0 Then dt = CStr(ws.Cells(rowNum, look("Basic.EvalDate")).value)
    Debug.Print "  [" & idx & "] row=" & rowNum & "  name=""" & nm & """  ID=" & pid & "  age=" & age & "  date=" & dt
End Sub
'=== ここまで ===



Public Sub RunEnsureCols()
    Call modEvalIOEntry.EnsureHeaderCol_BasicInfo(modSchema.GetEvalDataSheet())
End Sub


Public Sub Test_HeaderCheck()
    Dim ws As Worksheet, look As Object
    Set ws = modSchema.GetEvalDataSheet()
    Set look = BuildHeaderLookup(ws)

   Debug.Print "補助具:", modEvalIOEntry.FindColByHeaderExact(ws, "補助具")
Debug.Print "リスク:", modEvalIOEntry.FindColByHeaderExact(ws, "リスク")


End Sub


  ' ROMページ：数値欄=Label+22px、備考=Label+28px を“絶対座標”で揃える（親違いOK）
Public Sub ROM_AlignFix_Set20()
    Const targetSmall As Long = 24
    Const targetNote  As Long = 28

    Dim c As Object, mp As Object, pg As Object, i As Long
    Dim bestGap As Double

    ' --- ROMページ特定 ---
    For Each c In frmEval.Controls
        If TypeName(c) = "MultiPage" Then
            For i = 0 To c.Pages.Count - 1
                If InStr(1, CStr(c.Pages(i).caption), "ROM", vbTextCompare) > 0 _
                Or InStr(1, CStr(c.Pages(i).caption), "主要関節", vbTextCompare) > 0 Then
                    Set mp = c: Set pg = c.Pages(i): Exit For
                End If
            Next
            If Not pg Is Nothing Then Exit For
        End If
    Next
    If pg Is Nothing Then Debug.Print "[ROM_AlignFix_Set20] ROM page not found": Exit Sub

    Dim ctrl As Object, txt As Object, lbl As Object, tmp As Object, tmpZ As Object
    Dim ml As Boolean, isBig As Boolean
    Dim desired As Single, curr As Single
    Dim adjSmall As Long, adjNote As Long, skip As Long

    For Each ctrl In pg.Controls
        If TypeName(ctrl) = "TextBox" Then
            Set txt = ctrl

            ' --- 備考候補の判定 ---
            ml = False: On Error Resume Next: ml = txt.multiline: On Error GoTo 0
            isBig = (ml Or txt.Height >= 80 Or txt.Width >= 400 _
                    Or InStr(1, UCase$(txt.name), "NOTE") > 0 _
                    Or InStr(1, UCase$(txt.tag & ""), "NOTE") > 0)

            ' --- ラベル特定 ---
            Set lbl = Nothing
            If isBig Then
                ' ① 親内の「備考」ラベル
                For Each tmp In txt.parent.Controls
                    If TypeName(tmp) = "Label" Then
                        If InStr(1, CStr(tmp.caption), "備考", vbTextCompare) > 0 Then Set lbl = tmp: Exit For
                    End If
                Next tmp
                ' ② 見つからなければページ全体の「備考」ラベル
                If lbl Is Nothing Then
                    For Each tmpZ In pg.Controls
                        If TypeName(tmpZ) = "Label" Then
                            If InStr(1, CStr(tmpZ.caption), "備考", vbTextCompare) > 0 Then Set lbl = tmpZ: Exit For
                        End If
                    Next tmpZ
                End If
            End If
            ' ③-補：親内で見つからなければ、ページ全体で最近左（<=）ラベルを探す
If lbl Is Nothing Then
    bestGap = 1E+20
    Dim d As Double   ' ← ②で説明する d の宣言。既に上で宣言していれば不要
    For Each tmp In txt.parent.Controls
        If TypeName(tmp) = "Label" Then
            If tmp.Left <= txt.Left And tmp.Width <= 120 _
               And (LenB(CStr(tmp.caption)) >= 2 Or tmp.Width >= 12) Then

                d = AbsTop(txt) - AbsTop(tmp)  ' 上にある距離だけ採用
                If d >= 0 And d <= 120 Then
                    If d < bestGap Then Set lbl = tmp: bestGap = d
                End If
            End If
        End If
    Next tmp
End If



            ' --- 揃え ---
            If Not lbl Is Nothing Then
              desired = AbsTop(lbl) + IIf(isBig, targetNote, targetSmall)


                curr = AbsTop(txt)
                If Abs(curr - desired) > 0.5 Then
                    txt.Top = txt.Top + (desired - curr)
                    If isBig Then adjNote = adjNote + 1 Else adjSmall = adjSmall + 1
                Else
                    skip = skip + 1
                End If
            Else
                skip = skip + 1
            End If
        End If
    Next ctrl

    

    Debug.Print "[ROM_AlignFix_Set20] AdjustedSmall=" & adjSmall & "  AdjustedNote=" & adjNote & "  Skipped=" & skip
End Sub






' 備考だけを確実に整列：Label(「備考」)の24px下へ
Public Sub ROM_NoteFix_Once()
    Const target As Long = 24
    Dim c As Object, mp As Object, pg As Object, i As Long
    Dim ctrl As Object, lbl As Object, tb As Object
    Dim ml As Boolean, oldTop As Single

    ' ROMページ特定
    For Each c In frmEval.Controls
        If TypeName(c) = "MultiPage" Then
            Set mp = c
            For i = 0 To mp.Pages.Count - 1
                If InStr(1, CStr(mp.Pages(i).caption), "ROM", vbTextCompare) > 0 _
                Or InStr(1, CStr(mp.Pages(i).caption), "主要関節", vbTextCompare) > 0 Then
                    Set pg = mp.Pages(i): Exit For
                End If
            Next
            If Not pg Is Nothing Then Exit For
        End If
    Next
    If pg Is Nothing Then Debug.Print "[NoteFix] ROM page not found": Exit Sub

    ' 「備考」ラベル
    For Each ctrl In pg.Controls
        If TypeName(ctrl) = "Label" Then
            If InStr(1, CStr(ctrl.caption), "備考", vbTextCompare) > 0 Then Set lbl = ctrl: Exit For
        End If
    Next
    If lbl Is Nothing Then Debug.Print "[NoteFix] 備考ラベルなし": Exit Sub

    ' 最大サイズのMultiLineテキスト（備考本体）
    For Each ctrl In pg.Controls
        If TypeName(ctrl) = "TextBox" Then
            ml = False: On Error Resume Next: ml = ctrl.multiline: On Error GoTo 0
            If ml Or ctrl.Height >= 80 Or ctrl.Width >= 400 Then
                If tb Is Nothing Then
                    Set tb = ctrl
                ElseIf ctrl.Height * ctrl.Width > tb.Height * tb.Width Then
                    Set tb = ctrl
                End If
            End If
        End If
    Next
    If tb Is Nothing Then Debug.Print "[NoteFix] 備考テキストなし": Exit Sub

    oldTop = tb.Top
    tb.Top = lbl.Top + target
    Debug.Print "[NoteFix] Top: " & oldTop & " -> " & tb.Top & "  (Label.Top=" & lbl.Top & ", Δ=" & (tb.Top - lbl.Top) & ")"
End Sub



' 親チェーンを最大20段までで打ち切る安全版
Private Function AbsTop(ByVal o As Object) As Single
    On Error Resume Next
    Dim t As Single: t = NzTop(o)
    Dim p As Object: Set p = GetParentSafe(o)
    Dim guard As Long
    Do While Not p Is Nothing And guard < 20
        t = t + NzTop(p)
        Set p = GetParentSafe(p)
        guard = guard + 1
    Loop
    AbsTop = t
End Function

Private Function NzTop(ByVal o As Object) As Single
    On Error Resume Next
    NzTop = o.Top
End Function

Private Function GetParentSafe(ByVal o As Object) As Object
    On Error Resume Next
    Set GetParentSafe = o.parent
End Function



























' ROMページ内の全CheckBoxを「元位置から非累積で12px上」に配置（高さ/フォントは触らない）
Public Sub ROM_CheckBoxes_Up12_OnROM_Recursive_Once_V2()
    Dim mp As Object, pg As Object, i As Long, found As Boolean
    Dim stk As Collection, cont As Object, c As Object
    Dim j As Long, baseTop As Double, pos As Long
    Dim tagKey As String: tagKey = "CBBase="

    ' ?? ROMページを特定（Captionに ROM / ＲＯＭ / 主要関節 / 関節可動域 を含む）??
    For Each mp In frmEval.Controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.Count - 1
                Set pg = mp.Pages(i)
                If (InStr(1, CStr(pg.caption), "ROM", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "ＲＯＭ", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "主要関節", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "関節可動域", vbTextCompare) > 0) Then
                    found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then MsgBox "ROMページが見つかりません。", vbExclamation: Exit Sub

    ' ?? 入れ子(Frame, Page, MultiPage内のPage)も含めて走査（VBA Collection スタック）??
    Set stk = New Collection
    stk.Add pg

    Do While stk.Count > 0
        Set cont = stk(stk.Count): stk.Remove stk.Count

        On Error Resume Next
        For Each c In cont.Controls
            ' 子コンテナはスタックに積む
            Select Case TypeName(c)
                Case "Frame", "Page"
                    stk.Add c
                Case "MultiPage"
                    For j = 0 To c.Pages.Count - 1
                        stk.Add c.Pages(j)
                    Next
            End Select

            ' CheckBoxだけ Top を「元Top?12」に（非累積／高さは触らない）
            If TypeName(c) = "CheckBox" Then
                With c
                    pos = InStr(1, .tag, tagKey, vbTextCompare)
                    If pos > 0 Then
                        baseTop = val(Mid$(.tag, pos + Len(tagKey)))
                    Else
                        baseTop = .Top
                        If Len(.tag) > 0 Then
                            .tag = .tag & "|" & tagKey & CStr(baseTop)
                        Else
                            .tag = tagKey & CStr(baseTop)
                        End If
                    End If
                    .Top = baseTop - 12
                End With
            End If
        Next
        On Error GoTo 0
    Loop
End Sub







' ROMタブの最終検証：TextBox高さ≠20 と、CheckBoxが「元Top?12」からズレている件数を表示
Public Sub ROM_VerifyOnce()
    Dim mp As Object, pg As Object, c As Object, i As Long, found As Boolean
    Dim ngTB As Long, ngCB As Long, base As Double, pos As Long, cap As String

    ' ROMページ特定
    For Each mp In frmEval.Controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.Count - 1
                cap = CStr(mp.Pages(i).caption)
                If InStr(1, cap, "ROM", vbTextCompare) > 0 Or _
                   InStr(1, cap, "ＲＯＭ", vbTextCompare) > 0 Or _
                   InStr(1, cap, "主要関節", vbTextCompare) > 0 Or _
                   InStr(1, cap, "関節可動域", vbTextCompare) > 0 Then
                    Set pg = mp.Pages(i): found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then Debug.Print "[VERIFY] ROM page not found": Exit Sub

    ' 検証
    For Each c In pg.Controls
        If TypeName(c) = "TextBox" Then
            If c.multiline = False And c.Height <> 15 Then ngTB = ngTB + 1
        ElseIf TypeName(c) = "CheckBox" Then
            pos = InStr(1, c.tag, "CBBase=", vbTextCompare)
            If pos > 0 Then
                base = val(Mid$(c.tag, pos + 7))
                If Abs((base - 12) - c.Top) > 0.5 Then ngCB = ngCB + 1
            Else
                ngCB = ngCB + 1  ' 基準未記録もズレ扱い
            End If
        End If
    Next

    
End Sub




' ROMページ内の全ての単行TextBoxを高さ20に固定（Frameや入れ子も含めて再帰）
Public Sub ROM_Fix_TextBoxHeight_Recursive_OnROM_Once()
    Const h As Single = 15
    Dim mp As Object, pg As Object, i As Long, found As Boolean

    ' ROMページを特定
    For Each mp In frmEval.Controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.Count - 1
                Set pg = mp.Pages(i)
                If InStr(1, CStr(pg.caption), "ROM", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "ＲＯＭ", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "主要関節", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "関節可動域", vbTextCompare) > 0 Then
                    found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then MsgBox "ROMページが見つかりません。": Exit Sub

    ' 再帰処理開始
    Call FixTextBoxHeightRecursive(pg, h)
End Sub

Private Sub FixTextBoxHeightRecursive(cont As Object, h As Single)
    Dim c As Control, j As Long
    On Error Resume Next
    For Each c In cont.Controls
        Select Case TypeName(c)
            Case "Frame", "Page"
                Call FixTextBoxHeightRecursive(c, h)
            Case "MultiPage"
                For j = 0 To c.Pages.Count - 1
                    Call FixTextBoxHeightRecursive(c.Pages(j), h)
                Next
            Case "TextBox"
                If c.multiline = False Then c.Height = h
        End Select
    Next
    On Error GoTo 0
End Sub




'--- 対象ページ（筋力/MMT）の特定 ---
Private Function GetMMTPage() As Object
    Dim c As Object, mp As Object, i As Long
    For Each c In frmEval.Controls
        If TypeName(c) = "MultiPage" Then
            Set mp = c
            For i = 0 To mp.Pages.Count - 1
                Dim cap As String
                cap = mp.Pages(i).caption
                If InStr(cap, "MMT") > 0 Or InStr(cap, "筋力") > 0 Then
                    Set GetMMTPage = mp.Pages(i)
                    Exit Function
                End If
            Next
        End If
    Next
    Set GetMMTPage = Nothing
End Function

'--- 自動生成タグのものだけ削除 ---
Private Sub MMT_ClearGen(pg As Object)
    Dim idx As Long
    For idx = pg.Controls.Count - 1 To 0 Step -1
        If Left$(pg.Controls(idx).tag & "", 6) = "MMTGEN" Then
            pg.Controls.Remove pg.Controls(idx).name
        End If
    Next
End Sub

'--- 1ページぶんのレイアウト生成 ---
Private Sub BuildPage(pg As Object, items As Variant)
    Const ROW_H As Single = 24
    Const LBL_W As Single = 130
    Const COL_W As Single = 90
    Const gap As Single = 12

    Dim x0 As Single, y0 As Single
    x0 = 20: y0 = 28

    ' 見出し
    MakeLabel pg, "lblHdr_Muscle", "筋群", x0, y0 - 20, 60, 18
    MakeLabel pg, "lblHdr_R", "右", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLabel pg, "lblHdr_L", "左", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, Y As Single
    Y = y0
    For i = LBound(items) To UBound(items)
        Dim key As String: key = CStr(items(i))

        MakeLabel pg, "lbl_" & key, key, x0, Y + 3, LBL_W, 18
        MakeCombo pg, "cboR_" & key, x0 + LBL_W + gap, Y, COL_W, 18
        MakeCombo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, Y, COL_W, 18

        Y = Y + ROW_H
    Next i
End Sub

'--- Label 生成 ---
Private Sub MakeLabel(pg As Object, nm As String, cap As String, l As Single, t As Single, w As Single, h As Single)
    Dim o As MSForms.label
    Set o = pg.Controls.Add("Forms.Label.1", nm, True)
    With o
        .caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .tag = "MMTGEN"
    End With
End Sub

'--- ComboBox 生成（MMT 0～5） ---
Private Sub MakeCombo(pg As Object, nm As String, l As Single, t As Single, w As Single, h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.Controls.Add("Forms.ComboBox.1", nm, True)
    With o
        .Left = l: .Top = t: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .tag = "MMTGEN"
        .BoundColumn = 1
        .List = Split("0,1,2,3,4,5", ",")
    End With
End Sub



'=== MMT 兄弟ページ方式：MMT_上肢 / MMT_下肢 を追加して項目生成 ===
Public Sub MMT_BuildSiblingTabs()
    Dim mp As Object, iMMT As Long
    If Not FindMMT_MultiPage(mp, iMMT) Then
        MsgBox "MMT/筋力 ページが見つかりません。", vbExclamation: Exit Sub
    End If

    Dim pgUpper As Object, pgLower As Object
    Set pgUpper = EnsurePage(mp, "MMT_上肢", iMMT + 1)
    Set pgLower = EnsurePage(mp, "MMT_下肢", iMMT + 2)

    Dim upperItems, lowerItems
    upperItems = Array("肩屈曲", "肩伸展", "肩外転", "肩内旋", "肩外旋", "肘屈曲", "肘伸展", "前腕回内", "前腕回外", "手関節掌屈", "手関節背屈", "指屈曲", "指伸展", "母指対立")
    lowerItems = Array("股屈曲", "股伸展", "股外転", "股内転", "膝屈曲", "膝伸展", "足関節背屈", "足関節底屈", "母趾伸展")

    ClearGenerated pgUpper: ClearGenerated pgLower
    BuildMMTPage pgUpper, upperItems
    BuildMMTPage pgLower, lowerItems

    MsgBox "MMT_上肢／MMT_下肢 を作成しました。", vbInformation
End Sub

'--- MMTページを含む MultiPage とそのインデックスを取得 ---
Private Function FindMMT_MultiPage(ByRef mp As Object, ByRef idx As Long) As Boolean
    Dim c As Object, i As Long
    For Each c In frmEval.Controls
        If TypeName(c) = "MultiPage" Then
            For i = 0 To c.Pages.Count - 1
                Dim cap$: cap = c.Pages(i).caption
                If InStr(cap, "MMT") > 0 Or InStr(cap, "筋力") > 0 Then
                    Set mp = c: idx = i: FindMMT_MultiPage = True: Exit Function
                End If
            Next
        End If
    Next
End Function

'--- 指定位置にページを用意（あれば再利用） ---
Private Function EnsurePage(mp As Object, title As String, atIndex As Long) As Object
    Dim i As Long
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = title Then Set EnsurePage = mp.Pages(i): Exit Function
    Next
    Dim pg As Object
    Set pg = mp.Pages.Add
    pg.caption = title
    If atIndex >= 0 And atIndex < mp.Pages.Count Then mp.Pages(mp.Pages.Count - 1).Index = atIndex
    Set EnsurePage = pg
End Function

'--- 自動生成の掃除 ---
Private Sub ClearGenerated(pg As Object)
    Dim j As Long
    For j = pg.Controls.Count - 1 To 0 Step -1
        If Left$(pg.Controls(j).tag & "", 6) = "MMTGEN" Then pg.Controls.Remove pg.Controls(j).name
    Next
End Sub

'--- MMT 1ページ分のUI生成 ---
Private Sub BuildMMTPage(pg As Object, items As Variant)
    Const ROW_H As Single = 24, LBL_W As Single = 130, COL_W As Single = 90, gap As Single = 12
    Dim x0 As Single, y0 As Single: x0 = 20: y0 = 28

    MakeLbl pg, "lblHdrMus", "筋群", x0, y0 - 20, 60, 18
    MakeLbl pg, "lblHdrR", "右", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLbl pg, "lblHdrL", "左", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, Y As Single: Y = y0
    For i = LBound(items) To UBound(items)
        Dim key$: key = CStr(items(i))
        MakeLbl pg, "lbl_" & key, key, x0, Y + 3, LBL_W, 18
        MakeCbo pg, "cboR_" & key, x0 + LBL_W + gap, Y, COL_W, 18
        MakeCbo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, Y, COL_W, 18
        Y = Y + ROW_H
    Next
End Sub

Private Sub MakeLbl(pg As Object, nm$, cap$, l!, t!, w!, h!)
    Dim o As MSForms.label
    Set o = pg.Controls.Add("Forms.Label.1", nm, True)
    o.caption = cap: o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.tag = "MMTGEN"
End Sub

Private Sub MakeCbo(ByVal pg As Object, ByVal nm As String, _
                    ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.Controls.Add("Forms.ComboBox.1", nm, True)
    With o
        .Left = l: .Top = t: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .BoundColumn = 1
        .List = Split("0,1,2,3,4,5", ",")
        .tag = "MMTGEN"
    End With
End Sub





'=== MMTの子タブ（上肢／下肢）を、MMTページ内のFrameに収容して作成 ===
Public Sub MMT_BuildChildTabs_Frame()
    Dim pg As Object
    Set pg = GetMMTPage()
    If pg Is Nothing Then MsgBox "MMTページが見つかりません。", vbExclamation: Exit Sub

    Dim fra As MSForms.Frame, mp As MSForms.MultiPage
    On Error Resume Next
    Set fra = pg.Controls("fraMMTWrap")
    On Error GoTo 0
    If fra Is Nothing Then
        ' Frameをフォームに追加してから親をMMTページへ
        Set fra = frmEval.Controls.Add("Forms.Frame.1", "fraMMTWrap", True)
        Set fra.parent = pg
        With fra
            .caption = ""
            .Left = 12: .Top = 12
            .Width = 820: .Height = 420
            .tag = "MMTGEN"
        End With
    End If

    On Error Resume Next
    Set mp = fra.Controls("mpMMTChild")
    On Error GoTo 0
    If mp Is Nothing Then
        Set mp = fra.Controls.Add("Forms.MultiPage.1", "mpMMTChild", True)
        With mp
            .Left = 6: .Top = 6
            .Width = fra.Width - 12
            .Height = fra.Height - 12
            .Style = 0
            .TabsPerRow = 4
            .Pages.Clear
            .Pages.Add.caption = "上肢"
            .Pages.Add.caption = "下肢"
        End With
    End If

    ' 中身を構築
    MMT_ClearGen mp.Pages(0)
    MMT_ClearGen mp.Pages(1)
    BuildPage mp.Pages(0), Array("肩屈曲", "肩伸展", "肩外転", "肩内旋", "肩外旋", "肘屈曲", "肘伸展", "前腕回内", "前腕回外", "手関節掌屈", "手関節背屈", "指屈曲", "指伸展", "母指対立")
    BuildPage mp.Pages(1), Array("股屈曲", "股伸展", "股外転", "股内転", "膝屈曲", "膝伸展", "足関節背屈", "足関節底屈", "母趾伸展")

    MsgBox "MMTページ内に子タブ（上肢／下肢）を作成しました。", vbInformation
End Sub





Public Sub Swap_DailyLogList_ToMonthlyDraftBox()
    On Error GoTo EH

    Dim uf As frmEval
    Dim lb As MSForms.Control
    Dim host As Object          ' fraDailyLog
    Dim tb As MSForms.Control   ' TextBox

    ' 重要：New は使わない（Initializeで落ちる環境があるため）
    Set uf = frmEval            ' 起動中のインスタンスを参照

    Set lb = uf.Controls("lstDailyLogList")
    Set host = lb.parent        ' fraDailyLog のはず

    ' 既に作ってあればそれを使う
    On Error Resume Next
    Set tb = host.Controls("txtMonthlyMonitoringDraft")
    On Error GoTo EH

    If tb Is Nothing Then
        Set tb = host.Controls.Add("Forms.TextBox.1", "txtMonthlyMonitoringDraft", True)
    End If

    ' 位置とサイズを lstDailyLogList に合わせる
    tb.Left = lb.Left
    tb.Top = lb.Top
    tb.Width = lb.Width
    tb.Height = lb.Height

    ' 使い勝手の最小限（WordWrap等は触らない）
    tb.multiline = True
    tb.EnterKeyBehavior = True
    tb.ScrollBars = fmScrollBarsVertical

    ' ListBox は隠す（戻したい時に戻せる）
    lb.Visible = False
    tb.Visible = True

#If APP_DEBUG Then
    Debug.Print "[Swap] lb.Visible=" & lb.Visible, _
                "tb.Name=" & tb.name, _
                "L=" & tb.Left & " T=" & tb.Top & " W=" & tb.Width & " H=" & tb.Height
#End If

    Exit Sub

EH:
#If APP_DEBUG Then
    Debug.Print "[Swap][ERR]", Err.Number, Err.Description
#End If
End Sub


Public Sub Ensure_MonthlyDraftBox_UnderFraDailyLog()
    On Error GoTo EH

    Dim uf As Object
    Dim f As Object
    Dim lb As Object
    Dim tb As Object

    Set uf = VBA.UserForms(0)
    Set f = uf.Controls("fraDailyLog")
    Set lb = f.Controls("lstDailyLogList")

    ' 既存があれば取得、なければ作成（fraDailyLog 直下）
    On Error Resume Next
    Set tb = f.Controls("txtMonthlyMonitoringDraft")
    On Error GoTo EH

    If tb Is Nothing Then
        Set tb = f.Controls.Add("Forms.TextBox.1", "txtMonthlyMonitoringDraft", True)
    End If

    ' ListBoxと同じ矩形に合わせる（確定値）
    tb.Left = 12
    tb.Top = 294
    tb.Width = 987.3
    tb.Height = 140

    tb.multiline = True
    tb.EnterKeyBehavior = True
    tb.ScrollBars = fmScrollBarsVertical

    lb.Visible = False
    tb.Visible = True

    Debug.Print "[EnsureDraftBox] OK name=" & tb.name, _
                "L=" & tb.Left & " T=" & tb.Top & " W=" & tb.Width & " H=" & tb.Height
    Exit Sub

EH:
    Debug.Print "[EnsureDraftBox][ERR]", Err.Number, Err.Description
End Sub


Private Function JsonEsc_Local(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEsc_Local = s
End Function


Public Sub ListOpenWorkbooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        Debug.Print wb.name
    Next wb
End Sub


Public Sub Verify_MonthlyExport_OpensWorkbook()
    Debug.Print "Before: ActiveWB=" & ActiveWorkbook.name & "  Count=" & Workbooks.Count

    ' いま使ってる呼び出しと同じ
    Call ExportMonitoring_ToMonthlyWorkbook( _
        CDate(frmEval.Controls("txtDailyDate").Text), _
        frmEval.Controls("frHeader").Controls("txtHdrName").Text, _
        frmEval.Controls("txtMonthlyMonitoringDraft").Text)

    Debug.Print "After:  ActiveWB=" & ActiveWorkbook.name & "  Count=" & Workbooks.Count

    ' 開いているブック名を列挙
    ListOpenWorkbooks
End Sub











