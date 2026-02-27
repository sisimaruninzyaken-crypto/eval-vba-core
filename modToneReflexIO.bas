Attribute VB_Name = "modToneReflexIO"
Option Explicit

' ローカル定義（他と競合しないよう Private）
Private Const SEP_REC As String = "|"
Private Const SEP_KV  As String = ":"
Private Const SEP_RL  As String = ","

'========================================================
' 筋緊張・反射（痙縮含む） 保存：TONE_IO / TONE_NOTE
'========================================================
Public Sub SaveToneReflexToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    If owner Is Nothing Then If VBA.UserForms.Count > 0 Then Set owner = VBA.UserForms(0)

    Dim ctl As Object, mp As Object, pg As Object, target As Object
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection

    ' 1) 「筋緊張」 or 「反射」を含むページを特定
    On Error Resume Next
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "筋緊張") > 0 Or InStr(pg.caption, "反射") > 0 Then
                    Set target = pg: Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) 対象ページ内の ComboBox を収集（Frame内も掘る）
    q.Add target
    Do While q.Count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.Controls
            Set tmp = ch.Controls
            If Err.Number = 0 Then q.Add ch     ' 子あり
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop
    If combos.Count = 0 Then Exit Sub

    ' 3) 重複除去 → Top/Left で安定ソート
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    Dim uniq As New Collection, i As Long, j As Long
    For i = 1 To combos.Count
        If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
    Next i
    Dim arr() As Object: ReDim arr(1 To uniq.Count)
    For i = 1 To uniq.Count: Set arr(i) = uniq(i): Next i

    Const tol As Single = 6
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If (arr(j).Top < arr(i).Top - tol) _
            Or (Abs(arr(j).Top - arr(i).Top) <= tol And arr(j).Left < arr(i).Left) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    ' 4) R→L ペアでシリアライズ（8項目）
    Dim keys As Variant
    keys = Array( _
        "MAS_上肢屈筋群", "MAS_上肢伸筋群", "MAS_下肢屈筋群", "MAS_下肢伸筋群", _
        "反射_上腕二頭筋", "反射_上腕三頭筋", "反射_膝蓋腱", "反射_アキレス腱")

    Dim pos As Long: pos = 1
    Dim k As Long, vR As String, vL As String, s As String

    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For

        ' 右（R）この1行を差し替え
vR = CStr(arr(pos).value): If Len(vR) = 0 Then vR = CStr(arr(pos).Text)

' 左（L）この1行を差し替え
vL = CStr(arr(pos + 1).value): If Len(vL) = 0 Then vL = CStr(arr(pos + 1).Text)


        If Len(s) > 0 Then s = s & SEP_REC
        
        Debug.Print "[TONE][GRAB]"; keys(k); " | R="; vR; " L="; vL; " | Rnm="; arr(pos).name; " Lnm="; arr(pos + 1).name; " | Ridx="; arr(pos).ListIndex; " Lidx="; arr(pos + 1).ListIndex

        s = s & keys(k) & SEP_KV & "R=" & vR & SEP_RL & "L=" & vL

        pos = pos + 2
    Next k

    ' 5) 書き出し（TONE_IO）
    Dim c As Long: c = EnsureHeaderCol(ws, "TONE_IO")
    ws.Cells(r, c).value = s
    Debug.Print "[TONE][SAVE] row=" & r & " col=" & c & " len=" & Len(s)

    ' 6) 備考（最も大きい or MultiLine TextBox）→ TONE_NOTE
    Dim noteCtl As Object, box As Object, subCtl As Object, bestH As Single: bestH = 0
    Dim note As String, cNote As Long

    On Error Resume Next
    For Each box In target.Controls
        If TypeName(box) = "TextBox" Then
            If box.multiline Or box.Height > bestH Then Set noteCtl = box: bestH = box.Height
        ElseIf TypeName(box) = "Frame" Then
            For Each subCtl In box.Controls
                If TypeName(subCtl) = "TextBox" Then
                    If subCtl.multiline Or subCtl.Height > bestH Then Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            Next subCtl
        End If
    Next box
    On Error GoTo 0

    If Not noteCtl Is Nothing Then note = CStr(noteCtl.Text) Else note = ""
    cNote = EnsureHeaderCol(ws, "TONE_NOTE")
    ws.Cells(r, cNote).value = note
    Debug.Print "[TONE][SAVE][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note)
End Sub


'========================================================
' 筋緊張・反射（痙縮含む） 読み込み：TONE_IO / TONE_NOTE
'========================================================
Public Sub LoadToneReflexFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    If owner Is Nothing Then If VBA.UserForms.Count > 0 Then Set owner = VBA.UserForms(0)

    Dim ctl As Object, mp As Object, pg As Object, target As Object
    ' 1) 「筋緊張」or「反射」を含むページを特定
    On Error Resume Next
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "筋緊張") > 0 Or InStr(pg.caption, "反射") > 0 Then
                    Set target = pg: Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) TONE_IO を辞書にパース（key → Array(R, L)）
    Dim c As Long, s As String, recs As Variant, rec As Variant
    Dim kv As Variant, rl As Variant
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1

    c = EnsureHeaderCol(ws, "TONE_IO")
        s = ReadStr_Compat("IO_Tone", r, ws)
    Debug.Print "[TONE][LOAD] row=" & r & " col=" & c & " len=" & Len(s)

    If Len(s) > 0 Then
        recs = Split(s, SEP_REC)
        For Each rec In recs
            If Len(rec) = 0 Then GoTo cont
            kv = Split(rec, SEP_KV)              ' key : R=..,L=..
            If UBound(kv) < 1 Then GoTo cont
            rl = Split(kv(1), SEP_RL)            ' R=.. , L=..
            If UBound(rl) < 1 Then GoTo cont
            d(CStr(kv(0))) = Array( _
                CStr(Split(rl(0), "=")(1)), _
                CStr(Split(rl(1), "=")(1)))
cont:
        Next rec
    End If

    ' 3) 対象ページ内の ComboBox を収集（Frameも掘る）→ Top/Left ソート
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection
    q.Add target
    Do While q.Count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.Controls
            Set tmp = ch.Controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop
    If combos.Count = 0 Then Exit Sub

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    Dim uniq As New Collection, i As Long, j As Long
    For i = 1 To combos.Count
        If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
    Next i
    Dim arr() As Object: ReDim arr(1 To uniq.Count)
    For i = 1 To uniq.Count: Set arr(i) = uniq(i): Next i

    Const tol As Single = 6
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If (arr(j).Top < arr(i).Top - tol) _
            Or (Abs(arr(j).Top - arr(i).Top) <= tol And arr(j).Left < arr(i).Left) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    ' 4) 保存と同じキー順で反映（R→L）
    Dim keys As Variant
    keys = Array( _
        "MAS_上肢屈筋群", "MAS_上肢伸筋群", "MAS_下肢屈筋群", "MAS_下肢伸筋群", _
        "反射_上腕二頭筋", "反射_上腕三頭筋", "反射_膝蓋腱", "反射_アキレス腱")
    Dim pos As Long: pos = 1
    Dim k As Long, pair As Variant
    Dim tR As String, tL As String
    Dim ii As Long, found As Boolean

    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For
        If d.exists(keys(k)) Then
            pair = d(keys(k))              ' pair(0)=R, pair(1)=L

            ' --- R側（Value優先、合わなければリスト走査→必要ならAddItem） ---
            tR = CStr(pair(0))
            On Error Resume Next
            arr(pos).value = tR
            If Trim$(CStr(arr(pos).value)) <> Trim$(tR) Then
                found = False
                For ii = 0 To arr(pos).ListCount - 1
                    If StrComp(Trim$(arr(pos).List(ii, 0)), Trim$(tR), vbTextCompare) = 0 Then
                        arr(pos).ListIndex = ii: found = True: Exit For
                    End If
                Next ii
                If Not found And Len(tR) > 0 Then
                    arr(pos).AddItem tR
                    arr(pos).value = tR
                End If
            End If
            
 
 

          ' --- L側（リストから直接選択。なければ追加して選択） ---
tL = CStr(pair(1))
 j = -1
On Error Resume Next
For ii = 0 To arr(pos + 1).ListCount - 1
    If StrComp(Trim$(arr(pos + 1).List(ii, 0)), Trim$(tL), vbTextCompare) = 0 Then
        j = ii: Exit For
    End If
Next ii
If j >= 0 Then
    arr(pos + 1).ListIndex = j
    ' Valueが空のケース対策：選択文字をValueにも入れる
    If Len(CStr(arr(pos + 1).value)) = 0 Then arr(pos + 1).value = CStr(arr(pos + 1).List(j, 0))
ElseIf Len(tL) > 0 Then
    arr(pos + 1).AddItem tL
    arr(pos + 1).value = tL
End If
On Error GoTo 0


           

        Else
            Debug.Print "[TONE][MISS]"; keys(k); "（データなし）"
        End If
        pos = pos + 2
    Next k

    ' 5) 備考読み込み（TONE_NOTE）
    Dim cNote As Long, note As String
    Dim noteCtl As Object, box As Object, subCtl As Object, bestH As Single: bestH = 0

    cNote = EnsureHeaderCol(ws, "TONE_NOTE")
    note = CStr(ws.Cells(r, cNote).value)

    On Error Resume Next
    For Each box In target.Controls
        If TypeName(box) = "TextBox" Then
            If box.multiline Or box.Height > bestH Then Set noteCtl = box: bestH = box.Height
        ElseIf TypeName(box) = "Frame" Then
            For Each subCtl In box.Controls
                If TypeName(subCtl) = "TextBox" Then
                    If subCtl.multiline Or subCtl.Height > bestH Then Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            Next subCtl
        End If
    Next box
    On Error GoTo 0

    If Not noteCtl Is Nothing Then noteCtl.Text = note
    Debug.Print "[TONE][LOAD][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & _
                IIf(noteCtl Is Nothing, " (target textbox not found)", " -> " & noteCtl.name & " H=" & bestH)

If TypeOf owner Is MSForms.UserForm Then owner.Repaint



End Sub







Public Sub SetPainHeights(ByVal h As Single)
    With frmEval.Controls("Frame12")
        .Controls("fraPainFactors").Height = h
        .Controls("fraPainSite").Height = h
        Debug.Print "[heights]", .Controls("fraPainFactors").Height, .Controls("fraPainSite").Height
    End With
End Sub



Public Sub PlacePainFactorsBesideSite()
    Dim z As MSForms.Frame, pf As MSForms.Frame, ps As MSForms.Frame, lb As MSForms.label
    Set z = frmEval.Controls("Frame12")
    Set pf = z.Controls("fraPainFactors")    ' 誘因・軽減因子（枠）
    Set ps = z.Controls("fraPainSite")       ' 疼痛部位（枠）
    Set lb = z.Controls("lblPainFactors")    ' ラベル

    Dim M As Single, avail As Single
    M = 12                                   ' 余白
    ' 位置だけ調整：疼痛部位と同じTop、右隣へ
    pf.Top = ps.Top
    pf.Left = ps.Left + ps.Width + M

    ' はみ出し防止（必要なら幅だけ詰める）
    avail = z.Width - M - pf.Left
    If avail < pf.Width Then pf.Width = avail

    ' ラベルは枠の直上に揃える
    lb.Left = pf.Left
    lb.Top = pf.Top - lb.Height - 4

    Debug.Print "[PlacePF]", "Top=", pf.Top, "Left=", pf.Left, "W=", pf.Width
End Sub

