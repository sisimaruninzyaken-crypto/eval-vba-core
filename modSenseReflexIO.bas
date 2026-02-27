Attribute VB_Name = "modSenseReflexIO"
'==== modSenseReflexIO ====
Option Explicit

' 形式:
'   1レコードは  baseKey:R=<ListIndex>,L=<Value>
'   区切りは "|" 例) HyomenShokkaku:R=2,L=2|ShinbuIchi:R=1,L=1
Private Const SEP_REC As String = "|"
Private Const SEP_KV  As String = ":"
Private Const SEP_RL  As String = ","




Public Function SerializeRL(container As Object) As String
    Dim s As String
    Dim q As New Collection
    Dim node As Object, ch As Object
    Dim nm As String, base As String, side As String
    Dim dr As Object, dl As Object  ' Dictionary (late bind)

    Set dr = CreateObject("Scripting.Dictionary")
    Set dl = CreateObject("Scripting.Dictionary")
    dr.CompareMode = 1: dl.CompareMode = 1  ' TextCompare

    ' 幅優先でコンテナ配下を総なめ（Frame/Pages含む）
    q.Add container
    Do While q.Count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next

        ' 子コントロールを走査
        For Each ch In node.Controls
            ' 子がさらに Controls/Pages を持つならキューへ
            Dim dummy As Object, pg As Object
            Set dummy = ch.Controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear
            
            ' ComboBox のみ対象
            If TypeName(ch) = "ComboBox" Then
                nm = ch.name: side = "": base = ""
                ' 先頭プレフィクス
                If LCase$(Left$(nm, 5)) = "cbor_" Then side = "R": base = Mid$(nm, 6)
                If LCase$(Left$(nm, 5)) = "cbol_" Then side = "L": base = Mid$(nm, 6)
                ' 末尾サフィックス (_R/_L)
                If side = "" Then
                    If LCase$(Right$(nm, 2)) = "_r" Then side = "R": base = Left$(nm, Len(nm) - 2)
                    If LCase$(Right$(nm, 2)) = "_l" Then side = "L": base = Left$(nm, Len(nm) - 2)
                End If
                ' アンダースコア無し（cboR肩屈曲 など）
                If side = "" Then
                    If LCase$(Left$(nm, 4)) = "cbor" Then side = "R": base = Mid$(nm, 5)
                    If LCase$(Left$(nm, 4)) = "cbol" Then side = "L": base = Mid$(nm, 5)
                End If

                If side = "R" And Len(base) > 0 Then dr(base) = ch
                If side = "L" And Len(base) > 0 Then dl(base) = ch
            End If
        Next ch

        
        On Error GoTo 0
    Loop

    ' R/L が揃ったものだけ出力
    Dim k As Variant
    For Each k In dr.keys
        If dl.exists(k) Then
            If Len(s) > 0 Then s = s & SEP_REC
            s = s & CStr(k) & SEP_KV & "R=" & CStr(dr(k).ListIndex) & SEP_RL & "L=" & CStr(dl(k).value)
        End If
    Next k

    SerializeRL = s
End Function

Public Sub DeserializeRL(container As Object, ByVal payload As String)
    Dim recs As Variant, rec As Variant
    Dim kv As Variant, rl As Variant, base As String
    Dim rIdx As String, lVal As String
    Dim rCtl As MSForms.ComboBox, lCtl As MSForms.ComboBox

    If Len(payload) = 0 Then Exit Sub
    recs = Split(payload, SEP_REC)

    For Each rec In recs
        If Len(rec) = 0 Then GoTo NextRec
        kv = Split(rec, SEP_KV)
        If UBound(kv) < 1 Then GoTo NextRec

        base = CStr(kv(0))
        rl = Split(kv(1), SEP_RL)
        If UBound(rl) < 1 Then GoTo NextRec

        On Error Resume Next
        rIdx = val(Split(rl(0), "=")(1))   ' "R=2"
        lVal = CStr(Split(rl(1), "=")(1))  ' "L=1" など

        ' 再帰探索で深い階層のコンボも検出
        Set rCtl = FindCtlDeep(container, "cboR_" & base)
        Set lCtl = FindCtlDeep(container, "cboL_" & base)
        On Error GoTo 0

        If Not rCtl Is Nothing Then rCtl.ListIndex = rIdx
        If Not lCtl Is Nothing Then lCtl.value = lVal

NextRec:
        Set rCtl = Nothing: Set lCtl = Nothing
    Next rec
End Sub




'--- 列 I/O（ヘッダは可変） ---
Public Sub SaveRLToSheet(ws As Worksheet, ByVal r As Long, ByVal header As String, container As Object)
    Dim c As Long, s As String
    c = EnsureHeaderCol(ws, header)
    s = SerializeRL(container)
    ws.Cells(r, c).value = s
    Debug.Print "[RL][SAVE] row=" & r & " col=" & c & " len=" & Len(s)
End Sub

Public Sub LoadRLFromSheet(ws As Worksheet, ByVal r As Long, ByVal header As String, container As Object)
    Dim c As Long, s As String
    c = EnsureHeaderCol(ws, header)
    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[RL][LOAD] row=" & r & " col=" & c & " len=" & Len(s)
    Call DeserializeRL(container, s)
End Sub
'==== /modSenseReflexIO ====


Public Sub SaveSensoryToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim ctl As Object, mp As Object, pg As Object, target As Object
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection
    Dim i As Long, j As Long
    Dim c As Long, s As String

    ' ① MultiPage 内から Caption に「感覚」を含む Page を特定（例：感覚（表在・深部））
    On Error Resume Next
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "感覚") > 0 Then
                    Set target = pg
                    Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner  ' 念のため

    ' ② target 配下を幅優先で走査して ComboBox をすべて収集（Frame の中も掘る）
    q.Add target
    Do While q.Count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.Controls
            ' 子がさらに Controls を持つならキューに積む
            Set tmp = ch.Controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear

            If TypeName(ch) = "ComboBox" Then
                combos.Add ch
            End If
        Next ch
        On Error GoTo 0
    Loop

    ' ③ 位置で安定ソート（Top → Left）。行ずれ吸収のため Top は±6の許容で比較
    If combos.Count = 0 Then
        s = "" ' 何も無ければ空
        GoTo WRITE_OUT
    End If

   Dim arr() As Object, seen As Object, uniq As New Collection
Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
For i = 1 To combos.Count
    If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
Next i
ReDim arr(1 To uniq.Count)
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

        ' ④ 並び順でキー割当（右→左の順で1項目とみなす）
    '    想定順：表在_触覚, 表在_痛覚, 表在_温度覚, 深部_位置覚, 深部_振動覚
    Dim keys As Variant
    keys = Array("表在_触覚", "表在_痛覚", "表在_温度覚", "深部_位置覚", "深部_振動覚")

    s = ""
    Dim k As Long
    Dim pos As Long
    Dim vR As String, vL As String

    pos = 1
    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For   ' 安全

        ' --- 右側コンボ（R） ---
        If arr(pos).ListIndex >= 0 Then
            vR = CStr(arr(pos).List(arr(pos).ListIndex, 0))
        Else
            vR = CStr(arr(pos).Text)
            If Len(vR) = 0 Then vR = CStr(arr(pos).value)
        End If

        ' --- 左側コンボ（L） ---
        If pos + 1 <= UBound(arr) Then
            If arr(pos + 1).ListIndex >= 0 Then
                vL = CStr(arr(pos + 1).List(arr(pos + 1).ListIndex, 0))
            Else
                vL = CStr(arr(pos + 1).Text)
                If Len(vL) = 0 Then vL = CStr(arr(pos + 1).value)
            End If
        Else
            vL = ""
        End If

        ' --- 文字列組み立て ---
        If Len(s) > 0 Then s = s & SEP_REC
        s = s & keys(k) & SEP_KV & "R=" & vR & SEP_RL & "L=" & vL

        

        pos = pos + 2
    Next k


WRITE_OUT:
    ' ⑤ シート書き出し（IO_Sensory に今回組み立てた s を保存）
    c = EnsureHeader(ws, "IO_Sensory")
    ws.Cells(r, c).value = s
    Debug.Print "[SENSE][SAVE] row=" & r & " col=" & c & " len=" & Len(s)


' --- SENSE 備考を保存（SENSE_NOTE） ---
Dim noteCtl As Object, box As Object, subCtl As Object
Dim bestH As Single: bestH = 0
Dim note As String

On Error Resume Next
' ページ内で MultiLine または最も背の高い TextBox を選ぶ（＝読み込みと同じ基準）
For Each box In target.Controls
    If TypeName(box) = "TextBox" Then
        If box.multiline Or box.Height > bestH Then
            Set noteCtl = box: bestH = box.Height
        End If
    ElseIf TypeName(box) = "Frame" Then
        For Each subCtl In box.Controls
            If TypeName(subCtl) = "TextBox" Then
                If subCtl.multiline Or subCtl.Height > bestH Then
                    Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            End If
        Next subCtl
    End If
Next box
On Error GoTo 0

If Not noteCtl Is Nothing Then note = CStr(noteCtl.Text) Else note = ""

Dim cNote As Long
cNote = EnsureHeaderCol(ws, "SENSE_NOTE")
ws.Cells(r, cNote).value = note
Debug.Print "[SENSE][SAVE][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & _
            " <- " & IIf(noteCtl Is Nothing, "(not found)", TypeName(noteCtl) & ":" & noteCtl.name & " H=" & bestH)

End Sub









' 再帰的に子孫からコントロールを探す
Private Function FindCtlDeep(root As Object, ByVal ctlName As String) As Object
    Dim ch As Object, tmp As Object
    On Error Resume Next
    Set FindCtlDeep = root.Controls(ctlName) ' まず直下を試す
    On Error GoTo 0
    If Not FindCtlDeep Is Nothing Then Exit Function

    ' 子を順に掘る（Controlsを持たない場合はエラーを握りつぶす）
    For Each ch In root.Controls
        On Error Resume Next
        Set tmp = ch.Controls
        If Err.Number = 0 Then
            Set tmp = FindCtlDeep(ch, ctlName)
            If Not tmp Is Nothing Then
                Set FindCtlDeep = tmp
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next ch
End Function


Public Sub TraceSensoryComboNames(owner As Object)
    Dim ctl As Object, subCtl As Object, mp As Object, pg As Object, target As Object

    ' MultiPage内で Caption に「感覚」を含むページだけ特定
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "感覚") > 0 Then
                    Set target = pg
                    Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    If target Is Nothing Then Set target = owner

    ' 1階層＋Frame内のComboBoxだけを列挙（再帰なし、.Pagesにも触れない）
    For Each ctl In target.Controls
        If TypeName(ctl) = "ComboBox" Then
            Debug.Print "[SENSE][CB] "; ctl.name
        ElseIf TypeName(ctl) = "Frame" Then
            For Each subCtl In ctl.Controls
                If TypeName(subCtl) = "ComboBox" Then
                    Debug.Print "[SENSE][CB] "; subCtl.name
                End If
            Next subCtl
        End If
    Next ctl
End Sub

Public Sub LoadSensoryFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    
    If owner Is Nothing Then
    If VBA.UserForms.Count > 0 Then Set owner = VBA.UserForms(0)
End If

    
    
    
    
    Dim ctl As Object, mp As Object, pg As Object, target As Object
    Dim c As Long, s As String, recs As Variant, rec As Variant
    Dim kv As Variant, rl As Variant, k As Long, pos As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1

    ' 1) 「感覚」を含むタブを特定（例：感覚（表在・深部））
    On Error Resume Next
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "感覚") > 0 Then Set target = pg: Exit For
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) シートから SENSE_IO 取得→辞書にパース（R/LともValueで扱う）
    s = ReadStr_Compat("IO_Sensory", r, ws)
    s = ReadStr_Compat("IO_Sensory", r, ws)
    
    If Len(s) > 0 Then
        recs = Split(s, SEP_REC) ' "|" 区切り
        For Each rec In recs
            If Len(rec) = 0 Then GoTo cont
            kv = Split(rec, SEP_KV)  ' "key:R=..,L=.."
            If UBound(kv) < 1 Then GoTo cont
            rl = Split(kv(1), SEP_RL) ' "R=..","L=.."
            If UBound(rl) < 1 Then GoTo cont
            d(CStr(kv(0))) = Array(CStr(Split(rl(0), "=")(1)), CStr(Split(rl(1), "=")(1)))
cont:
        Next rec
    End If

    ' 3) 感覚タブ内の ComboBox を収集（Frame内も掘る：再帰なしの幅優先）
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection
    q.Add target
    Do While q.Count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.Controls
            Set tmp = ch.Controls
            If Err.Number = 0 Then q.Add ch  ' 子を掘る
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop

    ' 4) 重複除去→Top→Leftで安定ソート
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

    ' 5) 保存と同じキー順で反映（R→L の並びで1項目）
    Dim keys As Variant: keys = Array("表在_触覚", "表在_痛覚", "表在_温度覚", "深部_位置覚", "深部_振動覚")
    pos = 1
    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For
        If d.exists(keys(k)) Then
            Dim pair As Variant
            pair = d(keys(k))                  ' pair(0)=Rの値, pair(1)=Lの値
            On Error Resume Next
       ' R側（未登録の値でも確実に表示）
Dim tR As String: tR = CStr(pair(0))
Dim ii As Long, found As Boolean
arr(pos).value = tR
If Trim$(CStr(arr(pos).value)) <> Trim$(tR) Then
    For ii = 0 To arr(pos).ListCount - 1
        If StrComp(Trim$(arr(pos).List(ii, 0)), Trim$(tR), vbTextCompare) = 0 Then
            arr(pos).ListIndex = ii
            found = True
            Exit For
        End If
    Next ii
    If Not found And Len(tR) > 0 Then
        arr(pos).AddItem tR
        arr(pos).value = tR
    End If
End If

' L側：同様にフォールバック
Dim tL As String: tL = CStr(pair(1))
arr(pos + 1).value = tL
If Trim$(CStr(arr(pos + 1).value)) <> Trim$(tL) Then
    For ii = 0 To arr(pos + 1).ListCount - 1
        If StrComp(Trim$(arr(pos + 1).List(ii, 0)), Trim$(tL), vbTextCompare) = 0 Then
            arr(pos + 1).ListIndex = ii
            Exit For
        End If
    Next ii
End If

            On Error GoTo 0
            
        Else
            Debug.Print "[SENSE][MISS]"; keys(k); "（データなし）"
        End If
        pos = pos + 2
    Next k
    
    
    ' --- SENSE 備考を読み込み（SENSE_NOTE） ---
Dim cNote As Long, note As String
Dim box As Object, subCtl As Object, noteCtl As Object
Dim bestH As Single: bestH = 0

cNote = EnsureHeaderCol(ws, "SENSE_NOTE")
note = CStr(ws.Cells(r, cNote).value)

On Error Resume Next
' ページ内で MultiLine または最も背の高い TextBox を選ぶ
For Each box In target.Controls
    If TypeName(box) = "TextBox" Then
        If box.multiline Or box.Height > bestH Then
            Set noteCtl = box: bestH = box.Height
        End If
    ElseIf TypeName(box) = "Frame" Then
        For Each subCtl In box.Controls
            If TypeName(subCtl) = "TextBox" Then
                If subCtl.multiline Or subCtl.Height > bestH Then
                    Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            End If
        Next subCtl
    End If
Next box
On Error GoTo 0

If Not noteCtl Is Nothing Then
    noteCtl.Text = note
    
Else
    Debug.Print "[SENSE][LOAD][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & " (target textbox not found)"
End If

    
    
End Sub







