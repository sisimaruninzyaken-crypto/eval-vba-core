Attribute VB_Name = "modToneReflexIO"
Option Explicit

' 繝ｭ繝ｼ繧ｫ繝ｫ螳夂ｾｩ・井ｻ悶→遶ｶ蜷医＠縺ｪ縺・ｈ縺・Private・・
Private Const SEP_REC As String = "|"
Private Const SEP_KV  As String = ":"
Private Const SEP_RL  As String = ","

'========================================================
' 遲狗ｷ雁ｼｵ繝ｻ蜿榊ｰ・ｼ育吏邵ｮ蜷ｫ繧・・菫晏ｭ假ｼ啜ONE_IO / TONE_NOTE
'========================================================
Public Sub SaveToneReflexToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    If owner Is Nothing Then If VBA.UserForms.count > 0 Then Set owner = VBA.UserForms(0)

    Dim ctl As Object, mp As Object, pg As Object, target As Object
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection

    ' 1) 縲檎ｭ狗ｷ雁ｼｵ縲・or 縲悟渚蟆・阪ｒ蜷ｫ繧繝壹・繧ｸ繧堤音螳・
    On Error Resume Next
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "遲狗ｷ雁ｼｵ") > 0 Or InStr(pg.caption, "蜿榊ｰ・) > 0 Then
                    Set target = pg: Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) 蟇ｾ雎｡繝壹・繧ｸ蜀・・ ComboBox 繧貞庶髮・ｼ・rame蜀・ｂ謗倥ｋ・・
    q.Add target
    Do While q.count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.controls
            Set tmp = ch.controls
            If Err.Number = 0 Then q.Add ch     ' 蟄舌≠繧・
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop
    If combos.count = 0 Then Exit Sub

    ' 3) 驥崎､・勁蜴ｻ 竊・Top/Left 縺ｧ螳牙ｮ壹た繝ｼ繝・
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    Dim uniq As New Collection, i As Long, j As Long
    For i = 1 To combos.count
        If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
    Next i
    Dim arr() As Object: ReDim arr(1 To uniq.count)
    For i = 1 To uniq.count: Set arr(i) = uniq(i): Next i

    Const tol As Single = 6
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If (arr(j).Top < arr(i).Top - tol) _
            Or (Abs(arr(j).Top - arr(i).Top) <= tol And arr(j).Left < arr(i).Left) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    ' 4) R竊鱈 繝壹い縺ｧ繧ｷ繝ｪ繧｢繝ｩ繧､繧ｺ・・鬆・岼・・
    Dim keys As Variant
    keys = Array( _
        "MAS_荳願い螻育ｭ狗ｾ､", "MAS_荳願い莨ｸ遲狗ｾ､", "MAS_荳玖い螻育ｭ狗ｾ､", "MAS_荳玖い莨ｸ遲狗ｾ､", _
        "蜿榊ｰЮ荳願・莠碁ｭ遲・, "蜿榊ｰЮ荳願・荳蛾ｭ遲・, "蜿榊ｰЮ閹晁搭閻ｱ", "蜿榊ｰЮ繧｢繧ｭ繝ｬ繧ｹ閻ｱ")

    Dim pos As Long: pos = 1
    Dim k As Long, vR As String, vL As String, s As String

    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For

        ' 蜿ｳ・・・峨％縺ｮ1陦後ｒ蟾ｮ縺玲崛縺・
vR = CStr(arr(pos).value): If Len(vR) = 0 Then vR = CStr(arr(pos).text)

' 蟾ｦ・・・峨％縺ｮ1陦後ｒ蟾ｮ縺玲崛縺・
vL = CStr(arr(pos + 1).value): If Len(vL) = 0 Then vL = CStr(arr(pos + 1).text)


        If Len(s) > 0 Then s = s & SEP_REC
        
        Debug.Print "[TONE][GRAB]"; keys(k); " | R="; vR; " L="; vL; " | Rnm="; arr(pos).name; " Lnm="; arr(pos + 1).name; " | Ridx="; arr(pos).ListIndex; " Lidx="; arr(pos + 1).ListIndex

        s = s & keys(k) & SEP_KV & "R=" & vR & SEP_RL & "L=" & vL

        pos = pos + 2
    Next k

    ' 5) 譖ｸ縺榊・縺暦ｼ・ONE_IO・・
    Dim c As Long: c = EnsureHeaderCol(ws, "TONE_IO")
    ws.Cells(r, c).value = s
    Debug.Print "[TONE][SAVE] row=" & r & " col=" & c & " len=" & Len(s)

    ' 6) 蛯呵・ｼ域怙繧ょ､ｧ縺阪＞ or MultiLine TextBox・俄・ TONE_NOTE
    Dim noteCtl As Object, box As Object, subCtl As Object, bestH As Single: bestH = 0
    Dim note As String, cNote As Long

    On Error Resume Next
    For Each box In target.controls
        If TypeName(box) = "TextBox" Then
            If box.multiline Or box.Height > bestH Then Set noteCtl = box: bestH = box.Height
        ElseIf TypeName(box) = "Frame" Then
            For Each subCtl In box.controls
                If TypeName(subCtl) = "TextBox" Then
                    If subCtl.multiline Or subCtl.Height > bestH Then Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            Next subCtl
        End If
    Next box
    On Error GoTo 0

    If Not noteCtl Is Nothing Then note = CStr(noteCtl.text) Else note = ""
    cNote = EnsureHeaderCol(ws, "TONE_NOTE")
    ws.Cells(r, cNote).value = note
    Debug.Print "[TONE][SAVE][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note)
End Sub


'========================================================
' 遲狗ｷ雁ｼｵ繝ｻ蜿榊ｰ・ｼ育吏邵ｮ蜷ｫ繧・・隱ｭ縺ｿ霎ｼ縺ｿ・啜ONE_IO / TONE_NOTE
'========================================================
Public Sub LoadToneReflexFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    If owner Is Nothing Then If VBA.UserForms.count > 0 Then Set owner = VBA.UserForms(0)

    Dim ctl As Object, mp As Object, pg As Object, target As Object
    ' 1) 縲檎ｭ狗ｷ雁ｼｵ縲腔r縲悟渚蟆・阪ｒ蜷ｫ繧繝壹・繧ｸ繧堤音螳・
    On Error Resume Next
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "遲狗ｷ雁ｼｵ") > 0 Or InStr(pg.caption, "蜿榊ｰ・) > 0 Then
                    Set target = pg: Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) TONE_IO 繧定ｾ樊嶌縺ｫ繝代・繧ｹ・・ey 竊・Array(R, L)・・
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

    ' 3) 蟇ｾ雎｡繝壹・繧ｸ蜀・・ ComboBox 繧貞庶髮・ｼ・rame繧よ侍繧具ｼ俄・ Top/Left 繧ｽ繝ｼ繝・
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection
    q.Add target
    Do While q.count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.controls
            Set tmp = ch.controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop
    If combos.count = 0 Then Exit Sub

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    Dim uniq As New Collection, i As Long, j As Long
    For i = 1 To combos.count
        If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
    Next i
    Dim arr() As Object: ReDim arr(1 To uniq.count)
    For i = 1 To uniq.count: Set arr(i) = uniq(i): Next i

    Const tol As Single = 6
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If (arr(j).Top < arr(i).Top - tol) _
            Or (Abs(arr(j).Top - arr(i).Top) <= tol And arr(j).Left < arr(i).Left) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    ' 4) 菫晏ｭ倥→蜷後§繧ｭ繝ｼ鬆・〒蜿肴丐・・竊鱈・・
    Dim keys As Variant
    keys = Array( _
        "MAS_荳願い螻育ｭ狗ｾ､", "MAS_荳願い莨ｸ遲狗ｾ､", "MAS_荳玖い螻育ｭ狗ｾ､", "MAS_荳玖い莨ｸ遲狗ｾ､", _
        "蜿榊ｰЮ荳願・莠碁ｭ遲・, "蜿榊ｰЮ荳願・荳蛾ｭ遲・, "蜿榊ｰЮ閹晁搭閻ｱ", "蜿榊ｰЮ繧｢繧ｭ繝ｬ繧ｹ閻ｱ")
    Dim pos As Long: pos = 1
    Dim k As Long, pair As Variant
    Dim tR As String, tL As String
    Dim ii As Long, found As Boolean

    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For
        If d.exists(keys(k)) Then
            pair = d(keys(k))              ' pair(0)=R, pair(1)=L

            ' --- R蛛ｴ・・alue蜆ｪ蜈医∝粋繧上↑縺代ｌ縺ｰ繝ｪ繧ｹ繝郁ｵｰ譟ｻ竊貞ｿ・ｦ√↑繧陰ddItem・・---
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
            
 
 

          ' --- L蛛ｴ・医Μ繧ｹ繝医°繧臥峩謗･驕ｸ謚槭ゅ↑縺代ｌ縺ｰ霑ｽ蜉縺励※驕ｸ謚橸ｼ・---
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
    ' Value縺檎ｩｺ縺ｮ繧ｱ繝ｼ繧ｹ蟇ｾ遲厄ｼ夐∈謚樊枚蟄励ｒValue縺ｫ繧ょ・繧後ｋ
    If Len(CStr(arr(pos + 1).value)) = 0 Then arr(pos + 1).value = CStr(arr(pos + 1).List(j, 0))
ElseIf Len(tL) > 0 Then
    arr(pos + 1).AddItem tL
    arr(pos + 1).value = tL
End If
On Error GoTo 0


           

        Else
            Debug.Print "[TONE][MISS]"; keys(k); "・医ョ繝ｼ繧ｿ縺ｪ縺暦ｼ・
        End If
        pos = pos + 2
    Next k

    ' 5) 蛯呵・ｪｭ縺ｿ霎ｼ縺ｿ・・ONE_NOTE・・
    Dim cNote As Long, note As String
    Dim noteCtl As Object, box As Object, subCtl As Object, bestH As Single: bestH = 0

    cNote = EnsureHeaderCol(ws, "TONE_NOTE")
    note = CStr(ws.Cells(r, cNote).value)

    On Error Resume Next
    For Each box In target.controls
        If TypeName(box) = "TextBox" Then
            If box.multiline Or box.Height > bestH Then Set noteCtl = box: bestH = box.Height
        ElseIf TypeName(box) = "Frame" Then
            For Each subCtl In box.controls
                If TypeName(subCtl) = "TextBox" Then
                    If subCtl.multiline Or subCtl.Height > bestH Then Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            Next subCtl
        End If
    Next box
    On Error GoTo 0

    If Not noteCtl Is Nothing Then noteCtl.text = note
    Debug.Print "[TONE][LOAD][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & _
                IIf(noteCtl Is Nothing, " (target textbox not found)", " -> " & noteCtl.name & " H=" & bestH)

If TypeOf owner Is MSForms.UserForm Then owner.Repaint



End Sub







Public Sub SetPainHeights(ByVal h As Single)
    With frmEval.controls("Frame12")
        .controls("fraPainFactors").Height = h
        .controls("fraPainSite").Height = h
        Debug.Print "[heights]", .controls("fraPainFactors").Height, .controls("fraPainSite").Height
    End With
End Sub



Public Sub PlacePainFactorsBesideSite()
    Dim z As MSForms.Frame, pf As MSForms.Frame, ps As MSForms.Frame, lb As MSForms.label
    Set z = frmEval.controls("Frame12")
    Set pf = z.controls("fraPainFactors")    ' 隱伜屏繝ｻ霆ｽ貂帛屏蟄撰ｼ域棧・・
    Set ps = z.controls("fraPainSite")       ' 逍ｼ逞幃Κ菴搾ｼ域棧・・
    Set lb = z.controls("lblPainFactors")    ' 繝ｩ繝吶Ν

    Dim m As Single, avail As Single
    m = 12                                   ' 菴咏區
    ' 菴咲ｽｮ縺縺題ｪｿ謨ｴ・夂名逞幃Κ菴阪→蜷後§Top縲∝承髫｣縺ｸ
    pf.Top = ps.Top
    pf.Left = ps.Left + ps.Width + m

    ' 縺ｯ縺ｿ蜃ｺ縺鈴亟豁｢・亥ｿ・ｦ√↑繧牙ｹ・□縺題ｩｰ繧√ｋ・・
    avail = z.Width - m - pf.Left
    If avail < pf.Width Then pf.Width = avail

    ' 繝ｩ繝吶Ν縺ｯ譫縺ｮ逶ｴ荳翫↓謠・∴繧・
    lb.Left = pf.Left
    lb.Top = pf.Top - lb.Height - 4

    Debug.Print "[PlacePF]", "Top=", pf.Top, "Left=", pf.Left, "W=", pf.Width
End Sub

