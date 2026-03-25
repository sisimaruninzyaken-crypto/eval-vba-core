Attribute VB_Name = "modSenseReflexIO"
'==== modSenseReflexIO ====
Option Explicit

' 蠖｢蠑・
'   1繝ｬ繧ｳ繝ｼ繝峨・  baseKey:R=<ListIndex>,L=<Value>
'   蛹ｺ蛻・ｊ縺ｯ "|" 萓・ HyomenShokkaku:R=2,L=2|ShinbuIchi:R=1,L=1
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

    ' 蟷・━蜈医〒繧ｳ繝ｳ繝・リ驟堺ｸ九ｒ邱上↑繧・ｼ・rame/Pages蜷ｫ繧・・
    q.Add container
    Do While q.count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next

        ' 蟄舌さ繝ｳ繝医Ο繝ｼ繝ｫ繧定ｵｰ譟ｻ
        For Each ch In node.controls
            ' 蟄舌′縺輔ｉ縺ｫ Controls/Pages 繧呈戟縺､縺ｪ繧峨く繝･繝ｼ縺ｸ
            Dim dummy As Object, pg As Object
            Set dummy = ch.controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear
            
            ' ComboBox 縺ｮ縺ｿ蟇ｾ雎｡
            If TypeName(ch) = "ComboBox" Then
                nm = ch.name: side = "": base = ""
                ' 蜈磯ｭ繝励Ξ繝輔ぅ繧ｯ繧ｹ
                If LCase$(Left$(nm, 5)) = "cbor_" Then side = "R": base = Mid$(nm, 6)
                If LCase$(Left$(nm, 5)) = "cbol_" Then side = "L": base = Mid$(nm, 6)
                ' 譛ｫ蟆ｾ繧ｵ繝輔ぅ繝・け繧ｹ (_R/_L)
                If side = "" Then
                    If LCase$(Right$(nm, 2)) = "_r" Then side = "R": base = Left$(nm, Len(nm) - 2)
                    If LCase$(Right$(nm, 2)) = "_l" Then side = "L": base = Left$(nm, Len(nm) - 2)
                End If
                ' 繧｢繝ｳ繝繝ｼ繧ｹ繧ｳ繧｢辟｡縺暦ｼ・boR閧ｩ螻域峇 縺ｪ縺ｩ・・
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

    ' R/L 縺梧純縺｣縺溘ｂ縺ｮ縺縺大・蜉・
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
        lVal = CStr(Split(rl(1), "=")(1))  ' "L=1" 縺ｪ縺ｩ

        ' 蜀榊ｸｰ謗｢邏｢縺ｧ豺ｱ縺・嚴螻､縺ｮ繧ｳ繝ｳ繝懊ｂ讀懷・
        Set rCtl = FindCtlDeep(container, "cboR_" & base)
        Set lCtl = FindCtlDeep(container, "cboL_" & base)
        On Error GoTo 0

        If Not rCtl Is Nothing Then rCtl.ListIndex = rIdx
        If Not lCtl Is Nothing Then lCtl.value = lVal

NextRec:
        Set rCtl = Nothing: Set lCtl = Nothing
    Next rec
End Sub




'--- 蛻・I/O・医・繝・ム縺ｯ蜿ｯ螟会ｼ・---
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

    ' 竭 MultiPage 蜀・°繧・Caption 縺ｫ縲梧─隕壹阪ｒ蜷ｫ繧 Page 繧堤音螳夲ｼ井ｾ具ｼ壽─隕夲ｼ郁｡ｨ蝨ｨ繝ｻ豺ｱ驛ｨ・会ｼ・
    On Error Resume Next
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "諢溯ｦ・) > 0 Then
                    Set target = pg
                    Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner  ' 蠢ｵ縺ｮ縺溘ａ

    ' 竭｡ target 驟堺ｸ九ｒ蟷・━蜈医〒襍ｰ譟ｻ縺励※ ComboBox 繧偵☆縺ｹ縺ｦ蜿朱寔・・rame 縺ｮ荳ｭ繧よ侍繧具ｼ・
    q.Add target
    Do While q.count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.controls
            ' 蟄舌′縺輔ｉ縺ｫ Controls 繧呈戟縺､縺ｪ繧峨く繝･繝ｼ縺ｫ遨阪・
            Set tmp = ch.controls
            If Err.Number = 0 Then q.Add ch
            Err.Clear

            If TypeName(ch) = "ComboBox" Then
                combos.Add ch
            End If
        Next ch
        On Error GoTo 0
    Loop

    ' 竭｢ 菴咲ｽｮ縺ｧ螳牙ｮ壹た繝ｼ繝茨ｼ・op 竊・Left・峨り｡後★繧悟精蜿弱・縺溘ａ Top 縺ｯﾂｱ6縺ｮ險ｱ螳ｹ縺ｧ豈碑ｼ・
    If combos.count = 0 Then
        s = "" ' 菴輔ｂ辟｡縺代ｌ縺ｰ遨ｺ
        GoTo WRITE_OUT
    End If

   Dim arr() As Object, seen As Object, uniq As New Collection
Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
For i = 1 To combos.count
    If Not seen.exists(combos(i).name) Then seen(combos(i).name) = True: uniq.Add combos(i)
Next i
ReDim arr(1 To uniq.count)
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

        ' 竭｣ 荳ｦ縺ｳ鬆・〒繧ｭ繝ｼ蜑ｲ蠖難ｼ亥承竊貞ｷｦ縺ｮ鬆・〒1鬆・岼縺ｨ縺ｿ縺ｪ縺呻ｼ・
    '    諠ｳ螳夐・ｼ夊｡ｨ蝨ｨ_隗ｦ隕・ 陦ｨ蝨ｨ_逞幄ｦ・ 陦ｨ蝨ｨ_貂ｩ蠎ｦ隕・ 豺ｱ驛ｨ_菴咲ｽｮ隕・ 豺ｱ驛ｨ_謖ｯ蜍戊ｦ・
    Dim keys As Variant
    keys = Array("陦ｨ蝨ｨ_隗ｦ隕・, "陦ｨ蝨ｨ_逞幄ｦ・, "陦ｨ蝨ｨ_貂ｩ蠎ｦ隕・, "豺ｱ驛ｨ_菴咲ｽｮ隕・, "豺ｱ驛ｨ_謖ｯ蜍戊ｦ・)

    s = ""
    Dim k As Long
    Dim pos As Long
    Dim vR As String, vL As String

    pos = 1
    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For   ' 螳牙・

        ' --- 蜿ｳ蛛ｴ繧ｳ繝ｳ繝懶ｼ・・・---
        If arr(pos).ListIndex >= 0 Then
            vR = CStr(arr(pos).List(arr(pos).ListIndex, 0))
        Else
            vR = CStr(arr(pos).text)
            If Len(vR) = 0 Then vR = CStr(arr(pos).value)
        End If

        ' --- 蟾ｦ蛛ｴ繧ｳ繝ｳ繝懶ｼ・・・---
        If pos + 1 <= UBound(arr) Then
            If arr(pos + 1).ListIndex >= 0 Then
                vL = CStr(arr(pos + 1).List(arr(pos + 1).ListIndex, 0))
            Else
                vL = CStr(arr(pos + 1).text)
                If Len(vL) = 0 Then vL = CStr(arr(pos + 1).value)
            End If
        Else
            vL = ""
        End If

        ' --- 譁・ｭ怜・邨・∩遶九※ ---
        If Len(s) > 0 Then s = s & SEP_REC
        s = s & keys(k) & SEP_KV & "R=" & vR & SEP_RL & "L=" & vL

        

        pos = pos + 2
    Next k


WRITE_OUT:
    ' 竭､ 繧ｷ繝ｼ繝域嶌縺榊・縺暦ｼ・O_Sensory 縺ｫ莉雁屓邨・∩遶九※縺・s 繧剃ｿ晏ｭ假ｼ・
    c = EnsureHeader(ws, "IO_Sensory")
    ws.Cells(r, c).value = s
    Debug.Print "[SENSE][SAVE] row=" & r & " col=" & c & " len=" & Len(s)


' --- SENSE 蛯呵・ｒ菫晏ｭ假ｼ・ENSE_NOTE・・---
Dim noteCtl As Object, box As Object, subCtl As Object
Dim bestH As Single: bestH = 0
Dim note As String

On Error Resume Next
' 繝壹・繧ｸ蜀・〒 MultiLine 縺ｾ縺溘・譛繧りレ縺ｮ鬮倥＞ TextBox 繧帝∈縺ｶ・茨ｼ晁ｪｭ縺ｿ霎ｼ縺ｿ縺ｨ蜷後§蝓ｺ貅厄ｼ・
For Each box In target.controls
    If TypeName(box) = "TextBox" Then
        If box.multiline Or box.Height > bestH Then
            Set noteCtl = box: bestH = box.Height
        End If
    ElseIf TypeName(box) = "Frame" Then
        For Each subCtl In box.controls
            If TypeName(subCtl) = "TextBox" Then
                If subCtl.multiline Or subCtl.Height > bestH Then
                    Set noteCtl = subCtl: bestH = subCtl.Height
                End If
            End If
        Next subCtl
    End If
Next box
On Error GoTo 0

If Not noteCtl Is Nothing Then note = CStr(noteCtl.text) Else note = ""

Dim cNote As Long
cNote = EnsureHeaderCol(ws, "SENSE_NOTE")
ws.Cells(r, cNote).value = note
Debug.Print "[SENSE][SAVE][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & _
            " <- " & IIf(noteCtl Is Nothing, "(not found)", TypeName(noteCtl) & ":" & noteCtl.name & " H=" & bestH)

End Sub









' 蜀榊ｸｰ逧・↓蟄仙ｭｫ縺九ｉ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ繧呈爾縺・
Private Function FindCtlDeep(root As Object, ByVal ctlName As String) As Object
    Dim ch As Object, tmp As Object
    On Error Resume Next
    Set FindCtlDeep = root.controls(ctlName) ' 縺ｾ縺夂峩荳九ｒ隧ｦ縺・
    On Error GoTo 0
    If Not FindCtlDeep Is Nothing Then Exit Function

    ' 蟄舌ｒ鬆・↓謗倥ｋ・・ontrols繧呈戟縺溘↑縺・ｴ蜷医・繧ｨ繝ｩ繝ｼ繧呈升繧翫▽縺ｶ縺呻ｼ・
    For Each ch In root.controls
        On Error Resume Next
        Set tmp = ch.controls
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

    ' MultiPage蜀・〒 Caption 縺ｫ縲梧─隕壹阪ｒ蜷ｫ繧繝壹・繧ｸ縺縺醍音螳・
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "諢溯ｦ・) > 0 Then
                    Set target = pg
                    Exit For
                End If
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    If target Is Nothing Then Set target = owner

    ' 1髫主ｱ､・祈rame蜀・・ComboBox縺縺代ｒ蛻玲嫌・亥・蟶ｰ縺ｪ縺励・Pages縺ｫ繧りｧｦ繧後↑縺・ｼ・
    For Each ctl In target.controls
        If TypeName(ctl) = "ComboBox" Then
            Debug.Print "[SENSE][CB] "; ctl.name
        ElseIf TypeName(ctl) = "Frame" Then
            For Each subCtl In ctl.controls
                If TypeName(subCtl) = "ComboBox" Then
                    Debug.Print "[SENSE][CB] "; subCtl.name
                End If
            Next subCtl
        End If
    Next ctl
End Sub

Public Sub LoadSensoryFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    
    If owner Is Nothing Then
    If VBA.UserForms.count > 0 Then Set owner = VBA.UserForms(0)
End If

    
    
    
    
    Dim ctl As Object, mp As Object, pg As Object, target As Object
    Dim c As Long, s As String, recs As Variant, rec As Variant
    Dim kv As Variant, rl As Variant, k As Long, pos As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1

    ' 1) 縲梧─隕壹阪ｒ蜷ｫ繧繧ｿ繝悶ｒ迚ｹ螳夲ｼ井ｾ具ｼ壽─隕夲ｼ郁｡ｨ蝨ｨ繝ｻ豺ｱ驛ｨ・会ｼ・
    On Error Resume Next
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For Each pg In mp.Pages
                If InStr(pg.caption, "諢溯ｦ・) > 0 Then Set target = pg: Exit For
            Next pg
            If Not target Is Nothing Then Exit For
        End If
    Next ctl
    On Error GoTo 0
    If target Is Nothing Then Set target = owner

    ' 2) 繧ｷ繝ｼ繝医°繧・SENSE_IO 蜿門ｾ冷・霎樊嶌縺ｫ繝代・繧ｹ・・/L縺ｨ繧７alue縺ｧ謇ｱ縺・ｼ・
    s = ReadStr_Compat("IO_Sensory", r, ws)
    s = ReadStr_Compat("IO_Sensory", r, ws)
    
    If Len(s) > 0 Then
        recs = Split(s, SEP_REC) ' "|" 蛹ｺ蛻・ｊ
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

    ' 3) 諢溯ｦ壹ち繝門・縺ｮ ComboBox 繧貞庶髮・ｼ・rame蜀・ｂ謗倥ｋ・壼・蟶ｰ縺ｪ縺励・蟷・━蜈茨ｼ・
    Dim q As New Collection, node As Object, ch As Object, tmp As Object
    Dim combos As New Collection
    q.Add target
    Do While q.count > 0
        Set node = q(1): q.Remove 1
        On Error Resume Next
        For Each ch In node.controls
            Set tmp = ch.controls
            If Err.Number = 0 Then q.Add ch  ' 蟄舌ｒ謗倥ｋ
            Err.Clear
            If TypeName(ch) = "ComboBox" Then combos.Add ch
        Next ch
        On Error GoTo 0
    Loop

    ' 4) 驥崎､・勁蜴ｻ竊探op竊鱈eft縺ｧ螳牙ｮ壹た繝ｼ繝・
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

    ' 5) 菫晏ｭ倥→蜷後§繧ｭ繝ｼ鬆・〒蜿肴丐・・竊鱈 縺ｮ荳ｦ縺ｳ縺ｧ1鬆・岼・・
    Dim keys As Variant: keys = Array("陦ｨ蝨ｨ_隗ｦ隕・, "陦ｨ蝨ｨ_逞幄ｦ・, "陦ｨ蝨ｨ_貂ｩ蠎ｦ隕・, "豺ｱ驛ｨ_菴咲ｽｮ隕・, "豺ｱ驛ｨ_謖ｯ蜍戊ｦ・)
    pos = 1
    For k = LBound(keys) To UBound(keys)
        If pos + 1 > UBound(arr) Then Exit For
        If d.exists(keys(k)) Then
            Dim pair As Variant
            pair = d(keys(k))                  ' pair(0)=R縺ｮ蛟､, pair(1)=L縺ｮ蛟､
            On Error Resume Next
       ' R蛛ｴ・域悴逋ｻ骭ｲ縺ｮ蛟､縺ｧ繧ら｢ｺ螳溘↓陦ｨ遉ｺ・・
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

' L蛛ｴ・壼酔讒倥↓繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ
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
            Debug.Print "[SENSE][MISS]"; keys(k); "・医ョ繝ｼ繧ｿ縺ｪ縺暦ｼ・
        End If
        pos = pos + 2
    Next k
    
    
    ' --- SENSE 蛯呵・ｒ隱ｭ縺ｿ霎ｼ縺ｿ・・ENSE_NOTE・・---
Dim cNote As Long, note As String
Dim box As Object, subCtl As Object, noteCtl As Object
Dim bestH As Single: bestH = 0

cNote = EnsureHeaderCol(ws, "SENSE_NOTE")
note = CStr(ws.Cells(r, cNote).value)

On Error Resume Next
' 繝壹・繧ｸ蜀・〒 MultiLine 縺ｾ縺溘・譛繧りレ縺ｮ鬮倥＞ TextBox 繧帝∈縺ｶ
For Each box In target.controls
    If TypeName(box) = "TextBox" Then
        If box.multiline Or box.Height > bestH Then
            Set noteCtl = box: bestH = box.Height
        End If
    ElseIf TypeName(box) = "Frame" Then
        For Each subCtl In box.controls
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
    noteCtl.text = note
    
Else
    Debug.Print "[SENSE][LOAD][NOTE] row=" & r & " col=" & cNote & " len=" & Len(note) & " (target textbox not found)"
End If

    
    
End Sub







