Attribute VB_Name = "modCommonUtil"
Option Explicit

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long

Public Sub Tmp_ShowExcelPID()
    Dim pid As Long
    GetWindowThreadProcessId Application.hwnd, pid
    Debug.Print "[PID] ExcelPID=" & pid & "  Hwnd=" & Application.hwnd
End Sub

Public Sub Cleanup_DuplicateROMHeaders_KeepRightmost()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim lastCol As Long, c As Long, h As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary") ' key: header, value: Collection of columns

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' 1) ROM_* 縺ｮ菴咲ｽｮ繧貞・襍ｰ譟ｻ
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If Len(h) > 0 Then
            If LCase$(Left$(h, 4)) = "rom_" Then
                If Not map.exists(h) Then Set map(h) = New Collection
                map(h).Add c
            End If
        End If
    Next c

    ' 2) 蜿ｳ遶ｯ(譛螟ｧ蛻・縺縺第ｮ九＠縲∽ｻ悶・蜑企勁蟇ｾ雎｡縺ｨ縺励※蜿朱寔
    Dim toDel As New Collection
    Dim k As Variant, i As Long, keepCol As Long
    For Each k In map.keys
        If map(k).count > 1 Then
            keepCol = -1
            For i = 1 To map(k).count
                If map(k)(i) > keepCol Then keepCol = map(k)(i)
            Next i
            For i = 1 To map(k).count
                If map(k)(i) <> keepCol Then toDel.Add CLng(map(k)(i))
            Next i
        End If
    Next k

    ' 3) 髯埼・〒蜑企勁
    Dim arr() As Long, n As Long
    n = toDel.count
    If n > 0 Then
        ReDim arr(1 To n)
        For i = 1 To n
            arr(i) = toDel(i)
        Next i
        ' 髯埼・た繝ｼ繝・
        Dim j As Long, tmp As Long
        For i = 1 To n - 1
            For j = i + 1 To n
                If arr(i) < arr(j) Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            Next j
        Next i
        ' 蜑企勁螳溯｡・
        For i = 1 To n
            Debug.Print "[DUP-CLEAN] delete col", arr(i), "(" & ws.Cells(1, arr(i)).value & ")"
            ws.Columns(arr(i)).Delete
        Next i
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' 隕句・縺励′螳悟・荳閾ｴ縺吶ｋ縲御ｸ逡ｪ蜿ｳ縺ｮ蛻礼分蜿ｷ縲阪ｒ霑斐☆繝ｦ繝ｼ繝・ぅ繝ｪ繝・ぅ
Public Function HeaderCol_Compat_Rightmost(ByVal name As String, ByVal ws As Worksheet) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = lastCol To 1 Step -1
        If StrComp(CStr(ws.Cells(1, c).value), name, vbTextCompare) = 0 Then
            HeaderCol_Compat_Rightmost = c
            Exit Function
        End If
    Next c
End Function

' IO騾｣邨先枚蟄怜・ "key=value|key=value|..." 縺九ｉ縲∵欠螳嗅ey縺ｮ value 繧定ｿ斐☆邁｡譏薙ヱ繝ｼ繧ｵ
Public Function GetIOValue(ByVal ioStr As String, ByVal key As String) As String
    Dim token As Variant, klen As Long
    klen = Len(key) + 1 ' "key=" 縺ｮ髟ｷ縺・
    For Each token In Split(ioStr, "|")
        If Left$(token, klen) = key & "=" Then
            GetIOValue = Mid$(token, klen + 1) ' "=" 縺ｮ蠕後ｍ
            Exit Function
        End If
    Next token
End Function

' "key=value|key:...|" 縺ｮ豺ｷ蝨ｨ繧呈Φ螳壹・
' 謖・ｮ・key 縺ｮ蜿ｳ蛛ｴ・・ 縺ｾ縺溘・ : 縺ｮ蠕後ｍ・峨ｒ窶懊◎縺ｮ縺ｾ縺ｾ窶晁ｿ斐☆縲・
Public Function GetIOChunk(ByVal ioStr As String, ByVal key As String) As String
    Dim token As Variant, t As String, k1 As String, k2 As String, p As Long
    k1 = key & "=": k2 = key & ":"
    For Each token In Split(ioStr, "|")
        t = Trim$(CStr(token))
        If Left$(t, Len(k1)) = k1 Then
            GetIOChunk = Mid$(t, Len(k1) + 1)
            Exit Function
        ElseIf Left$(t, Len(k2)) = k2 Then
            GetIOChunk = Mid$(t, Len(k2) + 1)
            Exit Function
        End If
    Next token
End Function

' 萓・ chunk="R=,L=豸亥､ｱ" -> GetIOSubVal(chunk,"R")="" / GetIOSubVal(chunk,"L")="豸亥､ｱ"
Public Function GetIOSubVal(ByVal chunk As String, ByVal subkey As String) As String
    Dim parts As Variant, p As Variant, k As String, pos As Long, tail As String, nextPos As Long
    k = subkey & "="
    ' 繧ｫ繝ｳ繝槫玄蛻・ｊ縺ｧ襍ｰ譟ｻ・・=... , L=... , 縺ｪ縺ｩ・・
    parts = Split(chunk, ",")
    For Each p In parts
        p = Trim$(CStr(p))
        If Left$(p, Len(k)) = k Then
            tail = Mid$(p, Len(k) + 1)
            ' 繧ゅ＠ "R=xxx L=yyy" 縺ｿ縺溘＞縺ｫ繧ｫ繝ｳ繝樒┌縺励〒邯壹￥蝣ｴ蜷医↓蛯吶∴縲∵ｬ｡縺ｮ繧ｹ繝壹・繧ｹ縺ｾ縺ｧ繧貞､縺ｨ縺ｿ縺ｪ縺・
            nextPos = InStr(1, tail, " ", vbBinaryCompare)
            If nextPos > 0 Then
                GetIOSubVal = Left$(tail, nextPos - 1)
            Else
                GetIOSubVal = tail
            End If
            Exit Function
        End If
    Next p
End Function

Public Sub App_Main()

    ActiveWindow.Zoom = 100
    Application.WindowState = xlMaximized

    Load frmEval

    With frmEval
        .StartUpPosition = 0
        .Left = 0
        .Top = 0

        ' 逕ｻ髱｢縺ｫ蜿弱∪繧九ｈ縺・↓荳企剞
        If .Height > Application.UsableHeight + 156 Then .Height = Application.UsableHeight + 156

    End With

    ' 笘・蔵 縺ｾ縺夊｡ｨ遉ｺ・医％縺薙〒 InsideHeight 縺檎｢ｺ螳夲ｼ・
    frmEval.Show vbModeless

        Dim yBtn As Single
    yBtn = frmEval.InsideHeight - frmEval.controls("btnCloseCtl").Height - 12

    frmEval.controls("btnCloseCtl").Top = yBtn
    frmEval.controls("cmdSaveGlobal").Top = yBtn
    frmEval.controls("cmdClearGlobal").Top = yBtn

    frmEval.controls("mpPhys").Height = yBtn - frmEval.controls("mpPhys").Top - 12
    
    Debug.Print "[post-mpPhys] yBtn=" & yBtn & " mpPhysB=" & (frmEval.controls("mpPhys").Top + frmEval.controls("mpPhys").Height) & " InsideH=" & frmEval.InsideHeight

    
    
    If frmEval.Height > Application.UsableHeight - 40 Then frmEval.Height = Application.UsableHeight - 40

Call frmEval.AdjustBottomButtons


End Sub

'=== Basic.* 縺ｨ譌ｧ蛻暦ｼ域ｰ丞錐/隧穂ｾ｡譌･ 縺ｪ縺ｩ・峨ｒ 1 陦悟・縺縺大酔譛溘☆繧・=====================
Public Sub SyncBasicInfoColumns(ws As Worksheet, ByVal r As Long)
    Dim headersBasic As Variant
    Dim headersLegacy As Variant
    Dim i As Long
    Dim cb As Long, cL As Long
    Dim vB As Variant, vL As Variant

    ' Basic.* 邉ｻ繧偵梧ｭ｣縲阪→縺ｿ縺ｪ縺吶′縲・
    ' 迚・婿縺励°蜈･縺｣縺ｦ縺・↑縺・ｴ蜷医・縲∝・縺｣縺ｦ縺・ｋ譁ｹ縺九ｉ繧ゅ≧迚・婿縺ｸ繧ｳ繝斐・縺吶ｋ縲・
     headersBasic = Array("Basic.EvalDate", "Basic.Name", "Basic.Age", "Basic.Evaluator")
     headersLegacy = Array("隧穂ｾ｡譌･", "豌丞錐", "蟷ｴ鮨｢", "隧穂ｾ｡閠・)


    For i = LBound(headersBasic) To UBound(headersBasic)
        cb = modEvalIOEntry.FindColByHeaderExact(ws, headersBasic(i))
        cL = modEvalIOEntry.FindColByHeaderExact(ws, headersLegacy(i))


        ' 縺ｩ縺｡繧峨°縺ｮ蛻励′蟄伜惠縺励※縺・ｌ縺ｰ蜷梧悄蟇ｾ雎｡
        If cb > 0 Or cL > 0 Then
            If cb > 0 Then
                vB = ws.Cells(r, cb).value
            Else
                vB = vbNullString
            End If

            If cL > 0 Then
                vL = ws.Cells(r, cL).value
            Else
                vL = vbNullString
            End If

            ' 蜆ｪ蜈亥ｺｦ・・
            ' 1) Basic 蛛ｴ縺ｫ蛟､縺後≠縺｣縺ｦ譌ｧ蛻励′遨ｺ 竊・Basic 竊・譌ｧ蛻励∈繧ｳ繝斐・
            ' 2) Basic 蛛ｴ縺檎ｩｺ縺ｧ譌ｧ蛻励↓蛟､ 竊・譌ｧ蛻・竊・Basic 縺ｸ繧ｳ繝斐・
            If cb > 0 And Len(vB) > 0 And cL > 0 And Len(vL) = 0 Then
                ws.Cells(r, cL).value = vB
            ElseIf cL > 0 And Len(vL) > 0 And cb > 0 And Len(vB) = 0 Then
                ws.Cells(r, cb).value = vL
            End If
        End If
    Next i
End Sub
'====================================================================


Public Function ControlExists(parent As Object, ctrlName As String) As Boolean
    Dim c As Object
    For Each c In parent.controls
        If c.name = ctrlName Then
            ControlExists = True
            Exit Function
        End If
    Next
    ControlExists = False
End Function

Public Function SafeGetControl(ByVal parent As Object, ByVal nm As String) As Object
    
    Dim visited As Object

    If parent Is Nothing Then Exit Function
    If LenB(Trim$(nm)) = 0 Then Exit Function

    Set visited = CreateObject("Scripting.Dictionary")
    Set SafeGetControl = FindControlRecursive(parent, nm, visited)
End Function

Private Function FindControlRecursive(ByVal node As Object, ByVal targetName As String, ByVal visited As Object) As Object
    Dim child As Object
    Dim page As Object
    Dim found As Object
    Dim key As String

    If node Is Nothing Then Exit Function

    key = BuildVisitedKey(node)
    If LenB(key) > 0 Then
        If visited.exists(key) Then Exit Function
        visited.Add key, True
    End If

    If HasControlName(node, targetName) Then
        Set FindControlRecursive = node
        Exit Function
    End If

    If HasControls(node) Then
        For Each child In GetChildControls(node)
            Set found = FindControlRecursive(child, targetName, visited)
            If Not found Is Nothing Then
                Set FindControlRecursive = found
                Exit Function
            End If
        Next child
    End If

    If TypeName(node) = "MultiPage" Then
        On Error Resume Next
        For Each page In node.Pages
            On Error GoTo 0
            Set found = FindControlRecursive(page, targetName, visited)
            If Not found Is Nothing Then
                Set FindControlRecursive = found
                Exit Function
            End If
            On Error Resume Next
        Next page
        On Error GoTo 0
    End If
End Function

Private Function GetChildControls(ByVal parent As Object) As Collection
    Dim result As New Collection
    Dim c As Object
    
    On Error Resume Next

    For Each c In parent.controls
        result.Add c
    Next c
    On Error GoTo 0

    Set GetChildControls = result
End Function

Private Function HasControlName(ByVal ctrl As Object, ByVal nm As String) As Boolean
    Dim ctrlName As String

    On Error Resume Next
    ctrlName = CStr(ctrl.name)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0
    

    HasControlName = (StrComp(ctrlName, nm, vbTextCompare) = 0)
End Function

Private Function BuildVisitedKey(ByVal node As Object) As String
    Dim h As String
    Dim n As String

    On Error Resume Next
    h = Hex$(ObjPtr(node))
    If Err.Number <> 0 Then
        Err.Clear
        h = ""
    End If

    n = CStr(node.name)
    If Err.Number <> 0 Then
        Err.Clear
        n = ""
    End If
    On Error GoTo 0

    BuildVisitedKey = TypeName(node) & "|" & h & "|" & n
End Function



Public Sub Tighten_DailyLog_Boxes()
    Dim uf As Object: Set uf = frmEval

    Dim mp As Object: Set mp = uf.controls("MultiPage1")
    Dim pg As Object: Set pg = mp.Pages(7) ' 譌･縲・・險倬鹸

    Dim f As MSForms.Frame: Set f = pg.controls("fraDailyLog")
    Dim txtTraining As MSForms.TextBox: Set txtTraining = f.controls("txtDailyTraining")
    Dim txtReaction As MSForms.TextBox: Set txtReaction = f.controls("txtDailyReaction")
    Dim txtAbnormal As MSForms.TextBox: Set txtAbnormal = f.controls("txtDailyAbnormal")
    Dim txtPlan As MSForms.TextBox: Set txtPlan = f.controls("txtDailyPlan")
    Dim lst As MSForms.ListBox: Set lst = f.controls("lstDailyLogList")

    Const BOX_H As Single = 95

    txtTraining.Height = BOX_H
    txtReaction.Height = BOX_H
    txtAbnormal.Height = BOX_H
    txtPlan.Height = BOX_H

    Dim fieldsBottom As Single
    fieldsBottom = Application.Max(txtAbnormal.Top + txtAbnormal.Height, txtPlan.Top + txtPlan.Height)



    ' 繝ｩ繝吶Ν繧偵御ｸ隕ｧ縺ｮ逶ｴ荳翫阪↓鄂ｮ縺・
    Dim lbl As MSForms.label
    Set lbl = f.controls("lblDailyHistory")


    ' ListBox縺ｯ貅｢繧後◆繧芽・蜍輔〒繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ縺悟・繧具ｼ亥ｸｸ譎り｡ｨ遉ｺ縺ｯ莉墓ｧ倅ｸ翫〒縺阪↑縺・ｼ・
    lbl.Top = fieldsBottom + 15
    lst.Top = lbl.Top + lbl.Height + 4
    lst.Height = Application.Max(60, f.Height - lst.Top - 8)
    lst.IntegralHeight = False


End Sub

Public Function HasControls(ByVal o As Object) As Boolean
    Dim n As Long
    
    On Error Resume Next
    n = o.controls.count
    
    If Err.Number <> 0 Then
        Err.Clear
        HasControls = False
    Else
        HasControls = (n >= 0)
    End If
    On Error GoTo 0
End Function

Public Sub Verify_POST_TagUniqueness()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim dup As Long: dup = 0

    Dim p As Object, f2 As Object, f35 As Object, f36 As Object
    Set p = frmEval.controls("MultiPage1").Pages("Page2")
    Set f2 = p.controls("Frame2")
    Set f35 = f2.controls("Frame35")
    Set f36 = f2.controls("Frame36")

    dup = dup + CountDupTagsInFrame(seen, f35)
    dup = dup + CountDupTagsInFrame(seen, f36)

    Debug.Print "[VERIFY POST TAG UNIQUE] DUP=" & dup & " UNIQUE=" & seen.count
End Sub

Private Function CountDupTagsInFrame(ByVal seen As Object, ByVal fr As Object) As Long
    Dim c As Object, t As String
    Dim dup As Long: dup = 0

    For Each c In fr.controls
        If c.parent Is fr Then
            If TypeName(c) = "CheckBox" Or TypeName(c) = "ComboBox" Or TypeName(c) = "TextBox" Or TypeName(c) = "OptionButton" Then
                t = ""
                On Error Resume Next
                t = CStr(c.tag)
                On Error GoTo 0

                If Len(t) > 0 Then
                    If seen.exists(t) Then
                        dup = dup + 1
                        Debug.Print "DUP TAG=" & t
                    Else
                        seen.Add t, True
                    End If
                Else
                    Debug.Print "EMPTY TAG at " & TypeName(c) & " L=" & c.Left & " T=" & c.Top
                End If
            End If
        End If
    Next c

    CountDupTagsInFrame = dup
End Function

Public Sub Build_POST_Narrative()
    Dim posture As String, contr As String, note As String, noteC As String

    posture = CollectTrueTags("Posture.")
    contr = CollectTrueTags("Contracture.")

    note = GetTagText("Posture.蛯呵・)
    noteC = GetTagText("Contracture.蛯呵・)

    If posture = "" Then posture = "蟋ｿ蜍｢・夂音險倥↑縺・ Else posture = "蟋ｿ蜍｢・・ & posture
    If contr = "" Then contr = "諡倡ｸｮ・夂音險倥↑縺・ Else contr = "諡倡ｸｮ・・ & contr

    Debug.Print posture
    Debug.Print contr
    If note <> "" Then Debug.Print "蟋ｿ蜍｢蛯呵・ｼ・ & note
    If noteC <> "" Then Debug.Print "諡倡ｸｮ蛯呵・ｼ・ & noteC
End Sub

Private Function CollectTrueTags(ByVal prefix As String) As String
    Dim p As Object, f2 As Object, f35 As Object, f36 As Object
    Set p = frmEval.controls("MultiPage1").Pages("Page2")
    Set f2 = p.controls("Frame2")
    Set f35 = f2.controls("Frame35")
    Set f36 = f2.controls("Frame36")

    Dim res As String
    res = res & CollectInFrame(f35, prefix)
    res = res & CollectInFrame(f36, prefix)

    If Len(res) > 0 Then
        If Right$(res, 1) = "縲・ Then res = Left$(res, Len(res) - 1)
    End If
    CollectTrueTags = res
End Function

Private Function CollectInFrame(ByVal fr As Object, ByVal prefix As String) As String
    Dim c As Object, t As String, s As String
    For Each c In fr.controls
        If c.parent Is fr Then
            If TypeName(c) = "CheckBox" Then
                t = CStr(c.tag)
                If Len(t) > 0 And Left$(t, Len(prefix)) = prefix Then
                    If c.value = True Then s = s & Replace$(t, prefix, "") & "縲・
                End If
            End If
        End If
    Next
    CollectInFrame = s
End Function

Private Function GetTagText(ByVal tagName As String) As String
    Dim p As Object, f2 As Object, fr As Object, c As Object
    Set p = frmEval.controls("MultiPage1").Pages("Page2")
    Set f2 = p.controls("Frame2")

    ' Frame35
    Set fr = f2.controls("Frame35")
    For Each c In fr.controls
        If c.parent Is fr Then
            If TypeName(c) = "TextBox" And CStr(c.tag) = tagName Then
                GetTagText = Replace$(CStr(c.text), vbCrLf, " ")
                Exit Function
            End If
        End If
    Next

    ' Frame36
    Set fr = f2.controls("Frame36")
    For Each c In fr.controls
        If c.parent Is fr Then
            If TypeName(c) = "TextBox" And CStr(c.tag) = tagName Then
                GetTagText = Replace$(CStr(c.text), vbCrLf, " ")
                Exit Function
            End If
        End If
    Next
End Function


Public Function AsLongArray(ByVal v As Variant) As Variant
    If IsEmpty(v) Then Exit Function
    If IsArray(v) Then AsLongArray = v: Exit Function
    If IsObject(v) Then
        Dim i As Long, arr() As Long
        ReDim arr(1 To v.count)
        For i = 1 To v.count
            arr(i) = CLng(v(i))
        Next
        AsLongArray = arr
    End If
End Function

Public Function NormalizeName(ByVal s As String) As String
    s = Replace$(CStr(s), " ", "")   ' 蜊願ｧ偵せ繝壹・繧ｹ髯､蜴ｻ
    s = Replace$(s, "縲", "")        ' 蜈ｨ隗偵せ繝壹・繧ｹ髯､蜴ｻ
    NormalizeName = s
End Function



' 豌丞錐繧貞魂蛻ｷ逕ｨ縺ｫ莨丞ｭ怜喧縺吶ｋ
' 4?5譁・ｭ暦ｼ・繝ｻ4譁・ｭ礼岼繧偵・ｼ医Θ繝ｼ繧ｶ繝ｼ謖・ｮ夲ｼ・
Public Function MaskNameForPrint(ByVal s As String, Optional ByVal maskChar As String = "縲・) As String
    Dim n As Long, i As Long
    Dim out As String, ch As String

    s = Trim$(s)
    n = Len(s)

    If n <= 0 Then
        MaskNameForPrint = ""
        Exit Function
    End If

    For i = 1 To n
        ch = Mid$(s, i, 1)

        Select Case n
            Case 1
                out = out & ch
            Case 2
                If i = 2 Then out = out & maskChar Else out = out & ch
            Case 3
                If i = 2 Then out = out & maskChar Else out = out & ch
            Case 4, 5
                If (i = 2) Or (i = 4) Then out = out & maskChar Else out = out & ch
            Case Else
                ' 6譁・ｭ嶺ｻ･荳奇ｼ壼・謨ｰ菴咲ｽｮ繧剃ｼ丞ｭ暦ｼ・,4,6,...・・
                If (i Mod 2 = 0) Then out = out & maskChar Else out = out & ch
        End Select
    Next i

    MaskNameForPrint = out
End Function



' 譛亥腰菴阪・繝悶ャ繧ｯ縺ｫ縲∝茜逕ｨ閠・す繝ｼ繝医ｒ霑ｽ蜉縺励※縺・￥
' ymKey 萓・ "2026-02" 縺ｪ縺ｩ・域怦蜊倅ｽ阪〒荳諢上↓縺ｪ繧区枚蟄怜・・・
Public Sub ExportMonitoring_ToMonthlyWorkbook(ByVal dailyDate As Date, ByVal clientName As String, ByVal bodyText As String)
    
    If Len(Trim$(clientName)) = 0 Then
        clientName = frmEval.controls("frHeader").controls("txtHdrName").text
    End If

    
    
    
    Dim ymKey As String: ymKey = Format(dailyDate, "yyyy-mm")
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim saveName As String
    Dim savePath As String


    saveName = "Monitoring_" & ymKey & ".xlsx"

    Dim baseFolder As String
Dim monthFolder As String

baseFolder = ThisWorkbook.path & "\Monitoring"
monthFolder = baseFolder & "\" & ymKey

If Dir$(baseFolder, vbDirectory) = "" Then
    MkDir baseFolder
End If

If Dir$(monthFolder, vbDirectory) = "" Then
    MkDir monthFolder
End If

savePath = monthFolder & "\" & saveName



If Len(Dir$(savePath)) > 0 Then
    Set wbNew = Workbooks.Open(savePath)
Else
    ThisWorkbook.Worksheets("Monitoring").Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
End If


    ' 譌｢縺ｫ蜷悟錐繧ｷ繝ｼ繝医′縺ゅｌ縺ｰ蜑企勁縺励※菴懊ｊ逶ｴ縺暦ｼ井ｸ頑嶌縺埼°逕ｨ・・
    If Len(Trim$(clientName)) > 0 Then
    On Error Resume Next
    Application.DisplayAlerts = False
    wbNew.Worksheets(clientName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End If


    ' 繝・Φ繝励Ξ Monitoring 繧偵％縺ｮ譛医ヶ繝・け縺ｸ繧ｳ繝斐・
    ThisWorkbook.Worksheets("Monitoring").Copy After:=wbNew.Worksheets(wbNew.Worksheets.count)
    Set wsNew = wbNew.Worksheets(wbNew.Worksheets.count)
    
    
    Dim c As Long, firstCol As Long, lastCol As Long
With ThisWorkbook.Worksheets("Monitoring").UsedRange
    firstCol = .Column
    lastCol = .Column + .Columns.count - 1
End With

With ThisWorkbook.Worksheets("Monitoring")
    For c = firstCol To lastCol
        wsNew.Columns(c).ColumnWidth = .Columns(c).ColumnWidth
    Next c
End With





    ' 繧ｷ繝ｼ繝亥錐縺ｯ螳溷錐縺ｧOK
If Trim$(clientName) = "" Then
    MsgBox "蛻ｩ逕ｨ閠・錐繧貞・蜉帙＠縺ｦ縺上□縺輔＞", vbExclamation
    Exit Sub
End If

    ' 蜊ｰ蛻ｷ迚ｩ縺ｮ陦ｨ遉ｺ蜷阪□縺台ｼ丞ｭ暦ｼ医す繝ｼ繝亥錐縺ｯ螳溷錐・・
    wsNew.Range("C7").value = MaskNameForPrint(clientName)

    ' 譛ｬ譁・
    wsNew.Range("A32:J37").Merge
    Dim p As Long
Dim s As String

s = bodyText
p = InStr(1, s, "笆 繧ｳ繝｡繝ｳ繝医・閠・ｯ・, vbTextCompare)

If p > 0 Then
    s = Mid$(s, p + Len("笆 繧ｳ繝｡繝ｳ繝医・閠・ｯ・))
    s = Trim$(s)
Else
    s = bodyText ' 隕九▽縺九ｉ縺ｪ縺・凾縺ｯ菫晞匱縺ｧ蜈ｨ譁・
End If


Dim p1 As Long, p2 As Long
Dim tok As String, special As String

tok = "笆 縺薙・譛医↓險倬鹸縺輔ｌ縺溽音險倅ｺ矩・
p1 = InStr(1, bodyText, tok, vbTextCompare)

If p1 > 0 Then
    special = Mid$(bodyText, p1 + Len(tok))
    ' 谺｡縺ｮ隕句・縺励〒豁｢繧√ｋ・亥呵｣懶ｼ・
    p2 = InStr(1, special, "笆 繧ｳ繝｡繝ｳ繝医・閠・ｯ・, vbTextCompare)
    If p2 = 0 Then p2 = InStr(1, special, "笆 譛ｬ譁・, vbTextCompare)
    If p2 > 0 Then special = Left$(special, p2 - 1)
    special = Trim$(special)
Else
    special = ""
End If



Dim tmplNoSpecial As String
tmplNoSpecial = "縺薙・譛医・迚ｹ險倅ｺ矩・→縺ｪ繧玖ｨ倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ縺ｧ縺励◆縲・ & vbCrLf & _
                "菴楢ｪｿ髱｢縺ｫ螟ｧ縺阪↑螟牙虚縺ｯ縺ｪ縺上∵律縲・・繝ｪ繝上ン繝ｪ縺ｫ繧ょｮ牙ｮ壹＠縺ｦ蜿悶ｊ邨・∪繧後※縺・∪縺励◆縲・ & vbCrLf & _
                "莉雁ｾ後ｂ迴ｾ蝨ｨ縺ｮ迥ｶ諷九ｒ邯ｭ謖√〒縺阪ｋ繧医≧縲∝ｼ輔″邯壹″邨碁℃繧定ｦｳ蟇溘＠縺ｦ縺・″縺ｾ縺吶・

If Len(Trim$(special)) = 0 Or InStr(special, "迚ｹ險倅ｺ矩・→縺ｪ繧玖ｨ倬鹸縺ｯ縺ゅｊ縺ｾ縺帙ｓ") > 0 Then
    
    ' 迚ｹ險倅ｺ矩・↑縺・
    wsNew.Range("A24").value = tmplNoSpecial
    wsNew.Range("A31:J37").ClearContents

Else
    
    ' 迚ｹ險倅ｺ矩・≠繧・
    wsNew.Range("A24").value = special
    wsNew.Range("A31").value = s

End If



wsNew.Range("A24").WrapText = True
With wsNew.Range("A24:J29")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With






With wsNew.Range("A31:J37")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With


    'wsNew.Range("A32").value = bodyText
    
    wsNew.Range("A24:J29").Merge
    With wsNew.Range("A31: J37 ")
      .Font.name = "・ｭ・ｳ ・ｰ繧ｴ繧ｷ繝・け"
      .Font.Size = 11
    End With
    
    
   Application.PrintCommunication = False

 With wsNew.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .Zoom = 100

 End With
Application.PrintCommunication = True


    

    
    wsNew.Range("A31").WrapText = True
    
    wsNew.ResetAllPageBreaks

    Dim r As Long, lastRow As Long
With ThisWorkbook.Worksheets("Monitoring").UsedRange
    lastRow = .row + .rows.count - 1
End With

With ThisWorkbook.Worksheets("Monitoring")
    For r = 1 To lastRow
        wsNew.rows(r).RowHeight = .rows(r).RowHeight
    Next r
End With

With wsNew.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .Zoom = False
End With




    wbNew.Save
wbNew.Close SaveChanges:=True



End Sub



