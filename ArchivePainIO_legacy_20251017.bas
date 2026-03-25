Attribute VB_Name = "ArchivePainIO_legacy_20251017"
'=== modPainIO ===
Option Private Module

Option Explicit



'窶補・險ｭ螳夲ｼ磯屁蠖｢縺ｮ4轤ｹ・壼ｿ・ｦ√↓蠢懊§縺ｦ蠕後〒蟾ｮ縺玲崛縺亥庄閭ｽ・・窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Const PAGE_HINT As String = "逍ｼ逞・            ' 蟇ｾ雎｡繧ｿ繝悶・隕句・縺励・荳驛ｨ・井ｾ具ｼ壹檎名逞帙阪檎李縺ｿ縲阪↑縺ｩ・・
Private Const HEADER_IO As String = "IO_Pain"         ' 譛ｬ菴薙・繧ｷ繝ｪ繧｢繝ｩ繧､繧ｺ蛻・
Private Const HEADER_NOTE As String = ""              ' 蛯呵・・・夂樟陦御ｻ墓ｧ倥〒縺ｯ蟒・ｭ｢
     ' 蛯呵・・・井ｸ崎ｦ√↑繧臥ｩｺ譁・ｭ励↓縺吶ｋ・・
Private keys As Variant                                ' R/L繝壹い蛹悶・隲也炊繧ｭ繝ｼ縲ょｿ・ｦ√↓蠢懊§縺ｦ蝗ｺ螳壼喧蜿ｯ
' 萓具ｼ壼ｾ後〒蠢・ｦ√↑繧・Array("VAS","PainQual","PainCourse","PainSite","PainFactors","PainDuration")

' 蛹ｺ蛻・ｊ・医ユ繝ｳ繝励Ξ譌｢螳夲ｼ・
Private Const SEP_REC As String = "|"  ' 繝ｬ繧ｳ繝ｼ繝牙玄蛻・ｊ
Private Const SEP_KV  As String = ":"  ' 繧ｭ繝ｼ縺ｨ蛟､
Private Const SEP_RL  As String = ","  ' R/L 騾｣邨・




'窶補・繝代ヶ繝ｪ繝・けAPI・井ｿ晏ｭ假ｼ・窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Public Sub SavePainToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)

    Dim pg As Object
    Set pg = ResolvePainPage(owner)
      If pg Is Nothing Then Exit Sub

    

    Dim combos As Collection
    Set combos = New Collection
    CollectCombos pg, combos  ' Page逶ｴ荳具ｼ祈rame蜀・・ComboBox繧貞・蟶ｰ蜿朱寔

    If combos.count = 0 Then
       

    End If

    ' Top/Left 繧ｽ繝ｼ繝茨ｼ・OL=6縺ｮ邁｡譏馴明蛟､・啜op蜆ｪ蜈遺・Left・・
    Dim arr() As Variant: arr = ControlsToArray(combos)
    SortByTopLeft arr, 6

    ' R/L 繝壹い繝ｪ繝ｳ繧ｰ縺励※繧ｷ繝ｪ繧｢繝ｩ繧､繧ｺ
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As String, base As String, side As String, valText As String

    For i = LBound(arr) To UBound(arr)
        nm = arr(i)("Name")
        valText = ComboValueText(arr(i)("Ref"))
        
        

        
        base = BaseNameRL(nm, side) ' side="R" or "L" or ""
        If Not dict.exists(base) Then dict.Add base, CreateObject("Scripting.Dictionary")
        If Len(side) = 0 Then side = "V" ' 蜊倡峡鬆・岼・・/L縺ｧ縺ｪ縺・ｼ・
        dict(base)(side) = valText
        ' 莉ｻ諢擾ｼ壽里遏･縺ｮ繧ｭ繝ｼ鬆・ｒ菴ｿ縺・ｴ蜷医・ KEYS 繧貞ｾ梧ｮｵ縺ｧ蛻ｩ逕ｨ
        
      


        
    Next

    Dim parts As Collection: Set parts = New Collection
    Dim k As Variant, rec As String, vR As String, vL As String, vV As String

    ' 譌｢螳夲ｼ壽､懷・鬆・ｼ・op/Left・峨〒蜃ｺ蜉帙ょ崋螳夐・↓縺励◆縺・ｴ蜷医・ KEYS 繧剃ｽｿ逕ｨ
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
    
        ' === 謖∫ｶ壽悄髢難ｼ域焚蟄暦ｼ嗾xtPainDuration・峨ｒ蜊倡峡繧ｭ繝ｼ縺ｨ縺励※菫晏ｭ・===
    Dim durText As String
    On Error Resume Next
    durText = CStr(pg.controls("txtPainDuration").text)
    On Error GoTo 0
    If Len(Trim$(durText)) > 0 Then
        parts.Add "txtPainDuration" & SEP_KV & " " & durText
    End If

    

' === ListBox・郁､・焚驕ｸ謚橸ｼ峨・菫晏ｭ倥ｒ霑ｽ蜉 ===
Dim listBoxes As Collection
Dim c As Object, sel As String, j As Long, base2 As String
Set listBoxes = New Collection
CollectListBoxesRecursive pg, listBoxes

For Each c In listBoxes
    If TypeName(c) = "ListBox" Then
        sel = ""
        For j = 0 To c.ListCount - 1
            If c.Selected(j) Then
                If Len(sel) > 0 Then sel = sel & "/"
                sel = sel & CStr(c.List(j))
            End If
        Next
        ' 菴輔ｂ驕ｸ縺ｰ繧後※縺・↑縺・ｴ蜷医・遨ｺ縺ｮ縺ｾ縺ｾ・亥・蜉帙＠縺ｪ縺・ｼ・
        If Len(sel) > 0 Then
            ' 萓具ｼ嗟stPainQual 竊・PainQual, lstPainSite 竊・PainSite
            base2 = c.name
            If LCase$(Left$(base2, 3)) = "lst" Then base2 = Mid$(base2, 4)
            
            'If base2 = "PainSite" Then sel = NormalizePainSite(sel)
            parts.Add base2 & SEP_KV & " " & sel
        End If
    End If
Next
' === 霑ｽ蜉縺薙％縺ｾ縺ｧ ===

' === CheckBox・・rue縺ｮ縺ｿ・峨ｒ縺ｾ縺ｨ繧√※菫晏ｭ・===
Dim factors As Collection: Set factors = New Collection
CollectChecksRecursive pg, factors
If factors.count > 0 Then
    Dim uniq As Object: Set uniq = CreateObject("Scripting.Dictionary")
    Dim ii As Long: For ii = 1 To factors.count: uniq(factors(ii)) = 1: Next
    parts.Add "PainFactors" & SEP_KV & " " & Join(uniq.keys, "/")
End If




' === 霑ｽ蜉縺薙％縺ｾ縺ｧ ===

' === VAS・・ 縺ｧ繧ゆｿ晏ｭ假ｼ・===
Dim vasText As String
On Error Resume Next
vasText = CStr(pg.controls("fraVAS").controls("txtVAS").text)   ' TextBox 蜆ｪ蜈・
If Len(vasText) = 0 Then vasText = CStr(pg.controls("fraVAS").controls("sldVAS").value)  ' ScrollBar 莉｣譖ｿ
On Error GoTo 0

' 縲・縲阪ｂ譛牙柑蛟､縺ｨ縺励※菫晏ｭ倥☆繧・
If (Len(vasText) > 0) Or vasText = "0" Then
    parts.Add "VAS" & SEP_KV & " " & vasText
End If
' === 霑ｽ蜉縺薙％縺ｾ縺ｧ ===


    Dim outText As String: outText = JoinCollection(parts, SEP_REC)
    Debug.Print "[IO-FINAL]", outText


    ' 蛯呵・ｼ・age蜀・・譛螟ｧTextBox繝・く繧ｹ繝茨ｼ・
    Dim noteText As String: noteText = LargestTextBoxValue(pg)


    Debug.Print "[IO-FINAL]", outText

    ' 譖ｸ縺崎ｾｼ縺ｿ
    ws.Cells(r, EnsureHeaderCol(ws, HEADER_IO)).value = outText
    If LenB(HEADER_NOTE) > 0 Then
        ws.Cells(r, EnsureHeaderCol(ws, HEADER_NOTE)).value = noteText
    End If
    
End Sub

Private Function ResolvePainPage(ByVal owner As Object) As Object
    Dim mpPhys As Object
    Dim i As Long
    Dim pg As Object

    Set mpPhys = modCommonUtil.SafeGetControl(owner, "mpPhys")
    If mpPhys Is Nothing Then Exit Function

    For i = 0 To mpPhys.Pages.count - 1
        Set pg = mpPhys.Pages(i)
        If Not modCommonUtil.SafeGetControl(pg, "fraPainCourse") Is Nothing Then
            Set ResolvePainPage = pg
            Exit Function
        End If
        If Not modCommonUtil.SafeGetControl(pg, "fraVAS") Is Nothing Then
            Set ResolvePainPage = pg
            Exit Function
        End If
    Next
End Function



'窶補・陬懷勧・壼ｯｾ雎｡MultiPage縺ｨPage謗｢邏｢ 窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Function FindTargetMultiPage(ByVal owner As Object, ByVal hint As String, ByRef outPage As Object) As Object
    Dim ctl As Object, mp As Object, i As Long
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            Set mp = ctl
            For i = 0 To mp.Pages.count - 1
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

'窶補・陬懷勧・啀age驟堺ｸ九・ComboBox繧貞・蟶ｰ蜿朱寔・・rame蜀・性繧・・窶補補補補補補補補補補補補補補補補補補補補補補補補・
Private Sub CollectCombos(ByVal container As Object, ByRef bag As Collection)
    Dim ctl As Object
    For Each ctl In container.controls
        Select Case TypeName(ctl)
            Case "ComboBox": bag.Add ctl
            Case "Frame":    CollectCombos ctl, bag
        End Select
    Next
End Sub

'窶補・陬懷勧・咾ombos竊帝・蛻暦ｼ・ame/Top/Left/Ref・・窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Sub CollectListBoxesRecursive(ByVal container As Object, ByRef bag As Collection)
    Dim ctl As Object
    For Each ctl In container.controls
        Select Case TypeName(ctl)
            Case "ListBox"
                bag.Add ctl
            Case "Frame", "Page"
                CollectListBoxesRecursive ctl, bag
        End Select
    Next
End Sub

Private Function ControlsToArray(ByVal bag As Collection) As Variant
    Dim i As Long, o As Object
    Dim arr() As Variant
    ReDim arr(0 To bag.count - 1)
    For i = 1 To bag.count
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

'窶補・陬懷勧・啌/L 蝓ｺ蠎募錐謚ｽ蜃ｺ 窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Function BaseNameRL(ByVal name As String, ByRef side As String) As String
    Dim s As String: s = UCase$(name)
    If Right$(s, 2) = "_R" Then side = "R": BaseNameRL = Left$(name, Len(name) - 2): Exit Function
    If Right$(s, 2) = "_L" Then side = "L": BaseNameRL = Left$(name, Len(name) - 2): Exit Function
    side = "": BaseNameRL = name
End Function

'窶補・陬懷勧・咾ombo縺ｮ蛟､蜿門ｾ暦ｼ・tyle=2蟇ｾ遲厄ｼ啖alue竊鱈istIndex竊但ddItem・・窶補補補補補補補補補補補補補補補補・
Private Function ComboValueText(ByVal cbo As Object) As String
    On Error Resume Next
    Dim t As String
    t = CStr(cbo.value)
    If LenB(t) = 0 Then
        If cbo.ListIndex >= 0 Then t = CStr(cbo.List(cbo.ListIndex))
    End If
    If LenB(t) = 0 Then
        ' 譛ｪ逋ｻ骭ｲ蛟､縺悟・縺｣縺ｦ縺・ｋ繧ｱ繝ｼ繧ｹ縺ｫ蛯吶∴縺ｦText繧りｦ九ｋ
        t = CStr(cbo.text)
    End If
    ComboValueText = t
End Function

'窶補・陬懷勧・壽怙螟ｧTextBox縺ｮ蛟､・亥ｙ閠・Φ螳夲ｼ・窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Function LargestTextBoxValue(ByVal container As Object) As String
    Dim ctl As Object, area As Double, maxArea As Double, best As Object
    For Each ctl In container.controls
        If TypeName(ctl) = "TextBox" Then
            area = ctl.Width * ctl.Height
            If area > maxArea Then maxArea = area: Set best = ctl
        ElseIf TypeName(ctl) = "Frame" Then
            Dim s As String: s = LargestTextBoxValue(ctl)
            If LenB(s) > 0 And area = 0 Then LargestTextBoxValue = s ' 繝阪せ繝亥・縺ｧ豎ｺ縺ｾ縺｣縺溘ｉ謗｡逕ｨ
        End If
    Next
    If Not best Is Nothing Then LargestTextBoxValue = CStr(best.text)
End Function

'窶補・蟆冗黄繝ｦ繝ｼ繝・ぅ繝ｪ繝・ぅ 窶補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補補・
Private Function NzS(ByVal dict As Object, ByVal k As String) As String
    If dict.exists(k) Then NzS = CStr(dict(k)) Else NzS = ""
End Function

Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, s() As String
    ReDim s(1 To c.count)
    For i = 1 To c.count: s(i) = CStr(c(i)): Next
    JoinCollection = Join(s, sep)
End Function

Private Function CreateMap4(k1 As String, v1 As Variant, k2 As String, v2 As Variant, k3 As String, v3 As Variant, k4 As String, v4 As Variant) As Variant
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add k1, v1: d.Add k2, v2: d.Add k3, v3: d.Add k4, v4
    Set CreateMap4 = d
End Function






Private Sub CollectChecksRecursive(parent As Object, colL As Collection)
    Dim c As Object, nm2 As String
    For Each c In parent.controls
        If TypeName(c) = "CheckBox" Then
            If c.value = True Then
                nm2 = c.name
                If LCase$(Left$(nm2, 3)) = "chk" Then nm2 = Mid$(nm2, 4)
                colL.Add nm2
            End If
        ElseIf TypeName(c) = "Frame" Or TypeName(c) = "Page" Then
            CollectChecksRecursive c, colL
        End If
    Next
End Sub



Public Sub DumpPainFrames_Once()
    Dim pg As Object, f As Object, c As Object
    Dim uf As Object: Set uf = frmEval
Set pg = uf.mpPhys.Pages(4)

    Debug.Print "[Page]", pg.caption
    For Each f In pg.controls
        If TypeName(f) = "Frame" Then
            Debug.Print "[Frame]", f.name, "count", f.controls.count
            For Each c In f.controls
                If TypeName(c) = "CheckBox" Then Debug.Print "  [Chk]", c.name, c.value
                If TypeName(c) = "Frame" Then Debug.Print "  [SubFrame]", c.name, "count", c.controls.count
            Next
        End If
    Next
End Sub



Public Sub SavePain_CheckOnce()
    Dim uf As Object: Set uf = frmEval
    ' 繝√ぉ繝・け1縺､ON・郁ｪ伜屏・壼虚菴懊〒蠅玲が・・
    uf.controls("mpPhys").Pages(4).controls("fraPainFactors").controls("chkPainProv_Move").value = True
    ' 菫晏ｭ假ｼ医・繧ｿ繝ｳ逶ｸ蠖難ｼ・
    SaveEvaluation_Append_From uf
    ' 逶ｴ霑題｡後・IO/NOTE繧呈焚蛟､陦ｨ遉ｺ
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub



Public Sub SavePain_AppendTest_Once()
    Dim uf As Object: Set uf = frmEval
    On Error Resume Next
    uf.txtName.text = "讀懆ｨｼAppend"
    uf.controls("chkDiffOnly").value = False
    On Error GoTo 0

    ' 菫晏ｭ假ｼ亥・菴謎ｿ晏ｭ倥Ν繝ｼ繝茨ｼ・
    SaveEvaluation_Append_From uf

    ' 逶ｴ霑題｡後・IO/NOTE繧貞庄隕門喧
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub


Public Sub Test_SaveAtRow_Once()
    Dim ws As Worksheet, rr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    rr = 108   ' 竊・莉ｻ諢上・讀懆ｨｼ陦・

    'SaveAllSectionsToSheet ws, rr, frmEval

    Debug.Print "[WroteRow]", rr
   'Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 107).value), 180)
   'Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 108).value), 120)

End Sub



Public Sub SavePain_FillAndAppend_Once()
    Dim uf As Object: Set uf = frmEval
    On Error Resume Next
    uf.txtName.text = "讀懆ｨｼAppend3"
    uf.controls("chkDiffOnly").value = False
    With uf.controls("mpPhys").Pages(4)
        .controls("cmbPainOnset").ListIndex = 0
        .controls("cmbPainDurationUnit").ListIndex = 0
        .controls("cmbPainDayPeriod").ListIndex = 2
        Dim lb As MSForms.ListBox
        Set lb = .controls("lstPainQual"): If lb.ListCount > 0 Then lb.Selected(0) = True
        Set lb = .controls("lstPainSite"): If lb.ListCount > 0 Then lb.Selected(0) = True
        .controls("fraPainFactors").controls("chkPainProv_Move").value = True
    End With
    On Error GoTo 0

    SaveEvaluation_Append_From uf

    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets("EvalData")
    lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Debug.Print "[LastRow]", lr
    Debug.Print "[IO]", Left$(CStr(ws.Cells(lr, 156).value), 180)
    Debug.Print "[NOTE]", Left$(CStr(ws.Cells(lr, 157).value), 120)
End Sub





Private Function GetCtlVal(o As Object) As String
    On Error Resume Next
    GetCtlVal = "" & o.value
    If Len(GetCtlVal) = 0 Then GetCtlVal = "" & o.text
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

'=== [TEMP] 菫晏ｭ倡ｳｻ繝上Φ繝峨Λ迚ｹ螳壹せ繧ｭ繝｣繝・================================
Public Sub Scan_SaveHandlers()
    Dim vbComp As Object, cm As Object
    Dim lineCount As Long, i As Long
    Dim pat1 As String, pat2 As String, pat3 As String
    pat1 = "菫晏ｭ・         ' 繝｡繝・そ繝ｼ繧ｸ譁・ｨ
    pat2 = "Save"         ' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷阪く繝ｼ繝ｯ繝ｼ繝・
    pat3 = "Application.Caller" ' 繧ｷ繝ｼ繝医・繧ｿ繝ｳ諠ｳ螳・

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




