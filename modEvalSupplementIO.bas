Attribute VB_Name = "modEvalSupplementIO"
Option Explicit
Public gArchiveDeleteBasicID As String



'--- 繝倥ャ繝陦後°繧牙・逡ｪ蜿ｷ繧呈爾縺呻ｼ郁ｦ九▽縺九ｉ縺ｪ縺代ｌ縺ｰ 0・・
Private Function HeaderCol(ByVal hdrRow As Range, ByVal candidates As Variant) As Long
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim c As Range
        For Each c In hdrRow.Cells
            If Trim$(CStr(c.value)) = CStr(candidates(i)) Then
                HeaderCol = c.Column - hdrRow.Cells(1, 1).Column + 1 'CurrentRegion蜀・・逶ｸ蟇ｾ蛻・
                Exit Function
            End If
        Next c
    Next i
    HeaderCol = 0
End Function

'--- 繧｢繝ｼ繧ｫ繧､繝悶ヶ繝・け縺ｫ蜷悟錐繧ｷ繝ｼ繝医′縺ｪ縺代ｌ縺ｰ菴懊ｋ
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        On Error Resume Next
        ws.name = sheetName
        On Error GoTo 0
    End If
    Set GetOrCreateSheet = ws
End Function



Public Sub DeleteClientFrom_EvalData_ByName()
    Dim nm As String
    nm = Trim$(InputBox("蜑企勁縺励◆縺・茜逕ｨ閠・・豌丞錐・亥ｮ悟・荳閾ｴ・・, "EvalData 蜑企勁", ""))
    If nm = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim nameCol As Long: nameCol = 89 'CK

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, nameCol).End(xlUp).row

    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, nameCol).value) = nm Then
            ws.rows(r).Delete
            cnt = cnt + 1
        End If
    Next r

    MsgBox "EvalData: " & cnt & " 陦後ｒ蜑企勁縺励∪縺励◆縲・, vbInformation
End Sub



' 蜷悟ｧ灘酔蜷肴凾・壼炎髯､繝懊ち繝ｳ驕狗畑縺ｧ縺ｯ EvalData 縺九ｉ譛譁ｰ縺ｮID(遨ｺ縺ｧ縺ｪ縺・繧定・蜍戊｣懷勧蜿門ｾ励☆繧九◆繧√・壼ｸｸ縺ｯID蜈･蜉帙ｒ豎ゅａ縺ｪ縺・

' EvalData 蟆ら畑・壽ｰ丞錐・・K=89蛻暦ｼ我ｸ閾ｴ縺ｮ陦後ｒ縲悟挨繝悶ャ繧ｯ縺ｸ騾驕ｿ縲坂・縲悟・縺九ｉ蜑企勁縲・
Public Sub ArchiveAndDelete_EvalData_ByName()

    Dim nm As String
    nm = Trim$(InputBox("EvalData・夐驕ｿ竊貞炎髯､縺励◆縺・ｰ丞錐・亥ｮ悟・荳閾ｴ・・ & vbCrLf & "窶ｻ蜷悟ｧ灘酔蜷阪′縺・ｋ蝣ｴ蜷医・谺｡縺ｫID繧定◇縺阪∪縺・, "EvalData 蛻ｩ逕ｨ邨ゆｺ・・い繝ｼ繧ｫ繧､繝・, ""))
    If nm = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Const NAME_COL As Long = 89 'CK・域ｰ丞錐・・

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, NAME_COL).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "EvalData 縺ｫ繝・・繧ｿ陦後′縺ゅｊ縺ｾ縺帙ｓ縲・, vbExclamation
        Exit Sub
    End If
    
    
    Const ID_COL As Long = 82 'Basic.ID


'--- 蜷悟ｧ灘酔蜷阪メ繧ｧ繝・け・夊､・焚繝偵ャ繝医↑繧迂D繧定ｿｽ蜉縺ｧ閨槭￥ ---
Dim hitCount As Long: hitCount = 0
Dim r2 As Long
For r2 = lastRow To 2 Step -1
    If CStr(ws.Cells(r2, NAME_COL).value) = nm Then hitCount = hitCount + 1
Next r2

Dim pid As String: pid = ""
If hitCount >= 2 Then
    If Len(gArchiveDeleteBasicID) > 0 Then pid = gArchiveDeleteBasicID
        If pid = "" Then
    End If
    
    
    
    If pid = "" Then
        pid = Trim$(InputBox("蜷悟ｧ灘酔蜷阪′隕九▽縺九ｊ縺ｾ縺励◆縲ょ炎髯､・磯驕ｿ・峨＠縺溘＞ID繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・, "ID縺ｧ迚ｹ螳・, ""))
        If pid = "" Then
            MsgBox "ID譛ｪ蜈･蜉帙・縺溘ａ荳ｭ豁｢縺励∪縺励◆縲・, vbExclamation
            Exit Sub
        End If
    End If
End If


    

    Dim ans As VbMsgBoxResult
    ans = MsgBox("EvalData 縺ｮ豌丞錐=" & nm & " 繧偵い繝ｼ繧ｫ繧､繝悶∈騾驕ｿ縺励∝・繝・・繧ｿ縺九ｉ蜑企勁縺励∪縺吶ょｮ溯｡後＠縺ｾ縺吶°・・, _
                 vbYesNo + vbQuestion, "譛邨ら｢ｺ隱・)
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim wbArc As Workbook: Set wbArc = Workbooks.Add(xlWBATWorksheet)
    Dim wsA As Worksheet: Set wsA = wbArc.Worksheets(1)
    wsA.name = "EvalData"

    ' 繝倥ャ繝騾驕ｿ・・1:FW1 繧偵◎縺ｮ縺ｾ縺ｾ・・
    ws.Range("A1:FW1").Copy Destination:=wsA.Range("A1")

    Dim moved As Long: moved = 0
    Dim r As Long

    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, NAME_COL).value) = nm And (pid = "" Or CStr(ws.Cells(r, ID_COL).value) = pid) Then
            ' 陦碁驕ｿ・・:FW 縺ｮ陦鯉ｼ・
            Dim nextA As Long
            nextA = wsA.Cells(wsA.rows.count, 1).End(xlUp).row + 1
            ws.Range("A" & r & ":FW" & r).Copy Destination:=wsA.Range("A" & nextA)

            ' 蜈・°繧牙炎髯､
            ws.rows(r).Delete
            moved = moved + 1
        End If
    Next r

    ' 繧｢繝ｼ繧ｫ繧､繝紋ｿ晏ｭ・
    Dim arcPath As String, arcFile As String
    arcPath = ThisWorkbook.path
    If arcPath = "" Then arcPath = Environ$("TEMP")
    arcFile = arcPath & Application.PathSeparator & "EvalData_Archive_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    wbArc.SaveAs fileName:=arcFile, FileFormat:=xlOpenXMLWorkbook 'xlsx

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "螳御ｺ・ｼ哘valData 縺九ｉ " & moved & " 陦後ｒ騾驕ｿ竊貞炎髯､縺励∪縺励◆縲・ & vbCrLf & _
           "繧｢繝ｼ繧ｫ繧､繝厄ｼ・ & arcFile, vbInformation
           
           
           gArchiveDeleteBasicID = ""

End Sub




Public Sub AddHeaderArchiveDeleteButton()
    Dim f As Object: Set f = frmEval
    Dim hdr As MSForms.Frame: Set hdr = f.controls("frHeader")

    Dim btn As MSForms.CommandButton

    On Error Resume Next
    Set btn = hdr.controls("cmdArchiveDelete")
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = hdr.controls.Add("Forms.CommandButton.1", "cmdArchiveDelete", True)
        btn.caption = "邨ゆｺ・・炎髯､"
        btn.Width = 90
        btn.Height = hdr.controls("txtHdrPID").Height
        btn.Top = hdr.controls("txtHdrPID").Top
        btn.Left = 8
    End If
End Sub





Public Function GetLatestID_ForName(ByVal ws As Worksheet, ByVal nm As String, ByVal nameCol As Long, ByVal idCol As Long) As String
    Dim lastRow As Long, r As Long
    Dim idS As String
    
    lastRow = ws.Cells(ws.rows.count, nameCol).End(xlUp).row
    
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, nameCol).value) = nm Then
            idS = Trim$(CStr(ws.Cells(r, idCol).value))
            If idS <> "" Then
                GetLatestID_ForName = idS
                Exit Function
            End If
        End If
    Next r
End Function





