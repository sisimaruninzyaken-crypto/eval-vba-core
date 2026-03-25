Attribute VB_Name = "modParalysisIO"

Option Explicit

Private Function ColOf(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If Not f Is Nothing Then ColOf = f.Column
End Function

' ---- 縺ｾ縺ｨ繧∝叙蠕暦ｼ医く繝ｼ莉・Collection・・---
Public Function GetParalysisState(ByVal owner As frmEval) As Collection
    Dim col As New Collection
    On Error Resume Next
    col.Add GetCtlText(owner, "cboParalysisSide"), "鮗ｻ逞ｺ蛛ｴ"
    col.Add GetCtlText(owner, "cboParalysisType"), "鮗ｻ逞ｺ縺ｮ遞ｮ鬘・
    col.Add GetCtlText(owner, "cboBRS_Upper"), "BRS_荳願い"
    col.Add GetCtlText(owner, "cboBRS_Hand"), "BRS_謇区欠"
    col.Add GetCtlText(owner, "cboBRS_Lower"), "BRS_荳玖い"
    col.Add GetCtlCheck(owner, "chkSynergy"), "蜈ｱ蜷碁°蜍・
    col.Add GetCtlCheck(owner, "chkAssociatedRxn"), "騾｣蜷亥渚蠢・
    col.Add GetCtlText(owner, "txtParalysisMemo"), "鮗ｻ逞ｺ_蛯呵・
    Set GetParalysisState = col
End Function

' ---- 菫晏ｭ假ｼ夊ｦ句・縺励′辟｡縺代ｌ縺ｰ閾ｪ蜍輔〒菴懈・ ----
Public Sub SaveParalysisToSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim s As Collection: Set s = GetParalysisState(owner)
    Dim look As Object: Set look = BuildHeaderLookup(ws)

    Dim k As Variant, c As Long
    For Each k In Array("鮗ｻ逞ｺ蛛ｴ", "鮗ｻ逞ｺ縺ｮ遞ｮ鬘・, "BRS_荳願い", "BRS_謇区欠", "BRS_荳玖い", "蜈ｱ蜷碁°蜍・, "騾｣蜷亥渚蠢・, "鮗ｻ逞ｺ_蛯呵・)
        c = ResolveColOrCreate(ws, look, CStr(k))   ' 竊・隕句・縺苓・蜍慕函謌・
        ws.Cells(rowNum, c).value = s(CStr(k))
    Next k
End Sub

' ---- 隱ｭ霎ｼ・壼・縺後≠繧句ｴ蜷医・縺ｿ隱ｭ繧・亥ｮ牙・・・----
Public Sub LoadParalysisFromSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim look As Object: Set look = BuildHeaderLookup(ws)
    Dim c As Long

    Dim v As Variant

c = ResolveColumn(look, "鮗ｻ逞ｺ蛛ｴ"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboParalysisSide", v
c = ResolveColumn(look, "鮗ｻ逞ｺ縺ｮ遞ｮ鬘・): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboParalysisType", v
c = ResolveColumn(look, "BRS_荳願い"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Upper", v
c = ResolveColumn(look, "BRS_謇区欠"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Hand", v
c = ResolveColumn(look, "BRS_荳玖い"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Lower", v

    c = ResolveColumn(look, "蜈ｱ蜷碁°蜍・):        If c > 0 Then FindCtlDeep(owner, "chkSynergy").value = (ws.Cells(rowNum, c).value = "譛・)
    c = ResolveColumn(look, "騾｣蜷亥渚蠢・):        If c > 0 Then FindCtlDeep(owner, "chkAssociatedRxn").value = (ws.Cells(rowNum, c).value = "譛・)
    c = ResolveColumn(look, "鮗ｻ逞ｺ_蛯呵・):       If c > 0 Then FindCtlDeep(owner, "txtParalysisMemo").value = ws.Cells(rowNum, c).value
End Sub


' 蛟､縺後さ繝ｳ繝懊・繝ｪ繧ｹ繝医↓縺ゅｋ譎ゅ□縺鷹∈謚槭☆繧具ｼ育┌縺代ｌ縺ｰ譛ｪ驕ｸ謚橸ｼ・
Private Sub SetComboSafe(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cb As MSForms.ComboBox
    Dim s As String, i As Long, hit As Long

    s = CStr(v)
    Set cb = FindCtlDeep(owner, ctlName)
    If cb Is Nothing Then Exit Sub

    hit = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i)) = s Then hit = i: Exit For
    Next

    If hit >= 0 Then
        cb.ListIndex = hit              ' 竊・螳牙・縺ｫ驕ｸ謚・
    Else
        cb.ListIndex = -1               ' 竊・隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ繧ｯ繝ｪ繧｢・育ｩｺ・・
        ' 蠢・ｦ√↑繧峨％縺薙〒・喞b.AddItem s : cb.Value = s   ' 閾ｪ蜍輔〒鬆・岼繧定ｿｽ蜉縺励※驕ｸ謚・
    End If
End Sub

