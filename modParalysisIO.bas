Attribute VB_Name = "modParalysisIO"

Option Explicit

Private Function ColOf(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If Not f Is Nothing Then ColOf = f.Column
End Function

' ---- まとめ取得（キー付 Collection）----
Public Function GetParalysisState(ByVal owner As frmEval) As Collection
    Dim col As New Collection
    On Error Resume Next
    col.Add GetCtlText(owner, "cboParalysisSide"), "麻痺側"
    col.Add GetCtlText(owner, "cboParalysisType"), "麻痺の種類"
    col.Add GetCtlText(owner, "cboBRS_Upper"), "BRS_上肢"
    col.Add GetCtlText(owner, "cboBRS_Hand"), "BRS_手指"
    col.Add GetCtlText(owner, "cboBRS_Lower"), "BRS_下肢"
    col.Add GetCtlCheck(owner, "chkSynergy"), "共同運動"
    col.Add GetCtlCheck(owner, "chkAssociatedRxn"), "連合反応"
    col.Add GetCtlText(owner, "txtParalysisMemo"), "麻痺_備考"
    Set GetParalysisState = col
End Function

' ---- 保存：見出しが無ければ自動で作成 ----
Public Sub SaveParalysisToSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim s As Collection: Set s = GetParalysisState(owner)
    Dim look As Object: Set look = BuildHeaderLookup(ws)

    Dim k As Variant, c As Long
    For Each k In Array("麻痺側", "麻痺の種類", "BRS_上肢", "BRS_手指", "BRS_下肢", "共同運動", "連合反応", "麻痺_備考")
        c = ResolveColOrCreate(ws, look, CStr(k))   ' ← 見出し自動生成
        ws.Cells(rowNum, c).value = s(CStr(k))
    Next k
End Sub

' ---- 読込：列がある場合のみ読む（安全） ----
Public Sub LoadParalysisFromSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim look As Object: Set look = BuildHeaderLookup(ws)
    Dim c As Long

    Dim v As Variant

c = ResolveColumn(look, "麻痺側"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboParalysisSide", v
c = ResolveColumn(look, "麻痺の種類"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboParalysisType", v
c = ResolveColumn(look, "BRS_上肢"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Upper", v
c = ResolveColumn(look, "BRS_手指"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Hand", v
c = ResolveColumn(look, "BRS_下肢"): If c > 0 Then v = ws.Cells(rowNum, c).value: SetComboSafe owner, "cboBRS_Lower", v

    c = ResolveColumn(look, "共同運動"):        If c > 0 Then FindCtlDeep(owner, "chkSynergy").value = (ws.Cells(rowNum, c).value = "有")
    c = ResolveColumn(look, "連合反応"):        If c > 0 Then FindCtlDeep(owner, "chkAssociatedRxn").value = (ws.Cells(rowNum, c).value = "有")
    c = ResolveColumn(look, "麻痺_備考"):       If c > 0 Then FindCtlDeep(owner, "txtParalysisMemo").value = ws.Cells(rowNum, c).value
End Sub


' 値がコンボのリストにある時だけ選択する（無ければ未選択）
Private Sub SetComboSafe(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cB As MSForms.ComboBox
    Dim s As String, i As Long, hit As Long

    s = CStr(v)
    Set cB = FindCtlDeep(owner, ctlName)
    If cB Is Nothing Then Exit Sub

    hit = -1
    For i = 0 To cB.ListCount - 1
        If CStr(cB.List(i)) = s Then hit = i: Exit For
    Next

    If hit >= 0 Then
        cB.ListIndex = hit              ' ← 安全に選択
    Else
        cB.ListIndex = -1               ' ← 見つからなければクリア（空）
        ' 必要ならここで：cb.AddItem s : cb.Value = s   ' 自動で項目を追加して選択
    End If
End Sub

