Attribute VB_Name = "modROMIO"
Option Explicit

' ===== 入口：ROMを保存（見出しは無ければ作る） =====
Public Sub SaveROMToSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim look As Object
Set look = BuildHeaderLookup(ws)

    ' 上肢
    SaveROMblock ws, rowNum, owner, look, "Upper", "Shoulder", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")
    SaveROMblock ws, rowNum, owner, look, "Upper", "Elbow", Array("Flex", "Ext")
    SaveROMblock ws, rowNum, owner, look, "Upper", "Forearm", Array("Sup", "Pro")
    SaveROMblock ws, rowNum, owner, look, "Upper", "Wrist", Array("Dorsi", "Palmar", "Radial", "Ulnar")
    SaveROMMemo ws, rowNum, owner, look, "Upper"
    
    

    ' 下肢
    SaveROMblock ws, rowNum, owner, look, "Lower", "Hip", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")
    SaveROMblock ws, rowNum, owner, look, "Lower", "Knee", Array("Flex", "Ext")
    SaveROMblock ws, rowNum, owner, look, "Lower", "Ankle", Array("Dorsi", "Plantar", "Inv", "Ev")
    SaveROMMemo ws, rowNum, owner, look, "Lower"
    SaveROMTrunk ws, rowNum, owner, look
End Sub

' ===== 入口：ROMを読込（見出しがある列だけ読む） =====
Public Sub LoadROMFromSheet(ws As Worksheet, rowNum As Long, owner As frmEval)
    Dim look As Object: Set look = BuildHeaderLookup(ws)


#If APP_DEBUG Then
    Debug.Print "[ROM][LOAD]", _
                "row=" & rowNum, _
                "| IO_ROM.len=" & Len(ReadStr_Compat("IO_ROM", rowNum, ws))
#End If




    ' 上肢
    LoadROMblock ws, rowNum, owner, look, "Upper", "Shoulder", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")
    LoadROMblock ws, rowNum, owner, look, "Upper", "Elbow", Array("Flex", "Ext")
    LoadROMblock ws, rowNum, owner, look, "Upper", "Forearm", Array("Sup", "Pro")
    LoadROMblock ws, rowNum, owner, look, "Upper", "Wrist", Array("Dorsi", "Palmar", "Radial", "Ulnar")
    LoadROMMemo ws, rowNum, owner, look, "Upper"

    ' 下肢
    LoadROMblock ws, rowNum, owner, look, "Lower", "Hip", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")
    LoadROMblock ws, rowNum, owner, look, "Lower", "Knee", Array("Flex", "Ext")
    LoadROMblock ws, rowNum, owner, look, "Lower", "Ankle", Array("Dorsi", "Plantar", "Inv", "Ev")
    LoadROMMemo ws, rowNum, owner, look, "Lower"
    LoadROMTrunk ws, rowNum, owner, look
End Sub


' ==== 内部：1ブロック保存/読込・備考の保存/読込 ====

Private Sub SaveROMblock(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                         layer As String, joint As String, motions As Variant)
    Dim m As Variant, side As Variant
    For Each m In motions
        For Each side In Array("R", "L")
            Dim hdr$, ctl$, col&, v$
            hdr = "ROM_" & layer & "_" & joint & "_" & CStr(m) & "_" & CStr(side)
            ctl = "txtROM_" & layer & "_" & joint & "_" & CStr(m) & "_" & CStr(side)
            v = GetCtlText(owner, ctl)                              ' ModUtil
            col = ResolveColOrCreate(ws, look, hdr)                 ' ★無ければ見出し作成
            ws.Cells(rowNum, col).value = v
        Next side
    Next m
End Sub

Private Sub LoadROMblock(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                         layer As String, joint As String, motions As Variant)

    Debug.Print "[ROM][BLOCK]", _
                "layer=" & layer, _
                "| joint=" & joint, _
                "| keySample=" & "ROM_" & layer & "_" & joint & "_" & CStr(motions(LBound(motions))) & "_R", _
                "| colSample=" & HeaderCol_Compat("ROM_" & layer & "_" & joint & "_" & CStr(motions(LBound(motions))) & "_R", ws)

    Dim m As Variant, side As Variant
    For Each m In motions
        For Each side In Array("R", "L")
            Dim hdr$, ctl$, col&
            hdr = "ROM_" & layer & "_" & joint & "_" & CStr(m) & "_" & CStr(side)
            ctl = "txtROM_" & layer & "_" & joint & "_" & CStr(m) & "_" & CStr(side)
            col = ResolveColumn(look, hdr)                          ' 無ければスキップ
            Dim v As String: v = ReadStr_Compat(hdr, rowNum, ws)
If Len(v) > 0 Then FindCtlDeep(owner, ctl).text = v

        Next side
    Next m
End Sub


Private Sub SaveROMMemo(ws As Worksheet, rowNum As Long, owner As Object, look As Object, layer As String)
    Dim hdr$, ctl$, col&, v$
    hdr = "ROM_" & layer & "_Memo"
    ctl = "txtROM_" & layer & "_Memo"
    v = GetCtlText(owner, ctl)
    col = ResolveColOrCreate(ws, look, hdr)
    ws.Cells(rowNum, col).value = v
End Sub

Private Sub LoadROMMemo(ws As Worksheet, rowNum As Long, owner As Object, look As Object, layer As String)
    Dim hdr$, ctl$, col&
    hdr = "ROM_" & layer & "_Memo"
    ctl = "txtROM_" & layer & "_Memo"
    col = ResolveColumn(look, hdr)
    If col > 0 Then FindCtlDeep(owner, ctl).text = ws.Cells(rowNum, col).value
End Sub

Private Sub SaveROMTrunk(ws As Worksheet, rowNum As Long, owner As Object, look As Object)
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Flex", "txtROM_Trunk_Neck_Flex"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Ext", "txtROM_Trunk_Neck_Ext"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Rot_R", "txtROM_Trunk_Neck_Rot_R"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Rot_L", "txtROM_Trunk_Neck_Rot_L"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_LatFlex_R", "txtROM_Trunk_Neck_LatFlex_R"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_LatFlex_L", "txtROM_Trunk_Neck_LatFlex_L"
    
    
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Flex", "txtROM_Trunk_Trunk_Flex"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Ext", "txtROM_Trunk_Trunk_Ext"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Rot_R", "txtROM_Trunk_Trunk_Rot_R"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Rot_L", "txtROM_Trunk_Trunk_Rot_L"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_LatFlex_R", "txtROM_Trunk_Trunk_LatFlex_R"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_LatFlex_L", "txtROM_Trunk_Trunk_LatFlex_L"
    SaveROMTrunkValue ws, rowNum, owner, look, "Thorax_Expansion", "txtROM_Trunk_Thorax_ChestDiff"
    SaveROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Memo", "txtROM_Trunk_Memo"
End Sub

Private Sub LoadROMTrunk(ws As Worksheet, rowNum As Long, owner As Object, look As Object)
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Flex", "txtROM_Trunk_Neck_Flex"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Ext", "txtROM_Trunk_Neck_Ext"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Rot_R", "txtROM_Trunk_Neck_Rot_R"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_Rot_L", "txtROM_Trunk_Neck_Rot_L"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_LatFlex_R", "txtROM_Trunk_Neck_LatFlex_R"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Neck_LatFlex_L", "txtROM_Trunk_Neck_LatFlex_L"


    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Flex", "txtROM_Trunk_Trunk_Flex"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Ext", "txtROM_Trunk_Trunk_Ext"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Rot_R", "txtROM_Trunk_Trunk_Rot_R"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Rot_L", "txtROM_Trunk_Trunk_Rot_L"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_LatFlex_R", "txtROM_Trunk_Trunk_LatFlex_R"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_LatFlex_L", "txtROM_Trunk_Trunk_LatFlex_L"
    LoadROMTrunkValue ws, rowNum, owner, look, "Thorax_Expansion", "txtROM_Trunk_Thorax_ChestDiff"
    LoadROMTrunkValue ws, rowNum, owner, look, "ROM_Trunk_Memo", "txtROM_Trunk_Memo"
End Sub

Private Sub SaveROMTrunkValue(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                              ByVal header As String, ByVal ctlName As String)
    Dim col As Long
    Dim ctl As Object

    Set ctl = FindCtlDeep(owner, ctlName)
    If ctl Is Nothing Then Exit Sub
    
    col = ResolveColOrCreate(ws, look, header, LegacyTrunkHeaderName(header))
    If col = 0 Then Exit Sub

    On Error Resume Next
    ws.Cells(rowNum, col).Value2 = CStr(ctl.value)
    On Error GoTo 0
End Sub

Private Sub LoadROMTrunkValue(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                              ByVal header As String, ByVal ctlName As String)
    Dim col As Long
    Dim ctl As Object
    Dim v As Variant
    
    col = ResolveColumn(look, header)
    If col = 0 Then col = ResolveColumn(look, LegacyTrunkHeaderName(header))
    If col = 0 Then Exit Sub

    Set ctl = FindCtlDeep(owner, ctlName)
    If ctl Is Nothing Then Exit Sub
    
    v = ws.Cells(rowNum, col).Value2
    If IsError(v) Or IsNull(v) Then v = vbNullString
    
    On Error Resume Next
    ctl.text = CStr(v)
    On Error GoTo 0
End Sub

Private Function LegacyTrunkHeaderName(ByVal header As String) As String
    Select Case header
        Case "ROM_Trunk_Flex":      LegacyTrunkHeaderName = "ROM_Trunk_Trunk_Flex"
        Case "ROM_Trunk_Ext":       LegacyTrunkHeaderName = "ROM_Trunk_Trunk_Ext"
        Case "ROM_Trunk_Rot_R":     LegacyTrunkHeaderName = "ROM_Trunk_Trunk_Rot_R"
        Case "ROM_Trunk_Rot_L":     LegacyTrunkHeaderName = "ROM_Trunk_Trunk_Rot_L"
        Case "ROM_Trunk_LatFlex_R": LegacyTrunkHeaderName = "ROM_Trunk_Trunk_LatFlex_R"
        Case "ROM_Trunk_LatFlex_L": LegacyTrunkHeaderName = "ROM_Trunk_Trunk_LatFlex_L"
        Case "ROM_Trunk_Memo":      LegacyTrunkHeaderName = "ROM_Trunk_Trunk_Memo"
    End Select
End Function



'=== 基本情報の候補名を一覧表示（イミディエイトに出力） ===
Public Sub FindNamesForBasics()
    On Error Resume Next
    'Load frmEval: frmEval.Show vbModeless
    
    Dim keys, k, c As Object
    ' ← 欲しい項目のキーワード。必要なら増減してOK
    keys = Array("Date", "Age", "Sex", "Eval", "Evaluator", "Name", _
                 "Onset", "CareLevel", "Dementia", "ADL", "Needs", "NeedsPt", "NeedsFam", _
                 "Diagnosis", "Living", "生活", "主診断")
    
    Debug.Print "----- candidates -----"
    For Each k In keys
        For Each c In frmEval.Controls
            If InStr(1, c.name, CStr(k), vbTextCompare) > 0 Then
                Debug.Print k, typeName(c), c.name
            End If
        Next c
    Next k
    Debug.Print "----------------------"
End Sub




'=== Local: 1行目に見出しが無ければ作って列番号を返す ===
Private Function HeaderColEnsure(ws As Worksheet, ByVal header As String) As Long
    Dim m As Variant, lastCol As Long
    m = Application.Match(header, ws.rows(1), 0)
    If IsError(m) Then
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = header
        HeaderColEnsure = lastCol + 1
    Else
        HeaderColEnsure = CLng(m)
    End If
End Function



Public Sub Debug_ShowROM_Abd_R()
    Dim ctl As Object
    Set ctl = FindCtlDeep(frmEval, "txtROM_Upper_Shoulder_Abd_R")

    If ctl Is Nothing Then
        Debug.Print "[ROMCTL] txtROM_Upper_Shoulder_Abd_R not found"
    Else
        Debug.Print "[ROMCTL] Name=" & ctl.name & " Text='" & ctl.text & "'"
    End If
End Sub

