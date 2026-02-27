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
End Sub


' ==== 内部：1ブロック保存/読込・備考の保存/読込 ====

Private Sub SaveROMblock(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                         layer As String, joint As String, motions As Variant)
    Dim M As Variant, side As Variant
    For Each M In motions
        For Each side In Array("R", "L")
            Dim hdr$, ctl$, col&, v$
            hdr = "ROM_" & layer & "_" & joint & "_" & CStr(M) & "_" & CStr(side)
            ctl = "txtROM_" & layer & "_" & joint & "_" & CStr(M) & "_" & CStr(side)
            v = GetCtlText(owner, ctl)                              ' ModUtil
            col = ResolveColOrCreate(ws, look, hdr)                 ' ★無ければ見出し作成
            ws.Cells(rowNum, col).value = v
        Next side
    Next M
End Sub

Private Sub LoadROMblock(ws As Worksheet, rowNum As Long, owner As Object, look As Object, _
                         layer As String, joint As String, motions As Variant)

    Debug.Print "[ROM][BLOCK]", _
                "layer=" & layer, _
                "| joint=" & joint, _
                "| keySample=" & "ROM_" & layer & "_" & joint & "_" & CStr(motions(LBound(motions))) & "_R", _
                "| colSample=" & HeaderCol_Compat("ROM_" & layer & "_" & joint & "_" & CStr(motions(LBound(motions))) & "_R", ws)

    Dim M As Variant, side As Variant
    For Each M In motions
        For Each side In Array("R", "L")
            Dim hdr$, ctl$, col&
            hdr = "ROM_" & layer & "_" & joint & "_" & CStr(M) & "_" & CStr(side)
            ctl = "txtROM_" & layer & "_" & joint & "_" & CStr(M) & "_" & CStr(side)
            col = ResolveColumn(look, hdr)                          ' 無ければスキップ
            Dim v As String: v = ReadStr_Compat(hdr, rowNum, ws)
If Len(v) > 0 Then FindCtlDeep(owner, ctl).Text = v

        Next side
    Next M
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
    If col > 0 Then FindCtlDeep(owner, ctl).Text = ws.Cells(rowNum, col).value
End Sub


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
                Debug.Print k, TypeName(c), c.name
            End If
        Next c
    Next k
    Debug.Print "----------------------"
End Sub




'=== Local: 1行目に見出しが無ければ作って列番号を返す ===
Private Function HeaderColEnsure(ws As Worksheet, ByVal header As String) As Long
    Dim M As Variant, lastCol As Long
    M = Application.Match(header, ws.rows(1), 0)
    If IsError(M) Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = header
        HeaderColEnsure = lastCol + 1
    Else
        HeaderColEnsure = CLng(M)
    End If
End Function



Public Sub Debug_ShowROM_Abd_R()
    Dim ctl As Object
    Set ctl = FindCtlDeep(frmEval, "txtROM_Upper_Shoulder_Abd_R")

    If ctl Is Nothing Then
        Debug.Print "[ROMCTL] txtROM_Upper_Shoulder_Abd_R not found"
    Else
        Debug.Print "[ROMCTL] Name=" & ctl.name & " Text='" & ctl.Text & "'"
    End If
End Sub

