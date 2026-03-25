Attribute VB_Name = "modUFDumpSafe"
Option Explicit

'========================================================
' Safe Dump: 驥崎､・ｷ｡蝗槭ｒ髦ｲ縺弱｀ultiPage/Page 繧貞ｮ牙・縺ｫ謇ｱ縺・沿
' - Visited(Dictionary)縺ｧ螟夐㍾險ｪ蝠城亟豁｢
' - MultiPage.Pages 縺ｯ蟆ら畑繝ｫ繝ｼ繝暦ｼ・age縺ｮLeft/Width遲峨・隗ｦ繧峨↑縺・ｼ・
' - Page 縺ｯ縲檎峩荳気ontrols繧貞・謖吶阪□縺代ょ・蟶ｰ縺ｯ Frame 縺ｮ縺ｿ險ｱ蜿ｯ
'========================================================

Public Sub DumpFrmEvalTree_ToFile_Safe(Optional ByVal outPath As String = "")
    On Error GoTo ErrH

    Dim f As Object, ws As Object
    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")

    If outPath = "" Then
        outPath = Environ$("TEMP") & "\frmEval_tree_SAFE_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    End If

    Set f = CreateObject("Scripting.FileSystemObject")
    Set ws = f.CreateTextFile(outPath, True, True) 'Unicode

    ws.WriteLine "[SAFE-DUMP] " & Now
    ws.WriteLine "Root=frmEval"
    ws.WriteLine String(60, "-")

    DumpControl_Safe ws, frmEval, 0, visited

    ws.WriteLine String(60, "-")
    ws.WriteLine "[OUT] " & outPath
    ws.Close

#If APP_DEBUG Then
    Debug.Print "[DumpFrmEvalTree_ToFile_Safe] OUT=" & outPath
#End If
    Exit Sub

ErrH:
#If APP_DEBUG Then
    Debug.Print "[DumpFrmEvalTree_ToFile_Safe] ERR " & Err.Number & " " & Err.Description
#End If
    On Error Resume Next
    If Not ws Is Nothing Then ws.Close
End Sub

Private Sub DumpControl_Safe(ByVal ws As Object, ByVal parent As Object, ByVal depth As Long, ByVal visited As Object)
    On Error GoTo ErrH

    ' parent閾ｪ霄ｫ繧歎isited・亥酔荳繧､繝ｳ繧ｹ繧ｿ繝ｳ繧ｹ縺ｮ蜀崎ｨｪ髦ｲ豁｢・・
    Dim keyP As String
    keyP = ObjKey(parent)
    If keyP <> "" Then
        If visited.exists(keyP) Then Exit Sub
        visited.Add keyP, True
    End If

    ' parent縺ｮ逶ｴ荳九□縺大・謖呻ｼ・arent荳閾ｴ縺ｧ繝輔ぅ繝ｫ繧ｿ・・
    Dim c As MSForms.Control
    For Each c In parent.controls
        If (c.parent Is parent) Then
            DumpOne ws, c, depth + 1

            If TypeName(c) = "Frame" Then
                DumpControl_Safe ws, c, depth + 1, visited

            ElseIf TypeName(c) = "MultiPage" Then
                DumpMultiPage_Safe ws, c, depth + 1, visited

            Else
                ' 縺昴ｌ莉･螟悶・蜀榊ｸｰ縺励↑縺・
            End If
        End If
    Next c

    Exit Sub

ErrH:
#If APP_DEBUG Then
    Debug.Print "[DumpControl_Safe] ERR " & Err.Number & " " & Err.Description & " Parent=" & TypeName(parent)
#End If
End Sub

Private Sub DumpMultiPage_Safe(ByVal ws As Object, ByVal mp As MSForms.MultiPage, ByVal depth As Long, ByVal visited As Object)
    On Error GoTo ErrH

    Dim i As Long
    ws.WriteLine Indent(depth) & "(Pages.Count=" & mp.Pages.count & ")"

    For i = 0 To mp.Pages.count - 1
        Dim pg As MSForms.page
        Set pg = mp.Pages(i)

        ' Page縺ｯ菴咲ｽｮ諠・ｱ縺ｫ隗ｦ繧後↑縺・ｼ・38蝗樣∩・峨・ame/Caption縺縺・
        ws.WriteLine Indent(depth) & "[Page " & i & "] Name=" & SafeStr(pg, "Name") & " Caption=" & SafeStr(pg, "Caption")

        ' Page逶ｴ荳気ontrols繧貞・謖呻ｼ・age閾ｪ菴薙・蜀榊ｸｰ蟇ｾ雎｡縺ｫ縺励↑縺・ｼ・
        DumpPageChildren_Safe ws, pg, depth + 1, visited
    Next i

    Exit Sub

ErrH:
#If APP_DEBUG Then
    Debug.Print "[DumpMultiPage_Safe] ERR " & Err.Number & " " & Err.Description
#End If
End Sub

Private Sub DumpPageChildren_Safe(ByVal ws As Object, ByVal pg As MSForms.page, ByVal depth As Long, ByVal visited As Object)
    On Error GoTo ErrH

    Dim c As MSForms.Control
    For Each c In pg.controls
        If (c.parent Is pg) Then
            DumpOne ws, c, depth

            ' Page驟堺ｸ九・ Frame 縺縺大・蟶ｰOK・亥ｿ・ｦ√↑繧窺ultiPage繧０K・・
            If TypeName(c) = "Frame" Then
                DumpControl_Safe ws, c, depth, visited
            ElseIf TypeName(c) = "MultiPage" Then
                DumpMultiPage_Safe ws, c, depth, visited
            End If
        End If
    Next c

    Exit Sub

ErrH:
#If APP_DEBUG Then
    Debug.Print "[DumpPageChildren_Safe] ERR " & Err.Number & " " & Err.Description
#End If
End Sub

Private Sub DumpOne(ByVal ws As Object, ByVal c As MSForms.Control, ByVal depth As Long)
    On Error Resume Next

    Dim s As String
    s = Indent(depth) & TypeName(c) & " " & SafeStr(c, "Name")

    ' 蠎ｧ讓吶・Control縺ｮ縺ｿ・・age縺ｯ隗ｦ繧峨↑縺・ｼ・
    Dim L As Variant, t As Variant, w As Variant, h As Variant
    L = SafeNum(c, "Left")
    t = SafeNum(c, "Top")
    w = SafeNum(c, "Width")
    h = SafeNum(c, "Height")

    If Not IsEmpty(L) Then s = s & " L=" & Format$(L, "0.00")
    If Not IsEmpty(t) Then s = s & " T=" & Format$(t, "0.00")
    If Not IsEmpty(w) Then s = s & " W=" & Format$(w, "0.00")
    If Not IsEmpty(h) Then s = s & " H=" & Format$(h, "0.00")

    ws.WriteLine s
End Sub

Private Function Indent(ByVal n As Long) As String
    Indent = String$(n * 2, " ")
End Function

Private Function ObjKey(ByVal o As Object) As String
    On Error Resume Next
    ObjKey = TypeName(o) & ":" & SafeStr(o, "Name")
End Function

Private Function SafeStr(ByVal o As Object, ByVal propName As String) As String
    On Error GoTo EH
    Dim v As Variant
    v = CallByName(o, propName, VbGet)
    SafeStr = CStr(v)
    Exit Function
EH:
    SafeStr = ""
End Function

Private Function SafeNum(ByVal o As Object, ByVal propName As String) As Variant
    On Error GoTo EH
    Dim v As Variant
    v = CallByName(o, propName, VbGet)
    If IsNumeric(v) Then
        SafeNum = CSng(v)
    Else
        SafeNum = Empty
    End If
    Exit Function
EH:
    SafeNum = Empty
End Function



Public Sub DumpFrmEvalSafe()
    DumpFrmEvalTree_ToFile_Safe Environ$("TEMP") & "\frmEval_tree_SAFE_LAST.txt"
End Sub

