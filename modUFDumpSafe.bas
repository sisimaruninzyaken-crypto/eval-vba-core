Attribute VB_Name = "modUFDumpSafe"
Option Explicit

'========================================================
' Safe Dump: 重複巡回を防ぎ、MultiPage/Page を安全に扱う版
' - Visited(Dictionary)で多重訪問防止
' - MultiPage.Pages は専用ループ（PageのLeft/Width等は触らない）
' - Page は「直下Controlsを列挙」だけ。再帰は Frame のみ許可
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

    ' parent自身をVisited（同一インスタンスの再訪防止）
    Dim keyP As String
    keyP = ObjKey(parent)
    If keyP <> "" Then
        If visited.exists(keyP) Then Exit Sub
        visited.Add keyP, True
    End If

    ' parentの直下だけ列挙（Parent一致でフィルタ）
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If (c.parent Is parent) Then
            DumpOne ws, c, depth + 1

            If TypeName(c) = "Frame" Then
                DumpControl_Safe ws, c, depth + 1, visited

            ElseIf TypeName(c) = "MultiPage" Then
                DumpMultiPage_Safe ws, c, depth + 1, visited

            Else
                ' それ以外は再帰しない
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
    ws.WriteLine Indent(depth) & "(Pages.Count=" & mp.Pages.Count & ")"

    For i = 0 To mp.Pages.Count - 1
        Dim pg As MSForms.Page
        Set pg = mp.Pages(i)

        ' Pageは位置情報に触れない（438回避）。Name/Captionだけ
        ws.WriteLine Indent(depth) & "[Page " & i & "] Name=" & SafeStr(pg, "Name") & " Caption=" & SafeStr(pg, "Caption")

        ' Page直下Controlsを列挙（Page自体は再帰対象にしない）
        DumpPageChildren_Safe ws, pg, depth + 1, visited
    Next i

    Exit Sub

ErrH:
#If APP_DEBUG Then
    Debug.Print "[DumpMultiPage_Safe] ERR " & Err.Number & " " & Err.Description
#End If
End Sub

Private Sub DumpPageChildren_Safe(ByVal ws As Object, ByVal pg As MSForms.Page, ByVal depth As Long, ByVal visited As Object)
    On Error GoTo ErrH

    Dim c As MSForms.Control
    For Each c In pg.Controls
        If (c.parent Is pg) Then
            DumpOne ws, c, depth

            ' Page配下は Frame だけ再帰OK（必要ならMultiPageもOK）
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

    ' 座標はControlのみ（Pageは触らない）
    Dim l As Variant, t As Variant, w As Variant, h As Variant
    l = SafeNum(c, "Left")
    t = SafeNum(c, "Top")
    w = SafeNum(c, "Width")
    h = SafeNum(c, "Height")

    If Not IsEmpty(l) Then s = s & " L=" & Format$(l, "0.00")
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

