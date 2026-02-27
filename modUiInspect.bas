Attribute VB_Name = "modUiInspect"
' ===== modUiInspect.bas =====
Option Explicit

Public Sub DumpControlsTree(Optional ByVal uf As Object)
    ' 使い方：DumpControlsTree frmEval
    If uf Is Nothing Then
        
        Exit Sub
    End If
    
    Dim path As String: path = uf.name
    DumpChildren uf, path
End Sub

Private Sub DumpChildren(ByVal parent As Object, ByVal path As String)
    Dim i As Long, c As Object
    On Error Resume Next
    For i = 0 To parent.Controls.Count - 1
        Set c = parent.Controls(i)
        If Not c Is Nothing Then
            Dim tp$, nm$, cap$
            tp = TypeName(c): nm = GetSafeName(c): cap = GetSafeCaption(c)
            
            ' ネストを掘る（Frame / MultiPage / Page）
            If tp = "Frame" Then
                DumpChildren c, path & "." & nm
            ElseIf tp = "MultiPage" Then
                Dim p As Integer
                For p = 0 To c.Pages.Count - 1
                    DumpChildren c.Pages(p), path & "." & nm & ".Page(" & p & ")"
                Next
            ElseIf tp = "Page" Then
                DumpChildren c, path & "." & nm
            End If
        End If
        Set c = Nothing
    Next
    On Error GoTo 0
End Sub

Private Function GetSafeName(o As Object) As String
    On Error Resume Next
    GetSafeName = o.name
    If Err.Number <> 0 Then GetSafeName = "(no-name)"
    Err.Clear
End Function

Private Function GetSafeCaption(o As Object) As String
    On Error Resume Next
    GetSafeCaption = o.caption
    If Err.Number <> 0 Then GetSafeCaption = ""
    Err.Clear
End Function

' いま表示中のタブ（MultiPageの現在ページ）にある ComboBox 名を列挙
Public Sub ListCombosOnActivePage_Safe()
    Dim mp As Object, pg As Object
    Set mp = GetActiveOrFirstMultiPage(frmEval)
    If mp Is Nothing Then Debug.Print "[ERR] MultiPage not found": Exit Sub
    Set pg = mp.Pages(mp.value)
    Debug.Print "=== Combos on page:", SafeCaption(pg), "==="
    ListCombosRecursive pg
End Sub



Private Function FindFirstByType(container As Object, ByVal t As String) As Object
    Dim c As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = t Then Set FindFirstByType = c: Exit Function
        If HasControls(c) Then
            Set FindFirstByType = FindFirstByType(c, t)
            If Not FindFirstByType Is Nothing Then Exit Function
        End If
    Next
End Function

Private Sub ListCombosRecursive(parent As Object)
    Dim c As Object
    On Error Resume Next
    For Each c In parent.Controls
        If TypeName(c) = "ComboBox" Then Debug.Print c.name
        If HasControls(c) Then ListCombosRecursive c
    Next
End Sub




' 姿勢評価（表示中のページ）にある主要コントロールを一覧表示
Public Sub ListKeyCtrlsOnActivePage_Safe()
    Dim mp As Object, pg As Object
    Set mp = GetActiveOrFirstMultiPage(frmEval)
    If mp Is Nothing Then Debug.Print "[ERR] MultiPage not found": Exit Sub
    Set pg = mp.Pages(mp.value)
    

    ListByTypeRecursive pg, Array("TextBox", "CheckBox", "OptionButton", "ComboBox")
End Sub

' ??? helpers（前に貼ったものを流用）???
Private Function GetActiveOrFirstMultiPage(uf As Object) As Object
    If TypeName(uf.ActiveControl) = "MultiPage" Then
        Set GetActiveOrFirstMultiPage = uf.ActiveControl
    Else
        Set GetActiveOrFirstMultiPage = FindFirstByType(uf, "MultiPage")
    End If
End Function



Private Sub ListByTypeRecursive(parent As Object, wantTypes As Variant, Optional ByVal path As String = "")
    Dim c As Object, t$, nm$, cap$
    On Error Resume Next
    For Each c In parent.Controls
        t = TypeName(c)
        If IsWantedType(t, wantTypes) Then
            nm = SafeName(c): cap = SafeCaption(c)
            Debug.Print t, padRight(path & SafeName(parent), 30), nm, "Caption:", cap
        End If
        If HasControls(c) Then ListByTypeRecursive c, wantTypes, path & SafeName(parent) & "."
    Next
End Sub

Private Function IsWantedType(ByVal t As String, arr As Variant) As Boolean
    Dim i&
    For i = LBound(arr) To UBound(arr)
        If t = arr(i) Then IsWantedType = True: Exit Function
    Next
End Function

Private Function HasControls(obj As Object) As Boolean
    On Error Resume Next
    HasControls = (obj.Controls.Count >= 0)
End Function

Private Function SafeCaption(o As Object) As String
    On Error Resume Next
    SafeCaption = o.caption
End Function

Private Function SafeName(o As Object) As String
    On Error Resume Next
    SafeName = o.name
End Function

Private Function padRight(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then padRight = s Else padRight = s & Space$(n - Len(s))
End Function

