Attribute VB_Name = "modInterestIO"

Option Explicit

Private Const sep As String = "|"

Public Sub SaveInterestToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    SaveInterestCategory ws, r, owner, "Now", "Interest_Now", Array("テレビ・新聞", "家事", "散歩", "趣味", "人と話す", "その他")
    SaveInterestCategory ws, r, owner, "Past", "Interest_Past", Array("仕事", "家事・役割", "趣味活動", "外出・旅行", "地域活動", "その他")
    SaveInterestCategory ws, r, owner, "Want", "Interest_Want", Array("散歩・運動", "買い物", "趣味活動", "外出・旅行", "家のこと", "その他")
    SaveInterestCategory ws, r, owner, "Social", "Interest_Social", Array("買い物", "家族との時間", "友人交流", "地域活動", "外出", "その他")
End Sub

Public Sub LoadInterestFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    LoadInterestCategory ws, r, owner, "Now", "Interest_Now", Array("テレビ・新聞", "家事", "散歩", "趣味", "人と話す", "その他")
    LoadInterestCategory ws, r, owner, "Past", "Interest_Past", Array("仕事", "家事・役割", "趣味活動", "外出・旅行", "地域活動", "その他")
    LoadInterestCategory ws, r, owner, "Want", "Interest_Want", Array("散歩・運動", "買い物", "趣味活動", "外出・旅行", "家のこと", "その他")
    LoadInterestCategory ws, r, owner, "Social", "Interest_Social", Array("買い物", "家族との時間", "友人交流", "地域活動", "外出", "その他")
End Sub

Private Sub SaveInterestCategory(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, ByVal key As String, ByVal headerName As String, ByVal labels As Variant)
    Dim picks As Collection
    Set picks = New Collection

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        If GetCheckboxValue(owner, "chkInterest_" & key & "_" & CStr(i)) Then
            picks.Add CStr(labels(i))
        End If
    Next i

    Dim otherText As String
    otherText = SanitizeFreeText(GetTextboxText(owner, "txtInterest_" & key & "_Other"))
    If LenB(otherText) > 0 Then
        picks.Add "その他:" & otherText
    End If

    ws.Cells(r, EnsureHeaderCol(ws, headerName)).value = JoinCollection(picks, sep)
End Sub

Private Sub LoadInterestCategory(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, ByVal key As String, ByVal headerName As String, ByVal labels As Variant)
    Dim raw As String
    raw = Trim$(CStr(ws.Cells(r, EnsureHeaderCol(ws, headerName)).value))

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        SetCheckboxValue owner, "chkInterest_" & key & "_" & CStr(i), False
    Next i
    SetCheckboxValue owner, "chkInterest_" & key & "_Other", False
    SetTextboxText owner, "txtInterest_" & key & "_Other", ""

    If LenB(raw) = 0 Then Exit Sub

    Dim tokens As Variant
    tokens = Split(raw, sep)

    Dim token As Variant
    For Each token In tokens
        ApplyInterestToken owner, key, labels, Trim$(CStr(token))
    Next token
End Sub

Private Sub ApplyInterestToken(ByVal owner As Object, ByVal key As String, ByVal labels As Variant, ByVal token As String)
    If LenB(token) = 0 Then Exit Sub

    If Left$(token, 4) = "その他:" Then
        SetCheckboxValue owner, "chkInterest_" & key & "_Other", True
        SetTextboxText owner, "txtInterest_" & key & "_Other", Mid$(token, 5)
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        If StrComp(token, CStr(labels(i)), vbTextCompare) = 0 Then
            SetCheckboxValue owner, "chkInterest_" & key & "_" & CStr(i), True
            Exit For
        End If
    Next i
End Sub

Private Function GetCheckboxValue(ByVal owner As Object, ByVal controlName As String) As Boolean
    Dim ctl As Object
    Set ctl = FindControlByNameDeep(owner, controlName)
    If ctl Is Nothing Then Exit Function

    On Error GoTo EH
    GetCheckboxValue = CBool(ctl.value)
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetTextboxText(ByVal owner As Object, ByVal controlName As String) As String
    Dim ctl As Object
    Set ctl = FindControlByNameDeep(owner, controlName)
    If ctl Is Nothing Then Exit Function

    On Error GoTo EH
    GetTextboxText = Trim$(CStr(ctl.text))
    Exit Function
EH:
    Err.Clear
End Function

Private Sub SetCheckboxValue(ByVal owner As Object, ByVal controlName As String, ByVal value As Boolean)
    Dim ctl As Object
    Set ctl = FindControlByNameDeep(owner, controlName)
    If ctl Is Nothing Then Exit Sub

    On Error Resume Next
    ctl.value = value
    Err.Clear
End Sub

Private Sub SetTextboxText(ByVal owner As Object, ByVal controlName As String, ByVal value As String)
    Dim ctl As Object
    Set ctl = FindControlByNameDeep(owner, controlName)
    If ctl Is Nothing Then Exit Sub

    On Error Resume Next
    ctl.text = value
    Err.Clear
End Sub

Private Function FindControlByNameDeep(ByVal owner As Object, ByVal controlName As String) As Object
    If owner Is Nothing Then Exit Function

    Dim direct As Object
    Set direct = TryGetControlFromContainer(owner, controlName)
    If Not direct Is Nothing Then
        Set FindControlByNameDeep = direct
        Exit Function
    End If

    Set FindControlByNameDeep = FindControlInChildren(owner, controlName)
End Function

Private Function FindControlInChildren(ByVal container As Object, ByVal controlName As String) As Object
    Dim children As Object
    Set children = TryGetObjectMember(container, "Controls")
    If children Is Nothing Then Exit Function

    Dim child As Object
    For Each child In children
        If StrComp(GetObjectName(child), controlName, vbTextCompare) = 0 Then
            Set FindControlInChildren = child
            Exit Function
        End If

        Dim nested As Object
        Set nested = FindControlInChildren(child, controlName)
        If Not nested Is Nothing Then
            Set FindControlInChildren = nested
            Exit Function
        End If
    Next child
End Function

Private Function TryGetControlFromContainer(ByVal container As Object, ByVal controlName As String) As Object
    On Error GoTo EH
    Set TryGetControlFromContainer = container.Controls(controlName)
    Exit Function
EH:
    Err.Clear
End Function

Private Function TryGetObjectMember(ByVal target As Object, ByVal memberName As String) As Object
    On Error GoTo EH
    If StrComp(memberName, "Controls", vbTextCompare) = 0 Then
        Set TryGetObjectMember = target.Controls
    End If
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetObjectName(ByVal target As Object) As String
    On Error GoTo EH
    GetObjectName = CStr(CallByName(target, "Name", VbGet))
    Exit Function
EH:
    Err.Clear
End Function

Private Function SanitizeFreeText(ByVal src As String) As String
    SanitizeFreeText = Replace$(src, sep, "｜")
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To values.count
        If LenB(JoinCollection) > 0 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(values(i))
    Next i
End Function

