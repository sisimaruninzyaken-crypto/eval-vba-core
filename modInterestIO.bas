Attribute VB_Name = "modInterestIO"

Option Explicit

Private Const SEP As String = "|"
Private Const SEP_SAFE As String = "｜"
Public Const INTEREST_OTHER_PREFIX As String = "その他:"

Public Sub SaveInterestToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim key As Variant
    For Each key In InterestCategoryKeys()
        SaveInterestCategory ws, r, owner, CStr(key)
    Next key
End Sub

Public Sub LoadInterestFromSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim key As Variant
    For Each key In InterestCategoryKeys()
        LoadInterestCategory ws, r, owner, CStr(key)
    Next key
End Sub

Public Function InterestCategoryKeys() As Variant
    InterestCategoryKeys = Array("Now", "Past", "Want", "Social")
End Function

Public Function InterestHeaderName(ByVal key As String) As String
    InterestHeaderName = "Interest_" & key
End Function

Public Function InterestCategoryTitle(ByVal key As String) As String
    Select Case key
        Case "Now": InterestCategoryTitle = "現在の関心"
        Case "Past": InterestCategoryTitle = "過去の関心"
        Case "Want": InterestCategoryTitle = "やりたいこと"
        Case "Social": InterestCategoryTitle = "社会参加"
    End Select
End Function

Public Function InterestLabels(ByVal key As String) As Variant
    Select Case key
        Case "Now"
            InterestLabels = Array("テレビ・新聞", "家事", "散歩", "趣味", "人と話す")
        Case "Past"
            InterestLabels = Array("仕事", "家事・役割", "趣味活動", "外出・旅行", "地域活動")
        Case "Want"
            InterestLabels = Array("散歩・運動", "買い物", "趣味活動", "外出・旅行", "家のこと")
        Case "Social"
            InterestLabels = Array("買い物", "家族との時間", "友人交流", "地域活動", "外出")
        Case Else
            InterestLabels = Array()
    End Select
End Function

Private Sub SaveInterestCategory(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, ByVal key As String)
    Dim labels As Variant
    labels = InterestLabels(key)

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
        picks.Add INTEREST_OTHER_PREFIX & otherText
    End If

    ws.Cells(r, EnsureHeaderColLocal(ws, InterestHeaderName(key))).value = JoinCollection(picks, SEP)
End Sub

Private Sub LoadInterestCategory(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, ByVal key As String)
    Dim labels As Variant
    labels = InterestLabels(key)

    Dim col As Long
    col = FindHeaderColLocal(ws, InterestHeaderName(key))
    If col = 0 Then Exit Sub

    Dim raw As String
    raw = Trim$(CStr(ws.Cells(r, col).value))

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        SetCheckboxValue owner, "chkInterest_" & key & "_" & CStr(i), False
    Next i
    SetCheckboxValue owner, "chkInterest_" & key & "_Other", False
    SetTextboxText owner, "txtInterest_" & key & "_Other", ""

    If LenB(raw) = 0 Then Exit Sub

    Dim tokens As Variant
    tokens = Split(raw, SEP)

    Dim token As Variant
    For Each token In tokens
        ApplyInterestToken owner, key, labels, Trim$(CStr(token))
    Next token
End Sub

Private Sub ApplyInterestToken(ByVal owner As Object, ByVal key As String, ByVal labels As Variant, ByVal token As String)
    If LenB(token) = 0 Then Exit Sub

    If StrComp(Left$(token, Len(INTEREST_OTHER_PREFIX)), INTEREST_OTHER_PREFIX, vbTextCompare) = 0 Then
        SetCheckboxValue owner, "chkInterest_" & key & "_Other", True
        SetTextboxText owner, "txtInterest_" & key & "_Other", Mid$(token, Len(INTEREST_OTHER_PREFIX) + 1)
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

Private Function EnsureHeaderColLocal(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Long
    c = FindHeaderColLocal(ws, header)
    If c > 0 Then
        EnsureHeaderColLocal = c
        Exit Function
    End If

    c = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
    ws.Cells(1, c).value = header
    EnsureHeaderColLocal = c
End Function

Private Function FindHeaderColLocal(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), header, vbTextCompare) = 0 Then
            FindHeaderColLocal = c
            Exit Function
        End If
    Next c
End Function


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

Private Function FindControlByNameDeep(ByVal container As Object, ByVal controlName As String) As Object
    On Error GoTo EH
    If container Is Nothing Then Exit Function

    If StrComp(GetObjectName(container), controlName, vbTextCompare) = 0 Then
        Set FindControlByNameDeep = container
        Exit Function
    End If

    Dim pages As Object
    Set pages = TryGetObjectMember(container, "Pages")
    If Not pages Is Nothing Then
        Dim pg As Object
        For Each pg In pages
            Set FindControlByNameDeep = FindControlByNameDeep(pg, controlName)
            If Not FindControlByNameDeep Is Nothing Then Exit Function
        Next pg
    End If

    Dim controls As Object
    Set controls = TryGetObjectMember(container, "Controls")
    If controls Is Nothing Then Exit Function

    Dim child As Object
    For Each child In controls
        Set FindControlByNameDeep = FindControlByNameDeep(child, controlName)
        If Not FindControlByNameDeep Is Nothing Then Exit Function

    Next child

    Exit Function
EH:
    Err.Clear
End Function

Private Function TryGetObjectMember(ByVal target As Object, ByVal memberName As String) As Object
    On Error GoTo EH
    Select Case memberName
        Case "Controls"
            Set TryGetObjectMember = target.controls
        Case "Pages"
            Set TryGetObjectMember = target.pages
    End Select
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
    SanitizeFreeText = Replace$(src, SEP, SEP_SAFE)
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To values.count
        If LenB(JoinCollection) > 0 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(values(i))
    Next i
End Function

