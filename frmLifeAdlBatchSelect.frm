VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLifeAdlBatchSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmLifeAdlBatchSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLifeAdlBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents lstCandidates As MSForms.ListBox
Attribute lstCandidates.VB_VarHelpID = -1
Private WithEvents btnRefresh As MSForms.CommandButton
Attribute btnRefresh.VB_VarHelpID = -1
Private WithEvents btnFirstDue As MSForms.CommandButton
Attribute btnFirstDue.VB_VarHelpID = -1
Private WithEvents btnClear As MSForms.CommandButton
Attribute btnClear.VB_VarHelpID = -1
Private WithEvents btnExportBatch As MSForms.CommandButton
Attribute btnExportBatch.VB_VarHelpID = -1
Private WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

Private mCandidates As Collection
Private mLoading As Boolean

Private Sub UserForm_Initialize()
    Me.caption = "ADL CSV" & JPText(&H4E00, &H62EC, &H5BFE, &H8C61, &H8005, &H9078, &H629E)
    Me.Width = 860
    Me.Height = 430
    Me.StartUpPosition = 1

    BuildControls
    LoadCandidates
End Sub

Public Sub LoadCandidates()
    Dim item As Object
    Dim rowIndex As Long

    mLoading = True
    lstCandidates.Clear
    Set mCandidates = modLifeAdlBatchSelect.BuildLifeAdlBatchCandidates()

    For Each item In mCandidates
        lstCandidates.AddItem IIf(CBool(item("Selected")), "ON", vbNullString)
        rowIndex = lstCandidates.ListCount - 1
        lstCandidates.List(rowIndex, 1) = CStr(item("Name"))
        lstCandidates.List(rowIndex, 2) = CStr(item("EvaluateDate"))
        lstCandidates.List(rowIndex, 3) = CStr(item("Status"))
        lstCandidates.List(rowIndex, 4) = CStr(item("InsurerNo"))
        lstCandidates.List(rowIndex, 5) = CStr(item("InsuredNo"))
        lstCandidates.List(rowIndex, 6) = CStr(item("ExternalSystemKey"))
        lstCandidates.List(rowIndex, 7) = CStr(item("MissingReason"))
        lstCandidates.Selected(rowIndex) = CBool(item("Selected"))
    Next item
    mLoading = False
End Sub

Private Sub BuildControls()
    Dim topList As Single
    topList = 36

    AddHeaderLabel "hdrSelect", JPText(&H51FA, &H529B), 12, 14, 40
    AddHeaderLabel "hdrName", JPText(&H6C0F, &H540D), 58, 14, 82
    AddHeaderLabel "hdrEvalDate", JPText(&H8A55, &H4FA1, &H65E5), 150, 14, 70
    AddHeaderLabel "hdrStatus", JPText(&H72B6, &H614B), 226, 14, 68
    AddHeaderLabel "hdrInsurer", JPText(&H4FDD, &H967A, &H8005, &H756A, &H53F7), 304, 14, 76
    AddHeaderLabel "hdrInsured", JPText(&H88AB, &H4FDD, &H967A, &H8005, &H756A, &H53F7), 390, 14, 88
    AddHeaderLabel "hdrExternal", JPText(&H5916, &H90E8, &H30B7, &H30B9, &H30C6, &H30E0, &H7BA1, &H7406, &H756A, &H53F7), 490, 14, 128
    AddHeaderLabel "hdrMissing", JPText(&H4E0D, &H8DB3, &H7406, &H7531), 632, 14, 160

    Set lstCandidates = Me.controls.Add("Forms.ListBox.1", "lstCandidates", True)
    With lstCandidates
        .Left = 10
        .top = topList
        .Width = 820
        .Height = 300
        .ColumnCount = 8
        .ColumnWidths = "40 pt;86 pt;70 pt;70 pt;78 pt;92 pt;132 pt;210 pt"
        .MultiSelect = fmMultiSelectMulti
        .IntegralHeight = False
    End With

    Set btnRefresh = AddButton("btnRefresh", JPText(&H518D, &H8AAD, &H8FBC), 10, 350, 82)
    Set btnFirstDue = AddButton("btnFirstDue", "FIRST/DUE", 104, 350, 90)
    Set btnClear = AddButton("btnClear", JPText(&H89E3, &H9664), 206, 350, 68)
    Set btnExportBatch = AddButton("btnExportBatch", "CSV" & JPText(&H51FA, &H529B), 642, 350, 94)
    Set btnClose = AddButton("btnClose", "OK", 748, 350, 82)
End Sub

Private Function AddHeaderLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthValue As Single) As MSForms.label
    Dim lbl As MSForms.label
    Set lbl = Me.controls.Add("Forms.Label.1", controlName, True)
    lbl.caption = captionText
    lbl.Left = leftPos
    lbl.top = topPos
    lbl.Width = widthValue
    lbl.Height = 16
    Set AddHeaderLabel = lbl
End Function

Private Function AddButton(ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthValue As Single) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = Me.controls.Add("Forms.CommandButton.1", controlName, True)
    btn.caption = captionText
    btn.Left = leftPos
    btn.top = topPos
    btn.Width = widthValue
    Set AddButton = btn
End Function

Private Sub lstCandidates_Change()
    SyncSelectionMarks
End Sub

Private Sub btnRefresh_Click()
    LoadCandidates
End Sub

Private Sub btnFirstDue_Click()
    Dim i As Long
    If lstCandidates.ListCount = 0 Then Exit Sub
    mLoading = True
    For i = 0 To lstCandidates.ListCount - 1
        lstCandidates.Selected(i) = modLifeAdlBatchSelect.LifeAdlBatchShouldSelect(CStr(lstCandidates.List(i, 3)))
    Next i
    mLoading = False
    SyncSelectionMarks
End Sub

Private Sub btnClear_Click()
    Dim i As Long
    If lstCandidates.ListCount = 0 Then Exit Sub
    mLoading = True
    For i = 0 To lstCandidates.ListCount - 1
        lstCandidates.Selected(i) = False
    Next i
    mLoading = False
    SyncSelectionMarks
End Sub


Private Sub btnExportBatch_Click()
    Dim selectedItems As Collection
    Set selectedItems = BuildSelectedCandidates()
    modLifeCsvAdlExport.ExportLifeAdlBatchCsvFromCandidates selectedItems
End Sub

Private Function BuildSelectedCandidates() As Collection
    Dim result As Collection
    Dim i As Long
    Dim item As Object

    Set result = New Collection
    If mCandidates Is Nothing Then
        Set BuildSelectedCandidates = result
        Exit Function
    End If

    For i = 0 To lstCandidates.ListCount - 1
        If lstCandidates.Selected(i) Then
            Set item = mCandidates(i + 1)
            item("Selected") = True
            result.Add item
        Else
            mCandidates(i + 1)("Selected") = False
        End If
    Next i

    Set BuildSelectedCandidates = result
End Function
Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub SyncSelectionMarks()
    Dim i As Long
    If mLoading Then Exit Sub
    For i = 0 To lstCandidates.ListCount - 1
        If lstCandidates.Selected(i) Then
            lstCandidates.List(i, 0) = "ON"
        Else
            lstCandidates.List(i, 0) = vbNullString
        End If
    Next i
End Sub

Private Function JPText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    For i = LBound(codePoints) To UBound(codePoints)
        JPText = JPText & ChrW$(CLng(codePoints(i)))
    Next i
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
