VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLifeLink 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmLifeLink.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLifeLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOwner As frmEval
Private mButtonHooks As Collection

Private txtInsuredNo As MSForms.TextBox
Private txtInsurerNo As MSForms.TextBox
Private txtExternalSystemKey As MSForms.TextBox

Public Sub InitWithOwner(ByVal owner As frmEval)
    Set mOwner = owner
    LoadFromOwner
End Sub

Private Sub UserForm_Initialize()
    Me.caption = "LIFE" & ChrW(&H8A2D) & ChrW(&H5B9A)
    Me.Width = 430
    Me.Height = 238
    Me.StartUpPosition = 1

    Set mButtonHooks = New Collection
    BuildFormControls
End Sub
Private Sub BuildFormControls()
    Dim labelLeft As Single: labelLeft = 14
    Dim inputLeft As Single: inputLeft = 150
    Dim topPos As Single: topPos = 18
    Dim rowGap As Single: rowGap = 34
    Dim buttonTop As Single

    AddLabel "lblInsuredNo", ChrW(&H88AB) & ChrW(&H4FDD) & ChrW(&H967A) & ChrW(&H8005) & ChrW(&H756A) & ChrW(&H53F7), labelLeft, topPos
    Set txtInsuredNo = AddTextBox("txtInsuredNo", inputLeft, topPos - 2, 240)
    txtInsuredNo.IMEMode = fmIMEModeOff

    topPos = topPos + rowGap
    AddLabel "lblInsurerNo", ChrW(&H4FDD) & ChrW(&H967A) & ChrW(&H8005) & ChrW(&H756A) & ChrW(&H53F7), labelLeft, topPos
    Set txtInsurerNo = AddTextBox("txtInsurerNo", inputLeft, topPos - 2, 240)
    txtInsurerNo.IMEMode = fmIMEModeOff

    topPos = topPos + rowGap
    AddLabel "lblExternalSystemKey", ChrW(&H5916) & ChrW(&H90E8) & ChrW(&H30B7) & ChrW(&H30B9) & ChrW(&H30C6) & ChrW(&H30E0) & ChrW(&H7BA1) & ChrW(&H7406) & ChrW(&H756A) & ChrW(&H53F7), labelLeft, topPos
    Set txtExternalSystemKey = AddTextBox("txtExternalSystemKey", inputLeft, topPos - 2, 240)
    txtExternalSystemKey.IMEMode = fmIMEModeOff

    buttonTop = topPos + 38

    Dim btnBatch As MSForms.CommandButton
    Set btnBatch = Me.controls.Add("Forms.CommandButton.1", "btnOpenLifeAdlBatchSelect", True)
    btnBatch.caption = "ADL" & ChrW(&H4E00) & ChrW(&H62EC) & ChrW(&H9078) & ChrW(&H629E)
    btnBatch.Left = 14
    btnBatch.top = buttonTop
    btnBatch.Width = 112

    Dim batchHook As clsLifeLinkButtonHook
    Set batchHook = New clsLifeLinkButtonHook
    batchHook.Init Me, btnBatch, "batch_adl_select"
    mButtonHooks.Add batchHook
    Dim btnExport As MSForms.CommandButton
    Set btnExport = Me.controls.Add("Forms.CommandButton.1", "btnExportLifeAdlCsv", True)
    btnExport.caption = "ADL CSV" & ChrW(&H51FA) & ChrW(&H529B)
    btnExport.Left = 138
    btnExport.top = buttonTop
    btnExport.Width = 112

    Dim exportHook As clsLifeLinkButtonHook
    Set exportHook = New clsLifeLinkButtonHook
    exportHook.Init Me, btnExport, "export_adl_csv"
    mButtonHooks.Add exportHook

    Dim btnOk As MSForms.CommandButton
    Set btnOk = Me.controls.Add("Forms.CommandButton.1", "btnOk", True)
    btnOk.caption = "OK"
    btnOk.Left = 258
    btnOk.top = buttonTop
    btnOk.Width = 68

    Dim okHook As clsLifeLinkButtonHook
    Set okHook = New clsLifeLinkButtonHook
    okHook.Init Me, btnOk, "ok"
    mButtonHooks.Add okHook

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.caption = ChrW(&H30AD) & ChrW(&H30E3) & ChrW(&H30F3) & ChrW(&H30BB) & ChrW(&H30EB)
    btnCancel.Left = 334
    btnCancel.top = buttonTop
    btnCancel.Width = 80

    Dim cancelHook As clsLifeLinkButtonHook
    Set cancelHook = New clsLifeLinkButtonHook
    cancelHook.Init Me, btnCancel, "cancel"
    mButtonHooks.Add cancelHook
End Sub

Private Sub LoadFromOwner()
    If mOwner Is Nothing Then Exit Sub
    If txtInsuredNo Is Nothing Then Exit Sub
    txtInsuredNo.text = mOwner.GetLifeLinkFieldValue("InsuredNo")
    txtInsurerNo.text = mOwner.GetLifeLinkFieldValue("InsurerNo")
    txtExternalSystemKey.text = mOwner.GetLifeLinkFieldValue("ExternalSystemKey")
End Sub

Private Sub SaveToOwner()
    If mOwner Is Nothing Then Exit Sub
    mOwner.SetLifeLinkFieldValue "InsuredNo", txtInsuredNo.text
    mOwner.SetLifeLinkFieldValue "InsurerNo", txtInsurerNo.text
    mOwner.SetLifeLinkFieldValue "ExternalSystemKey", txtExternalSystemKey.text
End Sub
Private Sub ExportAdlCsvFromThisForm()
    SaveToOwner
    If mOwner Is Nothing Then
        MsgBox "frmEval is not open.", vbExclamation
        Exit Sub
    End If
    modLifeCsvAdlExport.ExportLifeAdlCsvFromOwner mOwner
End Sub
Public Sub HandleButtonClick(ByVal actionName As String)
    Select Case LCase$(Trim$(actionName))
        Case "ok"
            SaveToOwner
            Me.Hide
        Case "cancel"
            Me.Hide
        Case "export_adl_csv"
            ExportAdlCsvFromThisForm
        Case "batch_adl_select"
            SaveToOwner
            modLifeAdlBatchSelect.ShowLifeAdlBatchSelect
    End Select
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single) As MSForms.label
    Dim lbl As MSForms.label
    Set lbl = Me.controls.Add("Forms.Label.1", controlName, True)
    lbl.caption = captionText
    lbl.Left = leftPos
    lbl.top = topPos
    lbl.Width = 130
    Set AddLabel = lbl
End Function

Private Function AddTextBox(ByVal controlName As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthValue As Single) As MSForms.TextBox
    Dim txt As MSForms.TextBox
    Set txt = Me.controls.Add("Forms.TextBox.1", controlName, True)
    txt.Left = leftPos
    txt.top = topPos
    txt.Width = widthValue
    Set AddTextBox = txt
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
