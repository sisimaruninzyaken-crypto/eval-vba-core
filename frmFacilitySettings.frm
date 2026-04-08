VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFacilitySettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmFacilitySettings.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFacilitySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSaved As Boolean
Private mButtonHooks As Collection

Private txtFacilityName As MSForms.TextBox
Private txtFacilityNo As MSForms.TextBox
Private txtFacilityAddress As MSForms.TextBox
Private txtFacilityPhone As MSForms.TextBox

Public Property Get IsSaved() As Boolean
    IsSaved = mSaved
End Property

Private Sub UserForm_Initialize()
    Me.caption = "事業所設定（初回のみ）"
    Me.Width = 380
    Me.Height = 260

    Set mButtonHooks = New Collection
    BuildFormControls
    LoadExistingValues
End Sub

Private Sub BuildFormControls()
    Dim labelLeft As Single: labelLeft = 12
    Dim inputLeft As Single: inputLeft = 104
    Dim topPos As Single: topPos = 18
    Dim rowHeight As Single: rowHeight = 30

    AddLabel "lblFacilityName", "事業所名", labelLeft, topPos
    Set txtFacilityName = AddTextBox("txtFacilityName", inputLeft, topPos - 2, 250)
    SetImeToJapanese txtFacilityName

    topPos = topPos + rowHeight
    AddLabel "lblFacilityNo", "事業所No", labelLeft, topPos
    Set txtFacilityNo = AddTextBox("txtFacilityNo", inputLeft, topPos - 2, 250)

    topPos = topPos + rowHeight
    AddLabel "lblFacilityAddress", "住所", labelLeft, topPos
    Set txtFacilityAddress = AddTextBox("txtFacilityAddress", inputLeft, topPos - 2, 250)
    SetImeToJapanese txtFacilityAddress

    topPos = topPos + rowHeight
    AddLabel "lblFacilityPhone", "電話番号", labelLeft, topPos
    Set txtFacilityPhone = AddTextBox("txtFacilityPhone", inputLeft, topPos - 2, 250)

    Dim btnSave As MSForms.CommandButton
    Set btnSave = Me.controls.Add("Forms.CommandButton.1", "btnSave", True)
    btnSave.caption = "保存"
    btnSave.Left = 208
    btnSave.top = topPos + 40
    btnSave.Width = 68

    Dim saveHook As clsFacilityButtonHook
    Set saveHook = New clsFacilityButtonHook
    saveHook.Init Me, btnSave, "save"
    mButtonHooks.Add saveHook

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.caption = "キャンセル"
    btnCancel.Left = 286
    btnCancel.top = topPos + 40
    btnCancel.Width = 80

    Dim cancelHook As clsFacilityButtonHook
    Set cancelHook = New clsFacilityButtonHook
    cancelHook.Init Me, btnCancel, "cancel"
    mButtonHooks.Add cancelHook
End Sub

Private Sub LoadExistingValues()
    Dim facilityName As String
    Dim facilityNo As String
    Dim facilityAddress As String
    Dim facilityPhone As String

    modAppConfig.LoadFacilitySettings facilityName, facilityNo, facilityAddress, facilityPhone
    facilityName = ResolveFacilityNameForOutput(owner, facilityName)

    txtFacilityName.text = facilityName
    txtFacilityNo.text = facilityNo
    txtFacilityAddress.text = facilityAddress
    txtFacilityPhone.text = facilityPhone
End Sub

Public Sub HandleButtonClick(ByVal actionName As String)
    Select Case LCase$(actionName)
        Case "save"
            SaveAndClose
        Case "cancel"
            CancelAndClose
    End Select
End Sub

Private Sub SaveAndClose()
    If Not ValidateRequired() Then Exit Sub

    modAppConfig.SaveFacilitySettings txtFacilityName.text, txtFacilityNo.text, txtFacilityAddress.text, txtFacilityPhone.text
    mSaved = True
    Me.Hide
End Sub

Private Sub CancelAndClose()
    mSaved = False
    Me.Hide
End Sub

Private Function ValidateRequired() As Boolean
    If Len(Trim$(txtFacilityName.text)) = 0 Then
        MsgBox "事業所名を入力してください。", vbExclamation
        txtFacilityName.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtFacilityNo.text)) = 0 Then
        MsgBox "事業所Noを入力してください。", vbExclamation
        txtFacilityNo.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtFacilityAddress.text)) = 0 Then
        MsgBox "住所を入力してください。", vbExclamation
        txtFacilityAddress.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtFacilityPhone.text)) = 0 Then
        MsgBox "電話番号を入力してください。", vbExclamation
        txtFacilityPhone.SetFocus
        Exit Function
    End If

    ValidateRequired = True
End Function

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single) As MSForms.label
    Dim lbl As MSForms.label
    Set lbl = Me.controls.Add("Forms.Label.1", controlName, True)
    lbl.caption = captionText
    lbl.Left = leftPos
    lbl.top = topPos
    lbl.Width = 88
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

Private Sub SetImeToJapanese(ByVal targetTextBox As MSForms.TextBox)
    targetTextBox.IMEMode = fmIMEModeHiragana
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        mSaved = False
    End If
End Sub

