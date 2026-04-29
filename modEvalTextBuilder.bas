Attribute VB_Name = "modEvalTextBuilder"
Public Sub Preview_NameToHeader()
    Dim f As Object
    Set f = frmEval

    Dim hdr As MSForms.Frame
    Set hdr = f.controls("frHeader")
    Call Align_LoadPrevButton_NextToHdrKana(f)

    Dim btn As MSForms.CommandButton
    Set btn = hdr.controls("cmdClearHeader")
    
    Dim gap As Single: gap = 10
    Dim pad As Single: pad = 8

    Dim lbl As MSForms.label
    Dim txt As MSForms.TextBox
    Dim lblKana As MSForms.label
    Dim txtKana As MSForms.TextBox

    On Error Resume Next
    Set lbl = hdr.controls("lblHdrName")
    On Error GoTo 0
    If lbl Is Nothing Then
        Set lbl = hdr.controls.Add("Forms.Label.1", "lblHdrName", True)
        lbl.caption = ChrW(&H6C0F) & ChrW(&H540D)
        lbl.AutoSize = True
        lbl.Width = lbl.Width + 8
    End If

    On Error Resume Next
    Set txt = hdr.controls("txtHdrName")
    On Error GoTo 0
    If txt Is Nothing Then
        Set txt = hdr.controls.Add("Forms.TextBox.1", "txtHdrName", True)
        txt.SpecialEffect = f.controls("txtName").SpecialEffect
        txt.Font.name = f.controls("txtName").Font.name
        txt.Font.Size = f.controls("txtName").Font.Size
        txt.Height = f.controls("txtName").Height
        txt.Width = f.controls("txtName").Width
    End If
    txt.IMEMode = fmIMEModeHiragana

    On Error Resume Next
    Set lblKana = hdr.controls("lblHdrKana")
    On Error GoTo 0
    If lblKana Is Nothing Then
        Set lblKana = hdr.controls.Add("Forms.Label.1", "lblHdrKana", True)
        lblKana.caption = ChrW(&H3075) & ChrW(&H308A) & ChrW(&H304C) & ChrW(&H306A)
        lblKana.AutoSize = True
        lblKana.Width = lblKana.Width + 8
    End If

    On Error Resume Next
    Set txtKana = hdr.controls("txtHdrKana")
    On Error GoTo 0
    If txtKana Is Nothing Then
        Set txtKana = hdr.controls.Add("Forms.TextBox.1", "txtHdrKana", True)
        txtKana.SpecialEffect = txt.SpecialEffect
        txtKana.Font.name = txt.Font.name
        txtKana.Font.Size = txt.Font.Size
        txtKana.Height = txt.Height
        txtKana.Width = txt.Width
    End If
    txtKana.IMEMode = fmIMEModeHiragana
    txt.text = f.controls("txtName").text

    txt.top = btn.top + (btn.Height - txt.Height) / 2
    lbl.top = btn.top + (btn.Height - lbl.Height) / 2
    txt.Left = btn.Left - pad - txt.Width
    lbl.Left = txt.Left - gap - lbl.Width

    txtKana.Left = btn.Left - pad - txtKana.Width
    txtKana.top = 36
    lblKana.top = txtKana.top + (txtKana.Height - lblKana.Height) / 2
    lblKana.Left = txtKana.Left - gap - lblKana.Width

    Dim lblID As MSForms.label
    Dim txtID As MSForms.TextBox
    On Error Resume Next
    Set lblID = hdr.controls("lblHdrPID")
    On Error GoTo 0
    If lblID Is Nothing Then
        Set lblID = hdr.controls.Add("Forms.Label.1", "lblHdrPID", True)
        lblID.caption = "ID"
        lblID.AutoSize = True
        lblID.Width = lblID.Width + 8
    End If

    On Error Resume Next
    Set txtID = hdr.controls("txtHdrPID")
    On Error GoTo 0
    If txtID Is Nothing Then
        Set txtID = hdr.controls.Add("Forms.TextBox.1", "txtHdrPID", True)
        txtID.SpecialEffect = f.controls("txtPID").SpecialEffect
        txtID.Font.name = f.controls("txtPID").Font.name
        txtID.Font.Size = f.controls("txtPID").Font.Size
        txtID.Height = f.controls("txtPID").Height
        txtID.Width = f.controls("txtPID").Width
    End If
    txtID.text = f.controls("txtPID").text
    txtID.top = btn.top + (btn.Height - txtID.Height) / 2
    lblID.top = btn.top + (btn.Height - lblID.Height) / 2
    txtID.Left = lbl.Left - pad - txtID.Width
    lblID.Left = txtID.Left - gap - lblID.Width

    Dim lblInsured As MSForms.label
    Dim txtInsured As MSForms.TextBox
    On Error Resume Next
    Set lblInsured = hdr.controls("lblHdrInsuredNo")
    On Error GoTo 0
    If lblInsured Is Nothing Then
        Set lblInsured = hdr.controls.Add("Forms.Label.1", "lblHdrInsuredNo", True)
        lblInsured.caption = ChrW(&H88AB) & ChrW(&H4FDD) & ChrW(&H967A) & ChrW(&H8005) & ChrW(&H756A) & ChrW(&H53F7)
        lblInsured.AutoSize = True
        lblInsured.Width = lblInsured.Width + 8
    End If

    On Error Resume Next
    Set txtInsured = hdr.controls("txtInsuredNo")
    On Error GoTo 0
    If txtInsured Is Nothing Then
        Set txtInsured = hdr.controls.Add("Forms.TextBox.1", "txtInsuredNo", True)
        txtInsured.SpecialEffect = txtID.SpecialEffect
        txtInsured.Font.name = txtID.Font.name
        txtInsured.Font.Size = txtID.Font.Size
        txtInsured.Height = txtID.Height
        txtInsured.Width = 132
    End If
    txtInsured.IMEMode = fmIMEModeOff

    Dim lblInsurer As MSForms.label
    Dim txtInsurer As MSForms.TextBox
    On Error Resume Next
    Set lblInsurer = hdr.controls("lblHdrInsurerNo")
    On Error GoTo 0
    If lblInsurer Is Nothing Then
        Set lblInsurer = hdr.controls.Add("Forms.Label.1", "lblHdrInsurerNo", True)
        lblInsurer.caption = ChrW(&H4FDD) & ChrW(&H967A) & ChrW(&H8005) & ChrW(&H756A) & ChrW(&H53F7)
        lblInsurer.AutoSize = True
        lblInsurer.Width = lblInsurer.Width + 8
    End If

    On Error Resume Next
    Set txtInsurer = hdr.controls("txtInsurerNo")
    On Error GoTo 0
    If txtInsurer Is Nothing Then
        Set txtInsurer = hdr.controls.Add("Forms.TextBox.1", "txtInsurerNo", True)
        txtInsurer.SpecialEffect = txtInsured.SpecialEffect
        txtInsurer.Font.name = txtInsured.Font.name
        txtInsurer.Font.Size = txtInsured.Font.Size
        txtInsurer.Height = txtInsured.Height
        txtInsurer.Width = 132
    End If
    txtInsurer.IMEMode = fmIMEModeOff

    Dim lblExternal As MSForms.label
    Dim txtExternal As MSForms.TextBox
    On Error Resume Next
    Set lblExternal = hdr.controls("lblHdrExternalSystemKey")
    On Error GoTo 0
    If lblExternal Is Nothing Then
        Set lblExternal = hdr.controls.Add("Forms.Label.1", "lblHdrExternalSystemKey", True)
        lblExternal.caption = ChrW(&H5916) & ChrW(&H90E8) & ChrW(&H30B7) & ChrW(&H30B9) & ChrW(&H30C6) & ChrW(&H30E0) & ChrW(&H7BA1) & ChrW(&H7406) & ChrW(&H756A) & ChrW(&H53F7)
        lblExternal.AutoSize = True
        lblExternal.Width = lblExternal.Width + 8
    End If

    On Error Resume Next
    Set txtExternal = hdr.controls("txtExternalSystemKey")
    On Error GoTo 0
    If txtExternal Is Nothing Then
        Set txtExternal = hdr.controls.Add("Forms.TextBox.1", "txtExternalSystemKey", True)
        txtExternal.SpecialEffect = txtInsurer.SpecialEffect
        txtExternal.Font.name = txtInsurer.Font.name
        txtExternal.Font.Size = txtInsurer.Font.Size
        txtExternal.Height = txtInsurer.Height
        txtExternal.Width = 220
    End If
    txtExternal.IMEMode = fmIMEModeOff

    frmEval.RearrangeHeaderTopAreaLayout
End Sub
