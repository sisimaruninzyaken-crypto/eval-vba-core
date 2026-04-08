Attribute VB_Name = "modEvalTextBuilder"
Public Sub Preview_NameToHeader()
    Dim f As Object
    Set f = frmEval



    Dim hdr As MSForms.Frame
    Set hdr = f.controls("frHeader")
    Call Align_LoadPrevButton_NextToHdrKana(f)

    Dim Btn As MSForms.CommandButton
    Set Btn = hdr.controls("cmdClearHeader")
    
    Dim gap As Single: gap = 10
    Dim pad As Single: pad = 8

    Dim lbl As MSForms.label
    Dim txt As MSForms.TextBox
    Dim lblKana As MSForms.label
    Dim txtKana As MSForms.TextBox

    '--- create or get header label ---
    On Error Resume Next
    Set lbl = hdr.controls("lblHdrName")
    On Error GoTo 0
    If lbl Is Nothing Then
    Set lbl = hdr.controls.Add("Forms.Label.1", "lblHdrName", True)
    lbl.caption = "éÅñº"
    lbl.AutoSize = True
    lbl.Width = lbl.Width + 8   ' Å© Ç±Ç±
End If


    '--- create or get header textbox ---
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

    '--- create or get header kana label/textbox ---
    On Error Resume Next
    Set lblKana = hdr.controls("lblHdrKana")
    On Error GoTo 0
    If lblKana Is Nothing Then
        Set lblKana = hdr.controls.Add("Forms.Label.1", "lblHdrKana", True)
        lblKana.caption = "Ç”ÇËÇ™Ç»"
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
   'Call frmEval.EnsureHeaderLoadPrevButton

    '--- value sync (one-way preview) ---
    txt.text = f.controls("txtName").text

    '--- position: [éÅñº][txt] [cmdClearHeader][cmdSaveHeader][cmdCloseHeader] ---
    txt.top = Btn.top + (Btn.Height - txt.Height) / 2
    lbl.top = Btn.top + (Btn.Height - lbl.Height) / 2

    txt.Left = Btn.Left - pad - txt.Width
    lbl.Left = txt.Left - gap - lbl.Width

    txtKana.Left = Btn.Left - pad - txtKana.Width
    txtKana.top = 36
    lblKana.top = txtKana.top + (txtKana.Height - lblKana.Height) / 2
    lblKana.Left = txtKana.Left - gap - lblKana.Width

        '--- create or get header PID label/textbox ---
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

    '--- value sync (one-way preview) ---
    txtID.text = f.controls("txtPID").text

    '--- position: [ID][txt] [éÅñº][txt] [buttons...] ---
    txtID.top = Btn.top + (Btn.Height - txtID.Height) / 2
    lblID.top = Btn.top + (Btn.Height - lblID.Height) / 2

    txtID.Left = lbl.Left - pad - txtID.Width
    lblID.Left = txtID.Left - gap - lblID.Width




End Sub








