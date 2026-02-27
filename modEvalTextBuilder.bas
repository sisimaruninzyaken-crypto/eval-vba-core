Attribute VB_Name = "modEvalTextBuilder"
Public Sub Preview_NameToHeader()
    Dim f As Object
    Set f = frmEval



    Dim hdr As MSForms.Frame
    Set hdr = f.Controls("frHeader")
    Call Align_LoadPrevButton_NextToHdrKana(f)

    Dim btn As MSForms.CommandButton
    Set btn = hdr.Controls("cmdClearHeader")
    
    Dim gap As Single: gap = 10
    Dim pad As Single: pad = 8

    Dim lbl As MSForms.label
    Dim txt As MSForms.TextBox
    Dim lblKana As MSForms.label
    Dim txtKana As MSForms.TextBox

    '--- create or get header label ---
    On Error Resume Next
    Set lbl = hdr.Controls("lblHdrName")
    On Error GoTo 0
    If lbl Is Nothing Then
    Set lbl = hdr.Controls.Add("Forms.Label.1", "lblHdrName", True)
    lbl.caption = "éÅñº"
    lbl.AutoSize = True
    lbl.Width = lbl.Width + 8   ' Å© Ç±Ç±
End If


    '--- create or get header textbox ---
    On Error Resume Next
    Set txt = hdr.Controls("txtHdrName")
    On Error GoTo 0
    If txt Is Nothing Then
        Set txt = hdr.Controls.Add("Forms.TextBox.1", "txtHdrName", True)
        txt.SpecialEffect = f.Controls("txtName").SpecialEffect
        txt.Font.name = f.Controls("txtName").Font.name
        txt.Font.Size = f.Controls("txtName").Font.Size
        txt.Height = f.Controls("txtName").Height
        txt.Width = f.Controls("txtName").Width
    End If


        txt.IMEMode = fmIMEModeHiragana

    '--- create or get header kana label/textbox ---
    On Error Resume Next
    Set lblKana = hdr.Controls("lblHdrKana")
    On Error GoTo 0
    If lblKana Is Nothing Then
        Set lblKana = hdr.Controls.Add("Forms.Label.1", "lblHdrKana", True)
        lblKana.caption = "Ç”ÇËÇ™Ç»"
        lblKana.AutoSize = True
        lblKana.Width = lblKana.Width + 8
    End If

    On Error Resume Next
    Set txtKana = hdr.Controls("txtHdrKana")
    On Error GoTo 0
    If txtKana Is Nothing Then
        Set txtKana = hdr.Controls.Add("Forms.TextBox.1", "txtHdrKana", True)
        txtKana.SpecialEffect = txt.SpecialEffect
        txtKana.Font.name = txt.Font.name
        txtKana.Font.Size = txt.Font.Size
        txtKana.Height = txt.Height
        txtKana.Width = txt.Width
    End If

    txtKana.IMEMode = fmIMEModeHiragana
   'Call frmEval.EnsureHeaderLoadPrevButton

    '--- value sync (one-way preview) ---
    txt.Text = f.Controls("txtName").Text

    '--- position: [éÅñº][txt] [cmdClearHeader][cmdSaveHeader][cmdCloseHeader] ---
    txt.Top = btn.Top + (btn.Height - txt.Height) / 2
    lbl.Top = btn.Top + (btn.Height - lbl.Height) / 2

    txt.Left = btn.Left - pad - txt.Width
    lbl.Left = txt.Left - gap - lbl.Width

    txtKana.Left = btn.Left - pad - txtKana.Width
    txtKana.Top = 36
    lblKana.Top = txtKana.Top + (txtKana.Height - lblKana.Height) / 2
    lblKana.Left = txtKana.Left - gap - lblKana.Width

        '--- create or get header PID label/textbox ---
    Dim lblID As MSForms.label
    Dim txtID As MSForms.TextBox

    On Error Resume Next
    Set lblID = hdr.Controls("lblHdrPID")
    On Error GoTo 0
    If lblID Is Nothing Then
        Set lblID = hdr.Controls.Add("Forms.Label.1", "lblHdrPID", True)
        lblID.caption = "ID"
        lblID.AutoSize = True
        lblID.Width = lblID.Width + 8
    End If

    On Error Resume Next
    Set txtID = hdr.Controls("txtHdrPID")
    On Error GoTo 0
    If txtID Is Nothing Then
        Set txtID = hdr.Controls.Add("Forms.TextBox.1", "txtHdrPID", True)
        txtID.SpecialEffect = f.Controls("txtPID").SpecialEffect
        txtID.Font.name = f.Controls("txtPID").Font.name
        txtID.Font.Size = f.Controls("txtPID").Font.Size
        txtID.Height = f.Controls("txtPID").Height
        txtID.Width = f.Controls("txtPID").Width
    End If

    '--- value sync (one-way preview) ---
    txtID.Text = f.Controls("txtPID").Text

    '--- position: [ID][txt] [éÅñº][txt] [buttons...] ---
    txtID.Top = btn.Top + (btn.Height - txtID.Height) / 2
    lblID.Top = btn.Top + (btn.Height - lblID.Height) / 2

    txtID.Left = lbl.Left - pad - txtID.Width
    lblID.Left = txtID.Left - gap - lblID.Width




End Sub








