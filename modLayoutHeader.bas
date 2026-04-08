Attribute VB_Name = "modLayoutHeader"
Option Explicit

Public Sub Align_LoadPrevButton_NextToHdrKana(ByVal f As Object)
    Dim hdr As Object
    Dim kana As Object
    Dim Btn As Object
    Dim refBtn As Object

    Set hdr = SafeGetControl(f, "frHeader")
    If hdr Is Nothing Then Exit Sub

    Set kana = SafeGetControl(hdr, "txtHdrKana")
    If kana Is Nothing Then Exit Sub

    Set Btn = SafeGetControl(hdr, "cmdHdrLoadPrev")
    If Btn Is Nothing Then Exit Sub

    Set refBtn = SafeGetControl(hdr, "cmdSaveHeader")
    If refBtn Is Nothing Then Set refBtn = SafeGetControl(hdr, "cmdClearHeader")
    If refBtn Is Nothing Then Set refBtn = SafeGetControl(hdr, "cmdCloseHeader")

    Btn.Width = 180
    Btn.Height = 24

    If Not refBtn Is Nothing Then
        Btn.Font.name = refBtn.Font.name
        Btn.Font.Size = refBtn.Font.Size
        On Error Resume Next: Btn.SpecialEffect = refBtn.SpecialEffect: On Error GoTo 0
    End If

    If Btn.parent Is hdr Then
        Btn.Left = kana.Left + kana.Width + 12
        Btn.top = kana.top + 2
    Else
        Btn.Left = hdr.Left + kana.Left + kana.Width + 12
        Btn.top = hdr.top + kana.top + 2
    End If
End Sub
