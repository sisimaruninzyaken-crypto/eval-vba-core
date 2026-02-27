Attribute VB_Name = "modLayoutHeader"
Option Explicit

Public Sub Align_LoadPrevButton_NextToHdrKana(ByVal f As Object)
    Dim hdr As Object
    Dim kana As Object
    Dim btn As Object
    Dim refBtn As Object

    On Error Resume Next
    Set hdr = f.Controls("frHeader")
    If hdr Is Nothing Then Exit Sub

    Set kana = hdr.Controls("txtHdrKana")
    If kana Is Nothing Then Exit Sub

    Set btn = hdr.Controls("cmdHdrLoadPrev")
    If btn Is Nothing Then Set btn = f.Controls("btnLoadPrevCtl")
    If btn Is Nothing Then Exit Sub

    Set refBtn = hdr.Controls("cmdSaveHeader")
    If refBtn Is Nothing Then Set refBtn = hdr.Controls("cmdClearHeader")
    If refBtn Is Nothing Then Set refBtn = hdr.Controls("cmdCloseHeader")
    On Error GoTo 0

    btn.Width = 180
    btn.Height = 24

    If Not refBtn Is Nothing Then
        btn.Font.name = refBtn.Font.name
        btn.Font.Size = refBtn.Font.Size
        On Error Resume Next: btn.SpecialEffect = refBtn.SpecialEffect: On Error GoTo 0
    End If

    If btn.parent Is hdr Then
        btn.Left = kana.Left + kana.Width + 12
        btn.Top = kana.Top + 2
    Else
        btn.Left = hdr.Left + kana.Left + kana.Width + 12
        btn.Top = hdr.Top + kana.Top + 2
    End If
End Sub
