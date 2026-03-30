VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRomMirrorBtnHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1


Private Sub btn_Click()
    On Error Resume Next

    Dim c As Object
    For Each c In btn.parent.Controls
        If typeName(c) = "TextBox" Then
            Dim nm As String
            nm = CStr(c.name)
            If Right$(nm, 2) = "_R" Then
                Dim leftName As String
                leftName = Left$(nm, Len(nm) - 2) & "_L"
                Dim tL As MSForms.TextBox
                Set tL = btn.parent.Controls(leftName)
                If Not tL Is Nothing Then
                    If Trim$(CStr(c.text)) <> "" Then
                        tL.text = c.text
                    End If
                End If
            End If
        End If
    Next
End Sub

