Attribute VB_Name = "modHdrProbe"
Option Explicit

Public Sub RunHdrProbe()
    On Error GoTo EH

    Debug.Print "=== HDR PROBE START ==="

    ' frmEvalが開いてなければ開く（既に開いてればそのまま）
    If VBA.UserForms.Count = 0 Then
        frmEval.Show vbModeless
        DoEvents
    End If

    Debug.Print "[Form] Controls.Count=" & frmEval.Controls.Count

    Debug.Print "[frHeader] Type=" & TypeName(frmEval.Controls("frHeader"))
    Debug.Print "[frHeader] InsideW=" & frmEval.Controls("frHeader").InsideWidth & _
                " Left=" & frmEval.Controls("frHeader").Left & _
                " Top=" & frmEval.Controls("frHeader").Top & _
                " W=" & frmEval.Controls("frHeader").Width & _
                " H=" & frmEval.Controls("frHeader").Height

    Debug.Print "[txtHdrKana] Type=" & TypeName(frmEval.Controls("frHeader").Controls("txtHdrKana")) & _
                " Left=" & frmEval.Controls("frHeader").Controls("txtHdrKana").Left & _
                " Top=" & frmEval.Controls("frHeader").Controls("txtHdrKana").Top & _
                " W=" & frmEval.Controls("frHeader").Controls("txtHdrKana").Width & _
                " H=" & frmEval.Controls("frHeader").Controls("txtHdrKana").Height & _
                " Visible=" & frmEval.Controls("frHeader").Controls("txtHdrKana").Visible

    ' 目的のボタン
    Debug.Print "[cmdHdrLoadPrev] Exists? " & (Not (frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev") Is Nothing))
    If Not (frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev") Is Nothing) Then
        Debug.Print "[cmdHdrLoadPrev] Left=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").Left & _
                    " Top=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").Top & _
                    " W=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").Width & _
                    " H=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").Height & _
                    " Visible=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").Visible & _
                    " Caption=" & frmEval.Controls("frHeader").Controls("cmdHdrLoadPrev").caption
    End If

    Debug.Print "=== HDR PROBE END ==="
        Exit Sub
EH:
    Debug.Print "[HDR PROBE][ERR] "; Err.Number; Err.Description
End Sub

