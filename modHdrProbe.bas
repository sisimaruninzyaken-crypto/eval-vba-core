Attribute VB_Name = "modHdrProbe"
Option Explicit

Public Sub RunHdrProbe()
    On Error GoTo EH

    Debug.Print "=== HDR PROBE START ==="

    ' frmEvalが開いてなければ開く（既に開いてればそのまま）
    If VBA.UserForms.count = 0 Then
        frmEval.Show vbModeless
        DoEvents
    End If

    Debug.Print "[Form] Controls.Count=" & frmEval.controls.count

    Debug.Print "[frHeader] Type=" & TypeName(frmEval.controls("frHeader"))
    Debug.Print "[frHeader] InsideW=" & frmEval.controls("frHeader").InsideWidth & _
                " Left=" & frmEval.controls("frHeader").Left & _
                " Top=" & frmEval.controls("frHeader").Top & _
                " W=" & frmEval.controls("frHeader").Width & _
                " H=" & frmEval.controls("frHeader").Height

    Debug.Print "[txtHdrKana] Type=" & TypeName(frmEval.controls("frHeader").controls("txtHdrKana")) & _
                " Left=" & frmEval.controls("frHeader").controls("txtHdrKana").Left & _
                " Top=" & frmEval.controls("frHeader").controls("txtHdrKana").Top & _
                " W=" & frmEval.controls("frHeader").controls("txtHdrKana").Width & _
                " H=" & frmEval.controls("frHeader").controls("txtHdrKana").Height & _
                " Visible=" & frmEval.controls("frHeader").controls("txtHdrKana").Visible

    ' 目的のボタン
    Debug.Print "[cmdHdrLoadPrev] Exists? " & (Not (frmEval.controls("frHeader").controls("cmdHdrLoadPrev") Is Nothing))
    If Not (frmEval.controls("frHeader").controls("cmdHdrLoadPrev") Is Nothing) Then
        Debug.Print "[cmdHdrLoadPrev] Left=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").Left & _
                    " Top=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").Top & _
                    " W=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").Width & _
                    " H=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").Height & _
                    " Visible=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").Visible & _
                    " Caption=" & frmEval.controls("frHeader").controls("cmdHdrLoadPrev").caption
    End If

    Debug.Print "=== HDR PROBE END ==="
        Exit Sub
EH:
    Debug.Print "[HDR PROBE][ERR] "; Err.Number; Err.Description
End Sub

