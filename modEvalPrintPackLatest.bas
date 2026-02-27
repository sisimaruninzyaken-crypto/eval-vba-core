Attribute VB_Name = "modEvalPrintPackLatest"
Public Sub Run_PrintPack_LatestRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, 89).End(xlUp).row

    Dim r As Long, d As Variant, latest As Date, lastHit As Long
    latest = 0
    For r = 2 To lastR
        d = ws.Cells(r, 86).value
        If IsDate(d) Then
            If DateValue(d) > latest Then latest = DateValue(d)
        End If
    Next
    For r = 2 To lastR
        d = ws.Cells(r, 86).value
        If IsDate(d) Then
            If DateValue(d) = latest Then lastHit = r
        End If
    Next

    Dim nm As String, pid As String
    nm = CStr(ws.Cells(lastHit, 89).value)
    pid = CStr(ws.Cells(lastHit, 97).value) '空のままでもOK

    '既存の入口（InputBox版）を流用するため、まずはそのまま呼ぶ
    '※次の手で、Build_TestEval_PrintPack を「引数版」に分岐させて完全ワンクリック化する
    Call Build_TestEval_PrintPack_Forced(nm, pid)
End Sub




Public Sub Build_TestEval_PrintPack_Forced(ByVal nm As String, Optional ByVal idFilter As String = "")
    '既存の Build_TestEval_PrintPack の「InputBoxで取ってる nm / idFilter」を
    '外から注入できるようにしただけの入口。
    '中身は既存の本体を呼ぶ（本体が分割されている前提）。

    '↓この呼び出し名は、あなたの現行コードの「本体」に合わせて置換する前提
    '（次の手で、未定義になった行だけ差し替える）
   Call Build_TestEval_PrintPack

End Sub





