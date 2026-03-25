Attribute VB_Name = "modEvalPrintPackLatest"
Public Sub Run_PrintPack_LatestRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim lastR As Long: lastR = ws.Cells(ws.rows.count, 89).End(xlUp).row

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
    pid = CStr(ws.Cells(lastHit, 97).value) '遨ｺ縺ｮ縺ｾ縺ｾ縺ｧ繧０K

    '譌｢蟄倥・蜈･蜿｣・・nputBox迚茨ｼ峨ｒ豬∫畑縺吶ｋ縺溘ａ縲√∪縺壹・縺昴・縺ｾ縺ｾ蜻ｼ縺ｶ
    '窶ｻ谺｡縺ｮ謇九〒縲。uild_TestEval_PrintPack 繧偵悟ｼ墓焚迚医阪↓蛻・ｲ舌＆縺帙※螳悟・繝ｯ繝ｳ繧ｯ繝ｪ繝・け蛹悶☆繧・
    Call Build_TestEval_PrintPack_Forced(nm, pid)
End Sub




Public Sub Build_TestEval_PrintPack_Forced(ByVal nm As String, Optional ByVal idFilter As String = "")
    '譌｢蟄倥・ Build_TestEval_PrintPack 縺ｮ縲栗nputBox縺ｧ蜿悶▲縺ｦ繧・nm / idFilter縲阪ｒ
    '螟悶°繧画ｳｨ蜈･縺ｧ縺阪ｋ繧医≧縺ｫ縺励◆縺縺代・蜈･蜿｣縲・
    '荳ｭ霄ｫ縺ｯ譌｢蟄倥・譛ｬ菴薙ｒ蜻ｼ縺ｶ・域悽菴薙′蛻・牡縺輔ｌ縺ｦ縺・ｋ蜑肴署・峨・

    '竊薙％縺ｮ蜻ｼ縺ｳ蜃ｺ縺怜錐縺ｯ縲√≠縺ｪ縺溘・迴ｾ陦後さ繝ｼ繝峨・縲梧悽菴薙阪↓蜷医ｏ縺帙※鄂ｮ謠帙☆繧句燕謠・
    '・域ｬ｡縺ｮ謇九〒縲∵悴螳夂ｾｩ縺ｫ縺ｪ縺｣縺溯｡後□縺大ｷｮ縺玲崛縺医ｋ・・
   Call Build_TestEval_PrintPack

End Sub





