Attribute VB_Name = "FindBasicInfoCode"
' 繝励Ο繧ｸ繧ｧ繧ｯ繝亥・縺ｮ蜈ｨ繝｢繧ｸ繝･繝ｼ繝ｫ縺九ｉ繧ｭ繝ｼ繝ｯ繝ｼ繝峨ｒ讓ｪ譁ｭ讀懃ｴ｢
' 窶ｻ縲後ヵ繧｡繧､繝ｫ > 繧ｪ繝励す繝ｧ繝ｳ > 繧ｻ繧ｭ繝･繝ｪ繝・ぅ 繧ｻ繝ｳ繧ｿ繝ｼ > 繧ｻ繧ｭ繝･繝ｪ繝・ぅ 繧ｻ繝ｳ繧ｿ繝ｼ縺ｮ險ｭ螳・> 繝槭け繝ｭ縺ｮ險ｭ螳壹・
'    縺ｧ縲祁BA繝励Ο繧ｸ繧ｧ繧ｯ繝・繧ｪ繝悶ず繧ｧ繧ｯ繝医Δ繝・Ν縺ｸ縺ｮ菫｡鬆ｼ繧剃ｻ倅ｸ弱阪↓繝√ぉ繝・け縺悟ｿ・ｦ√〒縺吶・
Public Sub FindBasicInfoCode()
    Dim targets As Variant
    targets = Array("BasicInfo", "蝓ｺ譛ｬ諠・ｱ", "SaveBasicInfo", "LoadBasicInfo", "EnsureHeaderCol_BasicInfo", "隧穂ｾ｡譌･", "豌丞錐")

    On Error GoTo TrustErr
    Dim vbProj As Object, vbComp As Object, codeMod As Object
    Set vbProj = Application.VBE.ActiveVBProject

    Dim t As Variant, i As Long, lastLine As Long, lineText As String
    Debug.Print "---- BasicInfo 讀懃ｴ｢ ----"
    For Each vbComp In vbProj.VBComponents
        Set codeMod = vbComp.CodeModule
        If Not codeMod Is Nothing Then
            lastLine = codeMod.CountOfLines
            For i = 1 To lastLine
                lineText = codeMod.lines(i, 1)
                For Each t In targets
                    If InStr(1, lineText, CStr(t), vbTextCompare) > 0 Then
                        Debug.Print vbComp.name & "  L" & i & "  : " & Trim$(lineText)
                        Exit For
                    End If
                Next t
            Next i
        End If
    Next vbComp
    Debug.Print "---- 讀懃ｴ｢螳御ｺ・----"
    Exit Sub

TrustErr:
    Debug.Print "[ERROR] 繝励Ο繧ｸ繧ｧ繧ｯ繝医↓繧｢繧ｯ繧ｻ繧ｹ縺ｧ縺阪∪縺帙ｓ縲ゆｿ｡鬆ｼ險ｭ螳壹ｒ譛牙柑縺ｫ縺励※縺上□縺輔＞縲・
End Sub

