Attribute VB_Name = "modSchema"
Option Explicit

' ====== 蜈ｬ髢九お繝ｳ繝医Μ繝昴う繝ｳ繝・======
' dryRun:=True 縺ｧ繝ｭ繧ｰ縺ｮ縺ｿ縲・alse 縺ｧ螳滄圀縺ｫ繝ｪ繝阪・繝繝ｻ霑ｽ蜉繝ｻ荳ｦ縺ｳ譖ｿ縺医ｒ螳溯｡後・
Public Sub EnsureEvalDataSchema(Optional ByVal dryRun As Boolean = True)
    Dim ws As Worksheet
    Set ws = GetEvalDataSheet()

    Debug.Print "[SCHEMA] Start EvalData schema ensure. dryRun=" & dryRun

    ' 1) 蟋ｿ蜍｢縺ｮ讓呎ｺ門・繧ｻ繝・ヨ繧貞ｮ夂ｾｩ
    Dim desiredPosture As Collection
    Set desiredPosture = PostureDesiredHeaders()

    ' 2) 譌｢蟄倪・讓呎ｺ門錐縺ｸ縺ｮ繧ｨ繧､繝ｪ繧｢繧ｹ霎樊嶌
    Dim dictAlias As Object
    Set dictAlias = BuildPostureAliasDict()

    ' 3) 譌｢蟄伜・繧定ｵｰ譟ｻ縺励∬ｩｲ蠖薙☆繧九ｂ縺ｮ繧呈ｨ呎ｺ門錐縺ｸ謾ｹ蜷・
    ApplyHeaderAliases ws, dictAlias, dryRun

    ' 4) 谺謳榊・繧定｣懷ｮ鯉ｼ域忰蟆ｾ縺ｫ霑ｽ蜉・・
    EnsureHeaders ws, desiredPosture, dryRun
    
    Dim desiredBasic As Collection
    Set desiredBasic = BasicInfoDesiredHeaders()
    EnsureHeaders ws, desiredBasic, dryRun


    ' 5) 窶懷ｧｿ蜍｢窶昴ヶ繝ｭ繝・け蜀・・荳ｦ縺ｳ鬆・ｒ謖・ｮ夐・∈・医す繝ｼ繝亥・菴薙・鬆・ｺ上・蠕梧ｮｵ諡｡蠑ｵ・・
    ReorderPostureBlock ws, desiredPosture, dryRun

    Debug.Print "[SCHEMA] Done."
End Sub

' ====== 繧ｷ繝ｼ繝亥叙蠕・======
Public Function GetEvalDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise 5, , "EvalData 繧ｷ繝ｼ繝医′縺ゅｊ縺ｾ縺帙ｓ縲・
    Set GetEvalDataSheet = ws
End Function

' ====== 蟋ｿ蜍｢・壽ｨ呎ｺ門・螳夂ｾｩ ======
Private Function PostureDesiredHeaders() As Collection
    Dim c As New Collection

    ' 隧穂ｾ｡・医メ繧ｧ繝・け/繧ｳ繝ｳ繝・蛯呵・ｼ・
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_鬆ｭ驛ｨ蜑肴婿遯∝・"
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_蜀・レ"
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_蛛ｴ蠑ｯ"
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_菴灘ｹｹ蝗樊雷"
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_蜿榊ｼｵ閹・
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_鬪ｨ逶､蛯ｾ譁・
    c.Add "蟋ｿ蜍｢_隧穂ｾ｡_蛯呵・

    ' 諡倡ｸｮ・亥腰髢｢遽竊貞ｷｦ蜿ｳ・・
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_鬆ｸ驛ｨ"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧倬未遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧倬未遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_謇矩未遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_謇矩未遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧｡髢｢遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閧｡髢｢遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閹晞未遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_閹晞未遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_雜ｳ髢｢遽_R": c.Add "蟋ｿ蜍｢_諡倡ｸｮ_雜ｳ髢｢遽_L"
    c.Add "蟋ｿ蜍｢_諡倡ｸｮ_蛯呵・

    Set PostureDesiredHeaders = c
End Function

' ====== 繧ｨ繧､繝ｪ繧｢繧ｹ霎樊嶌讒狗ｯ会ｼ郁｡ｨ險俶昭繧娯・讓呎ｺ門錐・・======
' 縺薙％縺ｫ隕九▽縺九▲縺滓昭繧後ｒ縺ｩ繧薙←繧楢ｶｳ縺励※縺・￠縺ｰOK
Private Function BuildPostureAliasDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    ' --- 隧穂ｾ｡ ---
    d("蟋ｿ蜍｢_蜀・レ") = "蟋ｿ蜍｢_隧穂ｾ｡_蜀・レ"
    d("蜀・レ") = "蟋ｿ蜍｢_隧穂ｾ｡_蜀・レ"
    d("蟋ｿ蜍｢_鬆ｭ驛ｨ蜑肴婿遯∝・") = "蟋ｿ蜍｢_隧穂ｾ｡_鬆ｭ驛ｨ蜑肴婿遯∝・"
    d("鬆ｭ驛ｨ蜑肴婿遯∝・") = "蟋ｿ蜍｢_隧穂ｾ｡_鬆ｭ驛ｨ蜑肴婿遯∝・"
    d("蟋ｿ蜍｢_蛛ｴ蠑ｯ") = "蟋ｿ蜍｢_隧穂ｾ｡_蛛ｴ蠑ｯ"
    d("蛛ｴ蠑ｯ") = "蟋ｿ蜍｢_隧穂ｾ｡_蛛ｴ蠑ｯ"
    d("蟋ｿ蜍｢_菴灘ｹｹ蝗樊雷") = "蟋ｿ蜍｢_隧穂ｾ｡_菴灘ｹｹ蝗樊雷"
    d("菴灘ｹｹ蝗樊雷") = "蟋ｿ蜍｢_隧穂ｾ｡_菴灘ｹｹ蝗樊雷"
    d("蜿榊ｼｵ閹・) = "蟋ｿ蜍｢_隧穂ｾ｡_蜿榊ｼｵ閹・
    d("蟋ｿ蜍｢_蜿榊ｼｵ閹・) = "蟋ｿ蜍｢_隧穂ｾ｡_蜿榊ｼｵ閹・
    d("鬪ｨ逶､蛯ｾ譁・) = "蟋ｿ蜍｢_隧穂ｾ｡_鬪ｨ逶､蛯ｾ譁・
    d("蟋ｿ蜍｢_鬪ｨ逶､蛯ｾ譁・) = "蟋ｿ蜍｢_隧穂ｾ｡_鬪ｨ逶､蛯ｾ譁・

    ' 蛯呵・ｼ井ｸ頑ｮｵ・・
    d("蟋ｿ蜍｢_蛯呵・) = "蟋ｿ蜍｢_隧穂ｾ｡_蛯呵・
    d("蟋ｿ蜍｢_隧穂ｾ｡_蛯呵・ｼ井ｸ頑ｮｵ・・) = "蟋ｿ蜍｢_隧穂ｾ｡_蛯呵・
    d("蟋ｿ蜍｢隧穂ｾ｡_蛯呵・) = "蟋ｿ蜍｢_隧穂ｾ｡_蛯呵・

    ' --- 諡倡ｸｮ ---
    d("髢｢遽諡倡ｸｮ_鬆ｸ驛ｨ") = "蟋ｿ蜍｢_諡倡ｸｮ_鬆ｸ驛ｨ"
    d("諡倡ｸｮ_鬆ｸ驛ｨ") = "蟋ｿ蜍｢_諡倡ｸｮ_鬆ｸ驛ｨ"

    ' 蛛ｴ莉倥″蜷咲ｧｰ縺ｮ繧・ｌ・亥・隗偵・繧ｫ繝・さ遲会ｼ・
    d("髢｢遽諡倡ｸｮ_閧ｩ髢｢遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_R"
    d("髢｢遽諡倡ｸｮ_閧ｩ髢｢遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_L"
    d("髢｢遽諡倡ｸｮ_閧倬未遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧倬未遽_R"
    d("髢｢遽諡倡ｸｮ_閧倬未遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧倬未遽_L"
    d("髢｢遽諡倡ｸｮ_謇矩未遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_謇矩未遽_R"
    d("髢｢遽諡倡ｸｮ_謇矩未遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_謇矩未遽_L"
    d("髢｢遽諡倡ｸｮ_閧｡髢｢遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧｡髢｢遽_R"
    d("髢｢遽諡倡ｸｮ_閧｡髢｢遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閧｡髢｢遽_L"
    d("髢｢遽諡倡ｸｮ_閹晞未遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閹晞未遽_R"
    d("髢｢遽諡倡ｸｮ_閹晞未遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_閹晞未遽_L"
    d("髢｢遽諡倡ｸｮ_雜ｳ髢｢遽・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_雜ｳ髢｢遽_R"
    d("髢｢遽諡倡ｸｮ_雜ｳ髢｢遽・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_雜ｳ髢｢遽_L"

    ' 蛯呵・ｼ井ｸ区ｮｵ・・
    d("髢｢遽諡倡ｸｮ_蛯呵・) = "蟋ｿ蜍｢_諡倡ｸｮ_蛯呵・
    d("蟋ｿ蜍｢_髢｢遽諡倡ｸｮ_蛯呵・) = "蟋ｿ蜍｢_諡倡ｸｮ_蛯呵・


    ' --- 蜿ｳ/蟾ｦ 竊・R/L 螟画鋤邉ｻ・井ｸ狗ｷ壼玄蛻・ｊ・・--
    AddKoushukuSideAliases d, "閧ｩ髢｢遽"
    AddKoushukuSideAliases d, "閧倬未遽"
    AddKoushukuSideAliases d, "謇矩未遽"
    AddKoushukuSideAliases d, "閧｡髢｢遽"
    AddKoushukuSideAliases d, "閹晞未遽"
    AddKoushukuSideAliases d, "雜ｳ髢｢遽"
    
        ' --- 縲碁未遽縲阪ｒ逵√＞縺溽洒邵ｮ陦ｨ險倥・蜷ｸ蜿趣ｼ郁か/閧・謇・閧｡/閹・雜ｳ・・---
    AddKoushukuSideAliasesShort d, "閧ｩ", "閧ｩ髢｢遽"
    AddKoushukuSideAliasesShort d, "閧・, "閧倬未遽"
    AddKoushukuSideAliasesShort d, "謇・, "謇矩未遽"
    AddKoushukuSideAliasesShort d, "閧｡", "閧｡髢｢遽"
    AddKoushukuSideAliasesShort d, "閹・, "閹晞未遽"
    AddKoushukuSideAliasesShort d, "雜ｳ", "雜ｳ髢｢遽"

    
    Set BuildPostureAliasDict = d
End Function
    
    
    ' 萓具ｼ壼ｧｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_蜿ｳ 竊・蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_R
'     蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_蟾ｦ 竊・蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_L
Private Sub AddKoushukuSideAliases(ByVal d As Object, ByVal joint As String)
    d("蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_蜿ｳ") = "蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_R"
    d("蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_蟾ｦ") = "蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_L"
    ' 蠢ｵ縺ｮ縺溘ａ蜈ｨ隗偵き繝・さ迚医′谿九▲縺ｦ縺・◆蝣ｴ蜷医↓繧ょｯｾ蠢懶ｼ域里縺ｫ荳驛ｨ縺ｯ逋ｻ骭ｲ貂医∩縺縺碁㍾隍⑯K・・
    d("髢｢遽諡倡ｸｮ_" & joint & "・亥承・・) = "蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_R"
    d("髢｢遽諡倡ｸｮ_" & joint & "・亥ｷｦ・・) = "蟋ｿ蜍｢_諡倡ｸｮ_" & joint & "_L"
End Sub



' ====== 譌｢蟄倥・繝・ム縺ｫ繧ｨ繧､繝ｪ繧｢繧ｹ驕ｩ逕ｨ・域隼蜷搾ｼ・======
' ====== 譌｢蟄倥・繝・ム縺ｫ繧ｨ繧､繝ｪ繧｢繧ｹ驕ｩ逕ｨ・域隼蜷搾ｼ上・繝ｼ繧ｸ蟇ｾ蠢懶ｼ・======
Private Sub ApplyHeaderAliases(ByVal ws As Worksheet, ByVal dictAlias As Object, ByVal dryRun As Boolean)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long
    For j = lastCol To 1 Step -1      ' 蜿ｳ竊貞ｷｦ縺ｫ襍ｰ譟ｻ・壼ｾ後ｍ縺九ｉ縺ｮ譁ｹ縺悟・蜑企勁縺ｫ蠑ｷ縺・
        Dim srcHdr As String: srcHdr = Trim$(CStr(ws.Cells(1, j).value))
        If Len(srcHdr) = 0 Then GoTo ContinueLoop

        If dictAlias.exists(srcHdr) Then
            Dim dstHdr As String: dstHdr = CStr(dictAlias(srcHdr))
            Debug.Print "[SCHEMA][ALIAS] " & srcHdr & " -> " & dstHdr

            If Not dryRun Then
                Dim dstCol As Long: dstCol = FindColByHeaderExact(ws, dstHdr)
                If dstCol > 0 And dstCol <> j Then
                    ' 譌｢縺ｫ繧ｿ繝ｼ繧ｲ繝・ヨ蛻励′蟄伜惠・夂ｩｺ谺・ｒ蝓九ａ繧句ｽ｢縺ｧ繝槭・繧ｸ縺励∵立蛻励ｒ蜑企勁
                    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, j).End(xlUp).row
                    Dim r As Long
                    For r = 2 To lastRow
                        If Len(ws.Cells(r, dstCol).value) = 0 And Len(ws.Cells(r, j).value) > 0 Then
                            ws.Cells(r, dstCol).value = ws.Cells(r, j).value
                        End If
                    Next r
                    ws.Columns(j).Delete
                Else
                    ' 繧ｿ繝ｼ繧ｲ繝・ヨ蛻励′辟｡縺・ｼ壹◎縺ｮ縺ｾ縺ｾ謾ｹ蜷・
                    ws.Cells(1, j).value = dstHdr
                End If
            End If
        End If
ContinueLoop:
    Next j
End Sub

' 螳悟・荳閾ｴ縺ｧ隕句・縺怜・逡ｪ蜿ｷ繧定ｿ斐☆・育┌縺代ｌ縺ｰ0・・
Public Function FindColByHeaderExact(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindColByHeaderExact = c
            Exit Function
        End If
    Next c
    FindColByHeaderExact = 0
End Function


' ====== 谺謳阪・繝・ム縺ｮ陬懷ｮ鯉ｼ域忰蟆ｾ霑ｽ蜉・・======
Private Sub EnsureHeaders(ByVal ws As Worksheet, ByVal desired As Collection, ByVal dryRun As Boolean)
    Dim have As Object: Set have = CurrentHeaderSet(ws)
    Dim nm As Variant
    For Each nm In desired
        If Not have.exists(CStr(nm)) Then
            Debug.Print "[SCHEMA][ADD] " & CStr(nm)
            If Not dryRun Then
                Dim lastCol As Long
                lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
                ws.Cells(1, lastCol + 1).value = CStr(nm)
            End If
        End If
    Next nm
End Sub

' 迴ｾ蝨ｨ縺ｮ繝倥ャ繝髮・粋・・extCompare・・
Private Function CurrentHeaderSet(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long
    For j = 1 To lastCol
        Dim h As String: h = Trim$(CStr(ws.Cells(1, j).value))
        If Len(h) > 0 Then d(h) = j
    Next j
    Set CurrentHeaderSet = d
End Function

' ====== 蟋ｿ蜍｢繝悶Ο繝・け縺ｮ荳ｦ縺ｹ譖ｿ縺・======
' 譌｢蟄倥・ 窶懷ｧｿ蜍｢_*窶・蛻礼ｾ､繧偵‥esired縺ｮ鬆・↓蟾ｦ隧ｰ繧√〒蜀埼・鄂ｮ・井ｻ悶そ繧ｯ繧ｷ繝ｧ繝ｳ蛻励・逶ｸ蟇ｾ鬆・ｒ菫晄戟・・
Private Sub ReorderPostureBlock(ByVal ws As Worksheet, ByVal desired As Collection, ByVal dryRun As Boolean)
    Dim hdrIdx As Object: Set hdrIdx = CurrentHeaderSet(ws)

    ' 蟇ｾ雎｡蛻励・繧､繝ｳ繝・ャ繧ｯ繧ｹ蜿朱寔・亥ｭ伜惠縺吶ｋ繧ゅ・縺ｮ縺ｿ・・
    Dim targetCols As Collection: Set targetCols = New Collection
    Dim nm As Variant
    For Each nm In desired
        If hdrIdx.exists(CStr(nm)) Then
            targetCols.Add CLng(hdrIdx(CStr(nm)))
        End If
    Next nm
    If targetCols.count = 0 Then
        Debug.Print "[SCHEMA][ORDER] 蟋ｿ蜍｢_* 縺ｮ譌｢蟄伜・縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・
        Exit Sub
    End If

    ' 蟋ｿ蜍｢繝悶Ο繝・け縺ｮ迴ｾ蝨ｨ縺ｮ譛蟆上・譛螟ｧ菴咲ｽｮ
    Dim minC As Long, maxC As Long, i As Long
    minC = Columns.count: maxC = 0
    For i = 1 To targetCols.count
        minC = IIf(targetCols(i) < minC, targetCols(i), minC)
        maxC = IIf(targetCols(i) > maxC, targetCols(i), maxC)
    Next i

    ' 荳ｦ縺ｳ譖ｿ縺亥・縺ｮ髢句ｧ句・・茨ｼ晉樟繝悶Ο繝・け縺ｮ蜈磯ｭ菴咲ｽｮ・峨↓縲‥esired鬆・〒蜀埼・鄂ｮ
    ' 蠕後ｍ縺九ｉ Cut竊棚nsert 縺ｧ繧､繝ｳ繝・ャ繧ｯ繧ｹ縺壹ｌ繧貞屓驕ｿ
    Dim desiredExisting As Collection: Set desiredExisting = New Collection
    For Each nm In desired
        If hdrIdx.exists(CStr(nm)) Then desiredExisting.Add CStr(nm)
    Next nm

    Dim curPos As Long: curPos = minC
    Dim nameToCol As Object

    Set nameToCol = CurrentHeaderSet(ws) ' 譛譁ｰ蛹・
    Dim k As Long
    For k = desiredExisting.count To 1 Step -1
        Dim hName As String: hName = desiredExisting(k)
        Dim fromCol As Long: fromCol = CLng(nameToCol(hName))
        If fromCol <> curPos Then
            Debug.Print "[SCHEMA][MOVE] " & hName & "  Col " & fromCol & " -> " & curPos
            If Not dryRun Then
                ws.Columns(fromCol).Cut
                ws.Columns(curPos).Insert Shift:=xlToRight
            End If
            ' 蜀阪せ繧ｭ繝｣繝ｳ
            Set nameToCol = CurrentHeaderSet(ws)
        Else
            Debug.Print "[SCHEMA][KEEP] " & hName & " at Col " & curPos
        End If
        curPos = curPos + 1
    Next k

    Debug.Print "[SCHEMA][ORDER] 蟋ｿ蜍｢繝悶Ο繝・け荳ｦ縺ｳ譖ｿ縺亥ｮ御ｺ・・
End Sub


' 萓具ｼ壼ｧｿ蜍｢_諡倡ｸｮ_閧ｩ_蜿ｳ 竊・蟋ｿ蜍｢_諡倡ｸｮ_閧ｩ髢｢遽_R
Private Sub AddKoushukuSideAliasesShort(ByVal d As Object, ByVal shortJoint As String, ByVal fullJoint As String)
    d("蟋ｿ蜍｢_諡倡ｸｮ_" & shortJoint & "_蜿ｳ") = "蟋ｿ蜍｢_諡倡ｸｮ_" & fullJoint & "_R"
    d("蟋ｿ蜍｢_諡倡ｸｮ_" & shortJoint & "_蟾ｦ") = "蟋ｿ蜍｢_諡倡ｸｮ_" & fullJoint & "_L"
End Sub


Public Sub ListUnknownPostureHeaders()
    Dim ws As Worksheet: Set ws = GetEvalDataSheet()
    Dim desired As Collection: Set desired = PostureDesiredHeaders()
    Dim allow As Object: Set allow = CreateObject("Scripting.Dictionary")
    allow.CompareMode = 1
    Dim v
    For Each v In desired: allow(CStr(v)) = True: Next

    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long, h As String, unknown As Object: Set unknown = CreateObject("Scripting.Dictionary"): unknown.CompareMode = 1
    For j = 1 To lastCol
        h = Trim$(CStr(ws.Cells(1, j).value))
        If Len(h) > 0 Then
            If Left$(h, 3) = "蟋ｿ蜍｢_" Then
                If Not allow.exists(h) Then unknown(h) = j
            End If
        End If
    Next j

    If unknown.count = 0 Then
        Debug.Print "[SCHEMA][CHECK] 蟋ｿ蜍｢_* 縺ｮ譛ｪ遏･蛻励・縺ゅｊ縺ｾ縺帙ｓ縲・
    Else
        Dim k: For Each k In unknown.keys
            Debug.Print "[SCHEMA][CHECK][UNKNOWN] "; k; "  Col "; unknown(k)
        Next k
    End If
End Sub


Private Function BasicInfoDesiredHeaders() As Collection
    Dim c As New Collection

    c.Add "菴丞ｮ・憾豕・
    c.Add "菴丞ｮ・ｙ閠・
    c.Add "逶ｴ霑大・髯｢譌･"
    c.Add "逶ｴ霑鷹髯｢譌･"
    c.Add "豐ｻ逋らｵ碁℃"
    c.Add "蜷井ｽｵ逍ｾ謔｣繝ｻ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ"

    Set BasicInfoDesiredHeaders = c
End Function
