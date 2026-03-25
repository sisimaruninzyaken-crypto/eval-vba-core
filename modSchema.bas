Attribute VB_Name = "modSchema"
Option Explicit

' ====== ŒöŠJƒGƒ“ƒgƒŠƒ|ƒCƒ“ƒg ======
' dryRun:=True ‚ÅƒƒO‚Ì‚İBFalse ‚ÅÀÛ‚ÉƒŠƒl[ƒ€E’Ç‰ÁE•À‚Ñ‘Ö‚¦‚ğÀsB
Public Sub EnsureEvalDataSchema(Optional ByVal dryRun As Boolean = True)
    Dim ws As Worksheet
    Set ws = GetEvalDataSheet()

    Debug.Print "[SCHEMA] Start EvalData schema ensure. dryRun=" & dryRun

    ' 1) p¨‚Ì•W€—ñƒZƒbƒg‚ğ’è‹`
    Dim desiredPosture As Collection
    Set desiredPosture = PostureDesiredHeaders()

    ' 2) Šù‘¶¨•W€–¼‚Ö‚ÌƒGƒCƒŠƒAƒX«‘
    Dim dictAlias As Object
    Set dictAlias = BuildPostureAliasDict()

    ' 3) Šù‘¶—ñ‚ğ‘–¸‚µAŠY“–‚·‚é‚à‚Ì‚ğ•W€–¼‚Ö‰ü–¼
    ApplyHeaderAliases ws, dictAlias, dryRun

    ' 4) Œ‡‘¹—ñ‚ğ•âŠ®i––”ö‚É’Ç‰Áj
    EnsureHeaders ws, desiredPosture, dryRun
    
    Dim desiredBasic As Collection
    Set desiredBasic = BasicInfoDesiredHeaders()
    EnsureHeaders ws, desiredBasic, dryRun


    ' 5) gp¨hƒuƒƒbƒN“à‚Ì•À‚Ñ‡‚ğw’è‡‚ÖiƒV[ƒg‘S‘Ì‚Ì‡˜‚ÍŒã’iŠg’£j
    ReorderPostureBlock ws, desiredPosture, dryRun

    Debug.Print "[SCHEMA] Done."
End Sub

' ====== ƒV[ƒgæ“¾ ======
Public Function GetEvalDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise 5, , "EvalData ƒV[ƒg‚ª‚ ‚è‚Ü‚¹‚ñB"
    Set GetEvalDataSheet = ws
End Function

' ====== p¨F•W€—ñ’è‹` ======
Private Function PostureDesiredHeaders() As Collection
    Dim c As New Collection

    ' •]‰¿iƒ`ƒFƒbƒN/ƒRƒ“ƒ{/”õlj
    c.Add "p¨_•]‰¿_“ª•”‘O•û“Ëo"
    c.Add "p¨_•]‰¿_‰~”w"
    c.Add "p¨_•]‰¿_‘¤œ^"
    c.Add "p¨_•]‰¿_‘ÌŠ²‰ñù"
    c.Add "p¨_•]‰¿_”½’£•G"
    c.Add "p¨_•]‰¿_œ”ÕŒXÎ"
    c.Add "p¨_•]‰¿_”õl"

    ' Ski’PŠÖß¨¶‰Ej
    c.Add "p¨_Sk_èò•”"
    c.Add "p¨_Sk_Œ¨ŠÖß_R": c.Add "p¨_Sk_Œ¨ŠÖß_L"
    c.Add "p¨_Sk_•IŠÖß_R": c.Add "p¨_Sk_•IŠÖß_L"
    c.Add "p¨_Sk_èŠÖß_R": c.Add "p¨_Sk_èŠÖß_L"
    c.Add "p¨_Sk_ŒÒŠÖß_R": c.Add "p¨_Sk_ŒÒŠÖß_L"
    c.Add "p¨_Sk_•GŠÖß_R": c.Add "p¨_Sk_•GŠÖß_L"
    c.Add "p¨_Sk_‘«ŠÖß_R": c.Add "p¨_Sk_‘«ŠÖß_L"
    c.Add "p¨_Sk_”õl"

    Set PostureDesiredHeaders = c
End Function

' ====== ƒGƒCƒŠƒAƒX«‘\’zi•\‹L—h‚ê¨•W€–¼j ======
' ‚±‚±‚ÉŒ©‚Â‚©‚Á‚½—h‚ê‚ğ‚Ç‚ñ‚Ç‚ñ‘«‚µ‚Ä‚¢‚¯‚ÎOK
Private Function BuildPostureAliasDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    ' --- •]‰¿ ---
    d("p¨_‰~”w") = "p¨_•]‰¿_‰~”w"
    d("‰~”w") = "p¨_•]‰¿_‰~”w"
    d("p¨_“ª•”‘O•û“Ëo") = "p¨_•]‰¿_“ª•”‘O•û“Ëo"
    d("“ª•”‘O•û“Ëo") = "p¨_•]‰¿_“ª•”‘O•û“Ëo"
    d("p¨_‘¤œ^") = "p¨_•]‰¿_‘¤œ^"
    d("‘¤œ^") = "p¨_•]‰¿_‘¤œ^"
    d("p¨_‘ÌŠ²‰ñù") = "p¨_•]‰¿_‘ÌŠ²‰ñù"
    d("‘ÌŠ²‰ñù") = "p¨_•]‰¿_‘ÌŠ²‰ñù"
    d("”½’£•G") = "p¨_•]‰¿_”½’£•G"
    d("p¨_”½’£•G") = "p¨_•]‰¿_”½’£•G"
    d("œ”ÕŒXÎ") = "p¨_•]‰¿_œ”ÕŒXÎ"
    d("p¨_œ”ÕŒXÎ") = "p¨_•]‰¿_œ”ÕŒXÎ"

    ' ”õliã’ij
    d("p¨_”õl") = "p¨_•]‰¿_”õl"
    d("p¨_•]‰¿_”õliã’ij") = "p¨_•]‰¿_”õl"
    d("p¨•]‰¿_”õl") = "p¨_•]‰¿_”õl"

    ' --- Sk ---
    d("ŠÖßSk_èò•”") = "p¨_Sk_èò•”"
    d("Sk_èò•”") = "p¨_Sk_èò•”"

    ' ‘¤•t‚«–¼Ì‚Ì‚ä‚êi‘SŠpEƒJƒbƒR“™j
    d("ŠÖßSk_Œ¨ŠÖßi‰Ej") = "p¨_Sk_Œ¨ŠÖß_R"
    d("ŠÖßSk_Œ¨ŠÖßi¶j") = "p¨_Sk_Œ¨ŠÖß_L"
    d("ŠÖßSk_•IŠÖßi‰Ej") = "p¨_Sk_•IŠÖß_R"
    d("ŠÖßSk_•IŠÖßi¶j") = "p¨_Sk_•IŠÖß_L"
    d("ŠÖßSk_èŠÖßi‰Ej") = "p¨_Sk_èŠÖß_R"
    d("ŠÖßSk_èŠÖßi¶j") = "p¨_Sk_èŠÖß_L"
    d("ŠÖßSk_ŒÒŠÖßi‰Ej") = "p¨_Sk_ŒÒŠÖß_R"
    d("ŠÖßSk_ŒÒŠÖßi¶j") = "p¨_Sk_ŒÒŠÖß_L"
    d("ŠÖßSk_•GŠÖßi‰Ej") = "p¨_Sk_•GŠÖß_R"
    d("ŠÖßSk_•GŠÖßi¶j") = "p¨_Sk_•GŠÖß_L"
    d("ŠÖßSk_‘«ŠÖßi‰Ej") = "p¨_Sk_‘«ŠÖß_R"
    d("ŠÖßSk_‘«ŠÖßi¶j") = "p¨_Sk_‘«ŠÖß_L"

    ' ”õli‰º’ij
    d("ŠÖßSk_”õl") = "p¨_Sk_”õl"
    d("p¨_ŠÖßSk_”õl") = "p¨_Sk_”õl"


    ' --- ‰E/¶ ¨ R/L •ÏŠ·Œni‰ºü‹æØ‚èj---
    AddKoushukuSideAliases d, "Œ¨ŠÖß"
    AddKoushukuSideAliases d, "•IŠÖß"
    AddKoushukuSideAliases d, "èŠÖß"
    AddKoushukuSideAliases d, "ŒÒŠÖß"
    AddKoushukuSideAliases d, "•GŠÖß"
    AddKoushukuSideAliases d, "‘«ŠÖß"
    
        ' --- uŠÖßv‚ğÈ‚¢‚½’Zk•\‹L‚Ì‹zûiŒ¨/•I/è/ŒÒ/•G/‘«j ---
    AddKoushukuSideAliasesShort d, "Œ¨", "Œ¨ŠÖß"
    AddKoushukuSideAliasesShort d, "•I", "•IŠÖß"
    AddKoushukuSideAliasesShort d, "è", "èŠÖß"
    AddKoushukuSideAliasesShort d, "ŒÒ", "ŒÒŠÖß"
    AddKoushukuSideAliasesShort d, "•G", "•GŠÖß"
    AddKoushukuSideAliasesShort d, "‘«", "‘«ŠÖß"

    
    Set BuildPostureAliasDict = d
End Function
    
    
    ' —áFp¨_Sk_Œ¨ŠÖß_‰E ¨ p¨_Sk_Œ¨ŠÖß_R
'     p¨_Sk_Œ¨ŠÖß_¶ ¨ p¨_Sk_Œ¨ŠÖß_L
Private Sub AddKoushukuSideAliases(ByVal d As Object, ByVal joint As String)
    d("p¨_Sk_" & joint & "_‰E") = "p¨_Sk_" & joint & "_R"
    d("p¨_Sk_" & joint & "_¶") = "p¨_Sk_" & joint & "_L"
    ' ”O‚Ì‚½‚ß‘SŠpƒJƒbƒR”Å‚ªc‚Á‚Ä‚¢‚½ê‡‚É‚à‘Î‰iŠù‚Éˆê•”‚Í“o˜^Ï‚İ‚¾‚ªd•¡OKj
    d("ŠÖßSk_" & joint & "i‰Ej") = "p¨_Sk_" & joint & "_R"
    d("ŠÖßSk_" & joint & "i¶j") = "p¨_Sk_" & joint & "_L"
End Sub



' ====== Šù‘¶ƒwƒbƒ_‚ÉƒGƒCƒŠƒAƒX“K—pi‰ü–¼j ======
' ====== Šù‘¶ƒwƒbƒ_‚ÉƒGƒCƒŠƒAƒX“K—pi‰ü–¼^ƒ}[ƒW‘Î‰j ======
Private Sub ApplyHeaderAliases(ByVal ws As Worksheet, ByVal dictAlias As Object, ByVal dryRun As Boolean)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long
    For j = lastCol To 1 Step -1      ' ‰E¨¶‚É‘–¸FŒã‚ë‚©‚ç‚Ì•û‚ª—ñíœ‚É‹­‚¢
        Dim srcHdr As String: srcHdr = Trim$(CStr(ws.Cells(1, j).value))
        If Len(srcHdr) = 0 Then GoTo ContinueLoop

        If dictAlias.exists(srcHdr) Then
            Dim dstHdr As String: dstHdr = CStr(dictAlias(srcHdr))
            Debug.Print "[SCHEMA][ALIAS] " & srcHdr & " -> " & dstHdr

            If Not dryRun Then
                Dim dstCol As Long: dstCol = FindColByHeaderExact(ws, dstHdr)
                If dstCol > 0 And dstCol <> j Then
                    ' Šù‚Éƒ^[ƒQƒbƒg—ñ‚ª‘¶İF‹ó—“‚ğ–„‚ß‚éŒ`‚Åƒ}[ƒW‚µA‹Œ—ñ‚ğíœ
                    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, j).End(xlUp).row
                    Dim r As Long
                    For r = 2 To lastRow
                        If Len(ws.Cells(r, dstCol).value) = 0 And Len(ws.Cells(r, j).value) > 0 Then
                            ws.Cells(r, dstCol).value = ws.Cells(r, j).value
                        End If
                    Next r
                    ws.Columns(j).Delete
                Else
                    ' ƒ^[ƒQƒbƒg—ñ‚ª–³‚¢F‚»‚Ì‚Ü‚Ü‰ü–¼
                    ws.Cells(1, j).value = dstHdr
                End If
            End If
        End If
ContinueLoop:
    Next j
End Sub

' Š®‘Sˆê’v‚ÅŒ©o‚µ—ñ”Ô†‚ğ•Ô‚·i–³‚¯‚ê‚Î0j
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


' ====== Œ‡‘¹ƒwƒbƒ_‚Ì•âŠ®i––”ö’Ç‰Áj ======
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

' Œ»İ‚Ìƒwƒbƒ_W‡iTextComparej
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

' ====== p¨ƒuƒƒbƒN‚Ì•À‚×‘Ö‚¦ ======
' Šù‘¶‚Ì gp¨_*h —ñŒQ‚ğAdesired‚Ì‡‚É¶‹l‚ß‚ÅÄ”z’ui‘¼ƒZƒNƒVƒ‡ƒ“—ñ‚Í‘Š‘Î‡‚ğ•Ûj
Private Sub ReorderPostureBlock(ByVal ws As Worksheet, ByVal desired As Collection, ByVal dryRun As Boolean)
    Dim hdrIdx As Object: Set hdrIdx = CurrentHeaderSet(ws)

    ' ‘ÎÛ—ñ‚ÌƒCƒ“ƒfƒbƒNƒXûWi‘¶İ‚·‚é‚à‚Ì‚Ì‚İj
    Dim targetCols As Collection: Set targetCols = New Collection
    Dim nm As Variant
    For Each nm In desired
        If hdrIdx.exists(CStr(nm)) Then
            targetCols.Add CLng(hdrIdx(CStr(nm)))
        End If
    Next nm
    If targetCols.count = 0 Then
        Debug.Print "[SCHEMA][ORDER] p¨_* ‚ÌŠù‘¶—ñ‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñB"
        Exit Sub
    End If

    ' p¨ƒuƒƒbƒN‚ÌŒ»İ‚ÌÅ¬EÅ‘åˆÊ’u
    Dim minC As Long, maxC As Long, i As Long
    minC = Columns.count: maxC = 0
    For i = 1 To targetCols.count
        minC = IIf(targetCols(i) < minC, targetCols(i), minC)
        maxC = IIf(targetCols(i) > maxC, targetCols(i), maxC)
    Next i

    ' •À‚Ñ‘Ö‚¦æ‚ÌŠJn—ñiŒ»ƒuƒƒbƒN‚Ìæ“ªˆÊ’uj‚ÉAdesired‡‚ÅÄ”z’u
    ' Œã‚ë‚©‚ç Cut¨Insert ‚ÅƒCƒ“ƒfƒbƒNƒX‚¸‚ê‚ğ‰ñ”ğ
    Dim desiredExisting As Collection: Set desiredExisting = New Collection
    For Each nm In desired
        If hdrIdx.exists(CStr(nm)) Then desiredExisting.Add CStr(nm)
    Next nm

    Dim curPos As Long: curPos = minC
    Dim nameToCol As Object

    Set nameToCol = CurrentHeaderSet(ws) ' ÅV‰»
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
            ' ÄƒXƒLƒƒƒ“
            Set nameToCol = CurrentHeaderSet(ws)
        Else
            Debug.Print "[SCHEMA][KEEP] " & hName & " at Col " & curPos
        End If
        curPos = curPos + 1
    Next k

    Debug.Print "[SCHEMA][ORDER] p¨ƒuƒƒbƒN•À‚Ñ‘Ö‚¦Š®—¹B"
End Sub


' —áFp¨_Sk_Œ¨_‰E ¨ p¨_Sk_Œ¨ŠÖß_R
Private Sub AddKoushukuSideAliasesShort(ByVal d As Object, ByVal shortJoint As String, ByVal fullJoint As String)
    d("p¨_Sk_" & shortJoint & "_‰E") = "p¨_Sk_" & fullJoint & "_R"
    d("p¨_Sk_" & shortJoint & "_¶") = "p¨_Sk_" & fullJoint & "_L"
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
            If Left$(h, 3) = "p¨_" Then
                If Not allow.exists(h) Then unknown(h) = j
            End If
        End If
    Next j

    If unknown.count = 0 Then
        Debug.Print "[SCHEMA][CHECK] p¨_* ‚Ì–¢’m—ñ‚Í‚ ‚è‚Ü‚¹‚ñB"
    Else
        Dim k: For Each k In unknown.keys
            Debug.Print "[SCHEMA][CHECK][UNKNOWN] "; k; "  Col "; unknown(k)
        Next k
    End If
End Sub


Private Function BasicInfoDesiredHeaders() As Collection
    Dim c As New Collection

    c.Add "Z‘îó‹µ"
    c.Add "Z‘î”õl"
    c.Add "’¼‹ß“ü‰@“ú"
    c.Add "’¼‹ß‘Ş‰@“ú"
    c.Add "¡—ÃŒo‰ß"
    c.Add "‡•¹¾Š³EƒRƒ“ƒgƒ[ƒ‹"

    Set BasicInfoDesiredHeaders = c
End Function
