Attribute VB_Name = "MMT"

Option Explicit

' === 生成系（Direct） ===


Public Sub MMT_BuildChildTabs_Direct()
  
    Dim pg As Object
    Set pg = GetMMTPage(frmEval)
    If pg Is Nothing Then
        MsgBox "MMTページが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim host As Object, mpMMTChildGen As Object
    Dim pgUpper As Object, pgLower As Object
    
    ClearLegacyMMTOnPage14 frmEval
    


    '--- host確保（無ければ作る） ---
    Set host = GetMMTHost(pg)
    
    If host Is Nothing Then
        MsgBox "MMTホストが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' hostサイズは毎回追従（pgが小さくても破綻しない）
    host.Width = pg.InsideWidth - 12
    host.Height = pg.InsideHeight - 12
    
    '--- mp確保（host配下） ---
    Set mpMMTChildGen = GetMMTChildTabs(pg, host)
    
    If mpMMTChildGen Is Nothing Then

        MsgBox "子タブ(mpMMTChild)が作成できません。", vbExclamation
        Exit Sub
    End If

    ' mpサイズも毎回追従
    mpMMTChildGen.Width = host.InsideWidth
    mpMMTChildGen.Height = host.InsideHeight
    
    ' legacy stray cleanup
    PurgeStrayMMTControls pg, host

    '--- 子タブの中身を作り直す（MMTGENだけ消す） ---
    Set pgUpper = mpMMTChildGen.Pages(0)
    Set pgLower = mpMMTChildGen.Pages(1)
MMT_ClearGen pgUpper
MMT_ClearGen pgLower

BuildMMTPage pgUpper, Array("肩屈曲", "肩伸展", "肩外転", "肩内旋", _
                            "肩外旋", "肘屈曲", "肘伸展", _
                            "前腕回内", "前腕回外", _
                            "手関節掌屈", "手関節背屈", _
                            "指屈曲", "指伸展", "母指対立")

BuildMMTPage pgLower, Array("股屈曲", "股伸展", "股外転", "股内転", _
                            "膝屈曲", "膝伸展", _
                            "足関節背屈", "足関節底屈", _
                            "母趾伸展")
    DoEvents
    Resize_MMTChildHost_ToPage
    
    Exit Sub


RRTRACE:
    Debug.Print "[MMT ERROR]", Err.Number, Err.Description

End Sub

Private Sub ClearLegacyMMTOnPage14(ByVal frm As Object)
    Dim mp1 As Object
    Dim pgRoot As Object
    Dim host As Object
    Dim mpPhys As Object
    Dim i As Long
    Dim pg As Object

    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    Set mp1 = frm.controls("MultiPage1")
    If mp1 Is Nothing Then Exit Sub

    Set pgRoot = mp1.Pages(2)
    If pgRoot Is Nothing Then Exit Sub

    Set host = pgRoot.controls("Frame3")
    If host Is Nothing Then Exit Sub

    Set mpPhys = host.controls("mpPhys")
    If mpPhys Is Nothing Then Exit Sub
    On Error GoTo 0

    For i = 0 To mpPhys.Pages.count - 1
        Set pg = mpPhys.Pages(i)
        If LCase$(CStr(pg.name)) = "page14" Then
            RemoveLegacyMMTControlsFromPage pg
            Exit For
        End If
    Next i
End Sub

Private Sub RemoveLegacyMMTControlsFromPage(ByVal pg As Object)
    Dim i As Long
    Dim ctl As Object
    Dim nm As String
    Dim removedCount As Long
    Dim hiddenCount As Long
    Dim failedCount As Long
    Dim removeErrNo As Long
    Dim removeErrDesc As String
    
    If pg Is Nothing Then Exit Sub

    For i = pg.controls.count - 1 To 0 Step -1
    
        Set ctl = pg.controls(i)
        nm = CStr(ctl.name)
    

        If IsLegacyMMTControlName(nm) Then
            Err.Clear
            On Error GoTo REMOVE_FAILED
            pg.controls.Remove nm
            removedCount = removedCount + 1
#If APP_DEBUG Then
            Debug.Print "[MMT][LEGACY][REMOVE]", nm
#End If
            On Error GoTo 0
            GoTo NEXT_CONTROL

REMOVE_FAILED:
            removeErrNo = Err.Number
            removeErrDesc = Err.Description
            Err.Clear

            On Error GoTo HIDE_FAILED
            ctl.Visible = False
            ctl.Enabled = False
            hiddenCount = hiddenCount + 1
#If APP_DEBUG Then
            Debug.Print "[MMT][LEGACY][HIDE]", nm, "removeErr=" & removeErrNo, removeErrDesc
#End If
            On Error GoTo 0
            GoTo NEXT_CONTROL

HIDE_FAILED:
            failedCount = failedCount + 1
#If APP_DEBUG Then
            Debug.Print "[MMT][LEGACY][FAIL]", nm, _
                        " removeErr=" & removeErrNo & ":" & removeErrDesc, _
                        " hideErr=" & Err.Number & ":" & Err.Description
#End If
            Err.Clear
            On Error GoTo 0
        End If

NEXT_CONTROL:
        On Error GoTo 0
    Next i

#If APP_DEBUG Then
    Debug.Print "[MMT][LEGACY][SUMMARY] removed=" & removedCount & " hidden=" & hiddenCount & " failed=" & failedCount
#End If
End Sub

Private Function IsLegacyMMTControlName(ByVal nm As String) As Boolean
    Dim key As String

    If Len(nm) = 0 Then Exit Function

    If nm = "lblHdrMus" Or nm = "lblHdrR" Or nm = "lblHdrL" Then
        IsLegacyMMTControlName = True
        Exit Function
    End If

    key = GetLegacyMMTKeyByControlName(nm)
    If Len(key) > 0 Then IsLegacyMMTControlName = True
End Function

Private Function GetLegacyMMTKeyByControlName(ByVal nm As String) As String
    Dim keys As Variant
    Dim i As Long
    Dim key As String

    keys = LegacyMMTKeys()
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If nm = "lbl_" & key _
           Or nm = "cboR_" & key _
           Or nm = "cboL_" & key Then
            GetLegacyMMTKeyByControlName = key
            Exit Function
        End If
    Next i
End Function

Private Function LegacyMMTKeys() As Variant
    LegacyMMTKeys = Array("肩屈曲", "肩伸展", "肩外転", "肩内旋", "肩外旋", _
                      "肘屈曲", "肘伸展", "前腕回内", "前腕回外", _
                      "手関節掌屈", "手関節背屈", "指屈曲", "指伸展", "母指対立", _
                      "股屈曲", "股伸展", "股外転", "股内転", _
                      "膝屈曲", "膝伸展", "足関節背屈", "足関節底屈", "母趾伸展")
End Function



Public Function UseMMTChildTabs() As Boolean
    UseMMTChildTabs = True
End Function

Public Function GetMMTHost(ByVal pg As Object) As Object
    Dim host As Object
    Dim cand As Variant
    Dim i As Long
    Dim j As Long
    Dim c As Object
    Dim mpProbe As Object
    
    If pg Is Nothing Then Exit Function
    

    ' mpPhys-host first
    On Error Resume Next
    If TypeName(pg.parent) = "MultiPage" Then
        If LCase$(CStr(pg.parent.name)) = "mpphys" Then
            For i = 0 To pg.controls.count - 1
                Set c = pg.controls(i)
                If TypeName(c) = "Frame" Then
                    Set GetMMTHost = c
                    Exit Function
                End If
            Next i
        End If
    End If
    On Error GoTo 0
    
    For i = 0 To pg.controls.count - 1
        Set c = pg.controls(i)
        If TypeName(c) = "Frame" Then
            Set mpProbe = Nothing
            On Error Resume Next
            For j = 0 To c.controls.count - 1
                If TypeName(c.controls(j)) = "MultiPage" Then
                    Set mpProbe = c.controls(j)
                    Exit For
                End If
            Next j
            On Error GoTo 0

            If Not mpProbe Is Nothing Then
                If InStr(1, CStr(c.name), "MMT", vbTextCompare) > 0 _
                   Or InStr(1, CStr(mpProbe.name), "MMT", vbTextCompare) > 0 _
                   Or InStr(1, CStr(mpProbe.tag), "MMT", vbTextCompare) > 0 Then
                    Set GetMMTHost = c
                    Exit Function
                End If
            End If
        End If
    Next i
    
    
    ' 1) 候補名を優先
    For Each cand In Array("Frame9", "fraMMTWrap")
        On Error Resume Next
        Set host = SafeGetControl(pg, CStr(cand))
        On Error GoTo 0
        
        If Not host Is Nothing Then
            If TypeName(host) = "Frame" Then

                Set GetMMTHost = host
                Exit Function
            End If
            Set host = Nothing
        End If
    Next cand
    
#If APP_DEBUG Then
    Debug.Print "[MMT][HOST] not found -> skip"
#End If
End Function

Public Function GetMMTChildTabs(ByVal pg As Object, Optional ByVal host As Object = Nothing) As Object
    Dim mp As Object
    Dim i As Long
    
    If host Is Nothing Then Set host = GetMMTHost(pg)
    If host Is Nothing Then Exit Function
    
    Set mp = Nothing
    If TypeName(host) = "MultiPage" Then
        If LCase$(CStr(host.name)) = "mpmmtchildgen" _
           Or InStr(1, CStr(host.tag), "MMTGEN", vbTextCompare) > 0 Then
            Set mp = host
        End If
    Else
        On Error Resume Next
        For i = 0 To host.controls.count - 1
            If TypeName(host.controls(i)) = "MultiPage" Then
                If LCase$(CStr(host.controls(i).name)) = "mpmmtchildgen" _
                   Or InStr(1, CStr(host.controls(i).tag), "MMTGEN", vbTextCompare) > 0 Then
                    Set mp = host.controls(i)
                    Exit For
                End If
            End If
        Next i
        On Error GoTo 0
    End If
    

        If mp Is Nothing Then
            Set mp = host.controls.Add("Forms.MultiPage.1", "mpMMTChildGen", True)
            With mp
                .Left = 0
                .Top = 0
                .Width = host.InsideWidth
                .Height = host.InsideHeight
                .Style = 0

                .tag = "MMTGEN"
            End With
        End If
    
    
    If mp.Pages.count < 2 Then
        Do While mp.Pages.count < 2
            mp.Pages.Add
        Loop
            ElseIf mp.Pages.count > 2 Then
        Do While mp.Pages.count > 2
            mp.Pages.Remove mp.Pages.count - 1
        Loop
    End If
    
    mp.Pages(0).caption = ChrW(&H4E0A) & ChrW(&H80A2)
    mp.Pages(1).caption = ChrW(&H4E0B) & ChrW(&H80A2)
    
    Set GetMMTChildTabs = mp
End Function


Public Function GetMMTPage(ByVal frm As Object) As Object
    Dim ctl As Object, pg As Object

    If frm Is Nothing Then Exit Function

    ' フォーム直下の MultiPage を総なめ
    ' strict route first
    Set pg = GetMMTPage_FromPhys(frm)
    If Not pg Is Nothing Then
        Set GetMMTPage = pg
        Exit Function
    End If

    For Each ctl In frm.controls
        If TypeName(ctl) = "MultiPage" Then
            Dim i As Long
            For i = 0 To ctl.Pages.count - 1
                Set pg = ctl.Pages(i)
                If PageHasMMTSignature(pg) Then
                    Set GetMMTPage = pg
                    Exit Function
                End If
            Next i
        End If
    Next ctl
End Function


Private Function GetMMTPage_FromPhys(ByVal frm As Object) As Object
    Dim mp1 As Object, pgPhysRoot As Object, host As Object, mpPhys As Object
    Dim i As Long, cap As String

    On Error Resume Next
    Set mp1 = frm.controls("MultiPage1")
    If mp1 Is Nothing Then Exit Function

    Set pgPhysRoot = mp1.Pages(2)
    If pgPhysRoot Is Nothing Then Exit Function

    Set host = pgPhysRoot.controls("Frame3")
    If host Is Nothing Then Exit Function

    Set mpPhys = host.controls("mpPhys")
    If mpPhys Is Nothing Then Exit Function

    For i = 0 To mpPhys.Pages.count - 1
        cap = CStr(mpPhys.Pages(i).caption)
        If InStr(1, cap, "MMT", vbTextCompare) > 0 Then
            Set GetMMTPage_FromPhys = mpPhys.Pages(i)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function


Private Function PageHasMMTSignature(ByVal pg As Object) As Boolean
    Dim c As Object

    If pg Is Nothing Then Exit Function

    ' 「mpMMTChild」や「Frame9」など、MMTページ固有の痕跡で判定
    For Each c In pg.controls
        If LCase$(c.name) = "mpmmtchild" Then
            PageHasMMTSignature = True
            Exit Function
        End If
        If LCase$(c.name) = "frame9" Then
            PageHasMMTSignature = True
            Exit Function
        End If
        If InStr(1, c.name, "MMT", vbTextCompare) > 0 Then
            PageHasMMTSignature = True
            Exit Function
        End If
    Next c
End Function


Private Sub MakeCbo(ByVal pg As Object, ByVal nm As String, _
                    ByVal L As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.controls.Add("Forms.ComboBox.1", nm, True)
    o.Left = L: o.Top = t: o.Width = w: o.Height = h
    o.Style = MSForms.fmStyleDropDownList: o.BoundColumn = 1
    o.List = Split("0,1,2,3,4,5", ","): o.tag = "MMTGEN"
End Sub


Private Sub MakeLbl(ByVal pg As Object, ByVal nm As String, ByVal cap As String, _
                    ByVal L As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.label
    Set o = pg.controls.Add("Forms.Label.1", nm, True)
    o.caption = cap: o.Left = L: o.Top = t: o.Width = w: o.Height = h: o.tag = "MMTGEN"
End Sub



'--- 1ページ分のUI生成 ---
Private Sub BuildMMTPage(ByVal pg As Object, ByVal items As Variant)
    Const ROW_H As Single = 24, LBL_W As Single = 130, COL_W As Single = 90, gap As Single = 12
    Dim x0 As Single, y0 As Single: x0 = 20: y0 = 28

    MakeLbl pg, "lblHdrMus", "筋群", x0, y0 - 20, 60, 18
    MakeLbl pg, "lblHdrR", "右", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLbl pg, "lblHdrL", "左", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, y As Single: y = y0
    For i = LBound(items) To UBound(items)
        Dim key As String: key = CStr(items(i))
        MakeLbl pg, "lbl_" & key, key, x0, y + 3, LBL_W, 18
        MakeCbo pg, "cboR_" & key, x0 + LBL_W + gap, y, COL_W, 18
        MakeCbo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, y, COL_W, 18
        y = y + ROW_H
    Next
End Sub

Private Sub PurgeStrayMMTControls(ByVal pg As Object, ByVal host As Object)
    PurgeMMTNamedControlsInContainer pg
    If Not host Is Nothing Then PurgeMMTNamedControlsInContainer host
End Sub

Private Sub PurgeMMTNamedControlsInContainer(ByVal parent As Object)
    Dim i As Long
    Dim nm As String

    If parent Is Nothing Then Exit Sub

    On Error Resume Next
    For i = parent.controls.count - 1 To 0 Step -1
        nm = LCase$(CStr(parent.controls(i).name))
        If Left$(nm, 5) = "cbor_" _
           Or Left$(nm, 5) = "cbol_" _
           Or Left$(nm, 4) = "lbl_" Then
            parent.controls.Remove parent.controls(i).name
        End If
    Next i
    On Error GoTo 0
End Sub


Private Sub MMT_ClearGen(ByVal pg As Object)
    Dim idx As Long
    For idx = pg.controls.count - 1 To 0 Step -1
        If Left$(pg.controls(idx).tag & "", 6) = "MMTGEN" Then
            pg.controls.Remove pg.controls(idx).name
        End If
    Next
End Sub

'--- 子タブ内の全 ComboBox をいったんクリア ---
Private Sub MMT_ClearMMTCombos(ByVal mp As MSForms.MultiPage)
    Dim pg As MSForms.page
    Dim c As Object  '（ControlでもOK）

    For Each pg In mp.Pages
        For Each c In pg.controls
            If TypeName(c) = "ComboBox" Then
                On Error Resume Next
                c.ListIndex = -1   '選択解除（DropDownListでも有効）
                c.value = ""       '念のため空文字
                On Error GoTo 0
            End If
        Next
    Next
End Sub


' === 保存（行ベース） ===
Public Sub SaveMMTToSheet(ws As Worksheet, r As Long, owner As Object)
    Dim c As Long
    Dim s As String

    ' 1) ヘッダー列（MMT_IO）を用意
    c = FindColByHeaderExact(ws, "MMT_IO")
    If c = 0 Then
        c = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, c).value = "MMT_IO"
    End If

    ' 2) MMTを文字列化して保存
    s = MMT_SaveToString()              ' ←既にある関数（直下に見えているもの）を使う
    ws.Cells(r, c).value = s

    ' 3) ログ
    Debug.Print "[MMT][SAVE] row=" & r & " col=" & c & " len=" & Len(s)
End Sub

'=== 子タブ(MMT) → 文字列（保存用テスト） ===
Private Function MMT_SaveToString() As String
    Dim pg As Object, mp As Object, p As Long, c As Object
    Dim parts() As String, n As Long
    Set pg = GetMMTPage(frmEval)
    If pg Is Nothing Then Exit Function

    On Error Resume Next
    Set mp = GetMMTChildTabs(pg)
    On Error GoTo 0
    If mp Is Nothing Then Exit Function

    ReDim parts(0 To 0): n = -1

    For p = 0 To mp.Pages.count - 1
        For Each c In mp.Pages(p).controls
            If TypeName(c) = "ComboBox" Then
                Dim nm As String, side As String
                If Left$(c.name, 5) = "cboR_" Then
                    nm = Mid$(c.name, 6): side = "R"
                ElseIf Left$(c.name, 5) = "cboL_" Then
                    nm = Mid$(c.name, 6): side = "L"
                Else
                    nm = "": side = ""
                End If

                If nm <> "" And side = "R" Then
                    Dim rVal As String, lVal As String
                    rVal = CStr(c.value)
                    On Error Resume Next
                    lVal = CStr(mp.Pages(p).controls("cboL_" & nm).value)
                    On Error GoTo 0

                    n = n + 1
                    ReDim Preserve parts(0 To n)
                    parts(n) = p & "|" & nm & "|" & rVal & "|" & lVal
                End If
            End If
        Next c
    Next p

    If n >= 0 Then MMT_SaveToString = Join(parts, ";")
End Function


' === 読込（行ベース） ===
Public Sub LoadMMTFromSheet(ws As Worksheet, r As Long, owner As Object)

    

    Dim c As Long, s As String
    Dim pg As Object, mp As Object

    ' MMT_IO 列
    c = FindColByHeaderExact(ws, "MMT_IO")
    Debug.Print "[MMT] col=" & c
    If c = 0 Then Exit Sub

    s = ReadStr_Compat("IO_MMT", r, ws)
    Debug.Print "[MMT] s.len=" & Len(s)
    If Len(s) = 0 Then Exit Sub

    ' MMTページ
    Debug.Print "[MMT] before GetMMTPage"
    Set pg = GetMMTPage(owner)
    Debug.Print "[MMT] after GetMMTPage", TypeName(pg)

    If pg Is Nothing Then Exit Sub

    ' 子タブの存在確認
    Debug.Print "[MMT] before GetMMTChildTabs"
    Set mp = GetMMTChildTabs(pg)
    Debug.Print "[MMT] after GetMMTChildTabs", TypeName(mp)

    If mp Is Nothing Then Exit Sub

    Debug.Print "[MMT] before LoadFromString"
    MMT_LoadFromString_Core s
    Debug.Print "[MMT] after LoadFromString"

End Sub
'========================
' 直列化文字列 → 子タブへ復元
' 形式：  side|項目名|右値|左値 ; side|項目名|右値|左値 ; ...
'   side: 0=上肢, 1=下肢
' 右値/左値は空文字もあり得る
'========================
Private Sub MMT_LoadFromString_Core(ByVal s As String)
    Dim pg As Object, mp As MSForms.MultiPage
    Dim parts As Variant, itm As Variant
    Dim side As Long, key As String
    Dim vR As String, vL As String
    Dim p As MSForms.page
    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox

    Set pg = GetMMTPage(frmEval)
    If pg Is Nothing Then
        MsgBox "MMTページが見つかりません。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set mp = GetMMTChildTabs(pg)
    On Error GoTo 0
    If mp Is Nothing Then
        MsgBox "子タブ(mpMMTChild)が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 値だけクリア
    Call MMT_ClearMMTCombos(mp)

    parts = Split(s, ";")
    For Each itm In parts
        itm = Trim$(CStr(itm))
        If Len(itm) = 0 Then GoTo cont

        Dim f() As String, okR As Boolean, okL As Boolean, foundR As Boolean, foundL As Boolean
        f = Split(itm, "|")
        If UBound(f) < 1 Then GoTo cont

        side = val(f(0))
        key = CStr(f(1))
        vR = IIf(UBound(f) >= 2, CStr(f(2)), "")
        vL = IIf(UBound(f) >= 3, CStr(f(3)), "")

        If side < 0 Or side > mp.Pages.count - 1 Then
            Debug.Print "[LOAD][MMT] side不正: "; side; " / key="; key
            GoTo cont
        End If

        Set p = mp.Pages(side)

        ' 名前規則：cboR_ / cboL_ ＋ 項目名
        Set cboR = Nothing: Set cboL = Nothing
        On Error Resume Next
        Set cboR = p.controls("cboR_" & key)
        Set cboL = p.controls("cboL_" & key)
        On Error GoTo 0

        foundR = Not cboR Is Nothing
foundL = Not cboL Is Nothing

' 空（=値なし）は「何もしない＝OK」とみなす
okR = (Len(vR) = 0)
okL = (Len(vL) = 0)

If foundR And Len(vR) > 0 Then cboR.ListIndex = val(vR): okR = (cboR.ListIndex = val(vR))


If foundL And Len(vL) > 0 Then cboL.ListIndex = val(vL): okL = (cboL.ListIndex = val(vL))



cont:
    Next itm
End Sub


' === 共有ユーティリティ ===

'=== 1) 見出しを完全一致で検索して列番号を返す（無ければ0） ===
Public Function FindColByHeaderExact(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        FindColByHeaderExact = f.Column
    Else
        FindColByHeaderExact = 0
    End If
End Function



'=== 2) 見出し列を保証して列番号を返す（無ければ作る） ===
Public Function EnsureHeaderCol(ws As Worksheet, header As String) As Long
    Dim c As Long
    c = FindColByHeaderExact(ws, header)
    If c = 0 Then
        c = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, c).value = header
    End If
    EnsureHeaderCol = c
End Function





'=== 見出し列「MMT_IO」を探す（無ければ作る） ===
Private Function FindOrCreateHeader(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If CStr(ws.Cells(1, c).value) = header Then
            FindOrCreateHeader = c
            Exit Function
        End If
    Next
    '無ければ末尾+1に作成
    FindOrCreateHeader = lastCol + 1
    ws.Cells(1, FindOrCreateHeader).value = header
End Function


'=== 区切り文字 | の n 番目(1始まり)を返す ===
Private Function ParseField(ByVal rec As String, ByVal idx As Long) As String
    Dim a As Variant: a = Split(rec, "|")
    If idx >= 1 And idx - 1 <= UBound(a) Then
        ParseField = a(idx - 1)
    Else
        ParseField = ""
    End If
End Function



Public Sub MMT_DebugSurvey_Page14LegacyNames()
    Dim pg As Object
    Dim i As Long
    Dim ctl As Object
    Dim tp As String
    Dim nm As String
    Dim parentNm As String
    Dim vis As String
    Dim legacyHit As Boolean
    Dim scanned As Long
    Dim candidateCount As Long
    Dim matchedCount As Long

    Set pg = GetMMTPageByName_Page14(frmEval)
    If pg Is Nothing Then
        Debug.Print "[MMT][SURVEY] Page14 not found (root cause candidate=B: target traversal mismatch)"
        Exit Sub
    End If

    Debug.Print "[MMT][SURVEY] START page=" & SafeObjName(pg) & " parent=" & SafeObjName(pg.parent)

    For i = 0 To pg.controls.count - 1
        Set ctl = pg.controls(i)
        tp = TypeName(ctl)
        scanned = scanned + 1

        If tp = "Label" Or tp = "ComboBox" Then
            nm = CStr(ctl.name)
            parentNm = SafeObjName(ctl.parent)
            vis = SafeObjVisible(ctl)
            legacyHit = IsLegacyMMTControlName(nm)

            candidateCount = candidateCount + 1
            If legacyHit Then matchedCount = matchedCount + 1

            Debug.Print "[MMT][SURVEY]", _
                        "Name=" & nm, _
                        "Type=" & tp, _
                        "Parent=" & parentNm, _
                        "Visible=" & vis, _
                        "LegacyHit=" & CStr(legacyHit)
        End If
    Next i

    Debug.Print "[MMT][SURVEY][SUMMARY] scanned=" & scanned & " candidates=" & candidateCount & " matched=" & matchedCount & " unmatched=" & (candidateCount - matchedCount)

    If candidateCount = 0 Then
        Debug.Print "[MMT][SURVEY][CAUSE] B: Page14 direct controls has no Label/ComboBox; traversal target likely differs from actual legacy-control parent"
    ElseIf matchedCount = 0 Then
        Debug.Print "[MMT][SURVEY][CAUSE] A: name matching likely failed (LegacyMMTKeys/IsLegacyMMTControlName mismatch)"
    Else
        Debug.Print "[MMT][SURVEY][CAUSE] A/B not definitive; if UI unchanged, verify remove/hide runtime errors => C"
    End If
End Sub

Private Function GetMMTPageByName_Page14(ByVal frm As Object) As Object
    Dim mp1 As Object
    Dim pgRoot As Object
    Dim host As Object
    Dim mpPhys As Object
    Dim i As Long

    If frm Is Nothing Then Exit Function

    On Error Resume Next
    Set mp1 = frm.controls("MultiPage1")
    If mp1 Is Nothing Then Exit Function

    Set pgRoot = mp1.Pages(2)
    If pgRoot Is Nothing Then Exit Function

    Set host = pgRoot.controls("Frame3")
    If host Is Nothing Then Exit Function

    Set mpPhys = host.controls("mpPhys")
    If mpPhys Is Nothing Then Exit Function
    On Error GoTo 0

    For i = 0 To mpPhys.Pages.count - 1
        If LCase$(CStr(mpPhys.Pages(i).name)) = "page14" Then
            Set GetMMTPageByName_Page14 = mpPhys.Pages(i)
            Exit Function
        End If
    Next i
End Function

Private Function SafeObjName(ByVal obj As Object) As String
    On Error Resume Next
    SafeObjName = CStr(obj.name)
    If Err.Number <> 0 Then
        Err.Clear
        SafeObjName = "(name-error)"
    End If
    On Error GoTo 0
End Function

Private Function SafeObjVisible(ByVal obj As Object) As String
    On Error Resume Next
    SafeObjVisible = CStr(obj.Visible)
    If Err.Number <> 0 Then
        Err.Clear
        SafeObjVisible = "(visible-error)"
    End If
    On Error GoTo 0
End Function


