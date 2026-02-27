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

    Dim host As Object, mp As Object


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
    Set mp = GetMMTChildTabs(pg, host)
    
    If mp Is Nothing Then

        MsgBox "子タブ(mpMMTChild)が作成できません。", vbExclamation
        Exit Sub
    End If

    ' mpサイズも毎回追従
    mp.Width = host.InsideWidth
    mp.Height = host.InsideHeight

    '--- 子タブの中身を作り直す（MMTGENだけ消す） ---
    MMT_ClearGen mp.Pages(0)
    MMT_ClearGen mp.Pages(1)

    BuildMMTPage mp.Pages(0), Array("肩屈曲", "肩伸展", "肩外転", "肩内旋", "肩外旋", _
                                    "肘屈曲", "肘伸展", "前腕回内", "前腕回外", _
                                    "手関節掌屈", "手関節背屈", "指屈曲", "指伸展", "母指対立")
    BuildMMTPage mp.Pages(1), Array("股屈曲", "股伸展", "股外転", "股内転", _
                                    "膝屈曲", "膝伸展", "足関節背屈", "足関節底屈", "母趾伸展")
    
    DoEvents
    Resize_MMTChildHost_ToPage
    
    Exit Sub

   


RRTRACE:
    Debug.Print "[MMT ERROR]", Err.Number, Err.Description

End Sub


Public Function GetMMTHost(ByVal pg As Object) As Object
    Dim host As Object
    Dim cand As Variant
    Dim i As Long
    Dim j As Long
    Dim c As Object
    Dim mpProbe As Object
    
    If pg Is Nothing Then Exit Function
    
    ' 1) 候補名を優先
    For Each cand In Array("fraMMTHost", "Frame9", "fraMMTWrap")
        On Error Resume Next
        Set host = pg.Controls(CStr(cand))
        On Error GoTo 0
        
        If Not host Is Nothing Then
            If TypeName(host) = "Frame" Then
#If APP_DEBUG Then
                Debug.Print "[MMT][HOST] name-hit=" & CStr(cand) & " type=" & TypeName(host)
#End If
                Set GetMMTHost = host
                Exit Function
            End If
            Set host = Nothing
        End If
    Next cand
    
    ' 2) Frame を走査して特徴で推定
    For i = 0 To pg.Controls.Count - 1
        Set c = pg.Controls(i)
        If TypeName(c) = "Frame" Then
            
            ' mpMMTChild を持っているか（アクセスエラーもあり得るのでガード）
            Set mpProbe = Nothing
            On Error Resume Next
            Set mpProbe = c.Controls("mpMMTChild")
            On Error GoTo 0
            
            If Not mpProbe Is Nothing Then
#If APP_DEBUG Then
                Debug.Print "[MMT][HOST] inferred=" & c.name & " (has mpMMTChild)"
#End If
                Set GetMMTHost = c
                Exit Function
            End If
            
            ' MultiPage 子を持っているか
            On Error Resume Next
            If c.Controls.Count > 0 Then
                For j = 0 To c.Controls.Count - 1
                    If TypeName(c.Controls(j)) = "MultiPage" Then
#If APP_DEBUG Then
                        Debug.Print "[MMT][HOST] inferred=" & c.name & " (has MultiPage child)"
#End If
                        Set GetMMTHost = c
                        Exit Function
                    End If
                Next j
            End If
            On Error GoTo 0
        End If
    Next i
    
#If APP_DEBUG Then
    Debug.Print "[MMT][HOST] not found -> skip"
#End If
End Function

Public Function GetMMTChildTabs(ByVal pg As Object, Optional ByVal host As Object = Nothing) As Object
    Dim mp As Object
    Dim i As Long
    
    If host Is Nothing Then Set host = GetMMTHost(pg)
    If host Is Nothing Then Exit Function
    
    On Error Resume Next
    Set mp = host.Controls("mpMMTChild")
    On Error GoTo 0
    
    If mp Is Nothing Then
        On Error Resume Next
        For i = 0 To host.Controls.Count - 1
            If TypeName(host.Controls(i)) = "MultiPage" Then
                Set mp = host.Controls(i)
                Exit For
            End If
        Next i
        On Error GoTo 0
    End If
    
    If mp Is Nothing Then
        On Error Resume Next
        Set mp = host.Controls.Add("Forms.MultiPage.1", "mpMMTChild", True)
        On Error GoTo 0
        
        If mp Is Nothing Then Exit Function
        
        mp.Left = 0
        mp.Top = 0
        mp.Style = 0
        mp.Pages.Clear
        mp.Pages.Add.caption = ChrW(&H4E0A) & ChrW(&H80A2)
        mp.Pages.Add.caption = ChrW(&H4E0B) & ChrW(&H80A2)
    End If
    
    If mp.Pages.Count < 2 Then
        Do While mp.Pages.Count < 2
            mp.Pages.Add
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
    For Each ctl In frm.Controls
        If TypeName(ctl) = "MultiPage" Then
            Dim i As Long
            For i = 0 To ctl.Pages.Count - 1
                Set pg = ctl.Pages(i)
                If PageHasMMTSignature(pg) Then
                    Set GetMMTPage = pg
                    Exit Function
                End If
            Next i
        End If
    Next ctl
End Function

Private Function PageHasMMTSignature(ByVal pg As Object) As Boolean
    Dim c As Object

    If pg Is Nothing Then Exit Function

    ' 「mpMMTChild」や「Frame9」など、MMTページ固有の痕跡で判定
    For Each c In pg.Controls
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
                    ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.Controls.Add("Forms.ComboBox.1", nm, True)
    o.Left = l: o.Top = t: o.Width = w: o.Height = h
    o.Style = MSForms.fmStyleDropDownList: o.BoundColumn = 1
    o.List = Split("0,1,2,3,4,5", ","): o.tag = "MMTGEN"
End Sub


Private Sub MakeLbl(ByVal pg As Object, ByVal nm As String, ByVal cap As String, _
                    ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.label
    Set o = pg.Controls.Add("Forms.Label.1", nm, True)
    o.caption = cap: o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.tag = "MMTGEN"
End Sub



'--- 1ページ分のUI生成 ---
Private Sub BuildMMTPage(ByVal pg As Object, ByVal items As Variant)
    Const ROW_H As Single = 24, LBL_W As Single = 130, COL_W As Single = 90, gap As Single = 12
    Dim x0 As Single, y0 As Single: x0 = 20: y0 = 28

    MakeLbl pg, "lblHdrMus", "筋群", x0, y0 - 20, 60, 18
    MakeLbl pg, "lblHdrR", "右", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLbl pg, "lblHdrL", "左", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, Y As Single: Y = y0
    For i = LBound(items) To UBound(items)
        Dim key As String: key = CStr(items(i))
        MakeLbl pg, "lbl_" & key, key, x0, Y + 3, LBL_W, 18
        MakeCbo pg, "cboR_" & key, x0 + LBL_W + gap, Y, COL_W, 18
        MakeCbo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, Y, COL_W, 18
        Y = Y + ROW_H
    Next
End Sub

'--- 自動生成（MMTGEN）だけ掃除 ---
Private Sub MMT_ClearGen(ByVal pg As Object)
    Dim idx As Long
    For idx = pg.Controls.Count - 1 To 0 Step -1
        If Left$(pg.Controls(idx).tag & "", 6) = "MMTGEN" Then
            pg.Controls.Remove pg.Controls(idx).name
        End If
    Next
End Sub

'--- 子タブ内の全 ComboBox をいったんクリア ---
Private Sub MMT_ClearMMTCombos(ByVal mp As MSForms.MultiPage)
    Dim pg As MSForms.Page
    Dim c As Object  '（ControlでもOK）

    For Each pg In mp.Pages
        For Each c In pg.Controls
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
        c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
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
    Set pg = GetMMTPage()
    If pg Is Nothing Then Exit Function

    On Error Resume Next
    Set mp = GetMMTChildTabs(pg)
    On Error GoTo 0
    If mp Is Nothing Then Exit Function

    ReDim parts(0 To 0): n = -1

    For p = 0 To mp.Pages.Count - 1
        For Each c In mp.Pages(p).Controls
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
                    lVal = CStr(mp.Pages(p).Controls("cboL_" & nm).value)
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

    If mp Is Nothing Then
        Debug.Print "[MMT] building child tabs"
        MMT_BuildChildTabs_Direct owner
        Set mp = GetMMTChildTabs(pg)
        Debug.Print "[MMT] after rebuild", TypeName(mp)
        If mp Is Nothing Then Exit Sub
    End If

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
    Dim p As MSForms.Page
    Dim cboR As MSForms.ComboBox, cboL As MSForms.ComboBox

    Set pg = GetMMTPage()
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

        If side < 0 Or side > mp.Pages.Count - 1 Then
            Debug.Print "[LOAD][MMT] side不正: "; side; " / key="; key
            GoTo cont
        End If

        Set p = mp.Pages(side)

        ' 名前規則：cboR_ / cboL_ ＋ 項目名
        Set cboR = Nothing: Set cboL = Nothing
        On Error Resume Next
        Set cboR = p.Controls("cboR_" & key)
        Set cboL = p.Controls("cboL_" & key)
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
        c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, c).value = header
    End If
    EnsureHeaderCol = c
End Function





'=== 見出し列「MMT_IO」を探す（無ければ作る） ===
Private Function FindOrCreateHeader(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
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
