Attribute VB_Name = "modEvalDataArchive"
Option Explicit

Public gObj As Object



Public Sub DumpUFTree(ByVal uf As Object)
    On Error GoTo EH

#If APP_DEBUG Then
    Debug.Print String(90, "=")
    Debug.Print "[UF TREE]"; TypeName(uf); " Name="; uf.name
    Debug.Print "  UF: W=" & f2(uf.Width) & " H=" & f2(uf.Height) & _
                " InW=" & f2(uf.InsideWidth) & " InH=" & f2(uf.InsideHeight) & _
                " ScrollH=" & f2(NZ(uf.ScrollHeight)) & " ScrollW=" & f2(NZ(uf.ScrollWidth)) & _
                " ScrollBars=" & NZ(uf.ScrollBars)
#End If

    DumpControlsRecursive uf, 0
    Exit Sub

EH:
    Debug.Print "[DumpUFTree][ERR]"; Err.Number; Err.Description
End Sub

Private Sub DumpControlsRecursive(ByVal parent As Object, ByVal depth As Long)
    On Error GoTo EH

    Dim c As Object
    For Each c In parent.Controls
        DumpOne c, depth

        ' 子を持つ可能性があるものだけ潜る（Frame / MultiPage / Page）
        If HasControls(c) Then
            DumpControlsRecursive c, depth + 1
        End If
    Next
    Exit Sub

EH:
    Debug.Print Ind(depth) & "[RECURSE ERR] " & TypeName(parent) & " " & Err.Number & " " & Err.Description
End Sub

Private Sub DumpOne(ByVal c As Object, ByVal depth As Long)
    On Error GoTo EH

    Dim line As String
    line = Ind(depth) & "- " & TypeName(c) & "  Name=" & SafeName(c)

    line = line & "  L=" & f2(SafeProp(c, "Left")) & _
                  " T=" & f2(SafeProp(c, "Top")) & _
                  " W=" & f2(SafeProp(c, "Width")) & _
                  " H=" & f2(SafeProp(c, "Height"))

    ' よく事故る系も控えめに拾う（取れないプロパティは無視）
    line = line & "  Vis=" & SafeProp(c, "Visible")
    line = line & "  En=" & SafeProp(c, "Enabled")

    ' MultiPage / Page / FrameはInside/Scrollも出す
    If TypeName(c) = "MultiPage" Or TypeName(c) = "Page" Or TypeName(c) = "Frame" Then
        line = line & "  InH=" & f2(SafeProp(c, "InsideHeight")) & " InW=" & f2(SafeProp(c, "InsideWidth"))
        line = line & "  ScrH=" & f2(SafeProp(c, "ScrollHeight")) & " ScrW=" & f2(SafeProp(c, "ScrollWidth"))
        line = line & "  ScrBars=" & SafeProp(c, "ScrollBars")
    End If

#If APP_DEBUG Then
    Debug.Print line
#End If

    ' MultiPage の Pages を明示的に列挙
    If TypeName(c) = "MultiPage" Then
        DumpMultiPagePages c, depth + 1
    End If

    Exit Sub

EH:
#If APP_DEBUG Then
    Debug.Print Ind(depth) & "[DUMP ERR] " & TypeName(c) & " " & Err.Number & " " & Err.Description
#End If
End Sub

Private Sub DumpMultiPagePages(ByVal mp As Object, ByVal depth As Long)
    On Error GoTo EH

    Dim i As Long, pg As Object
    For i = 0 To mp.Pages.Count - 1
        Set pg = mp.Pages(i)
#If APP_DEBUG Then
        Debug.Print Ind(depth) & "* Page(" & i & ") Name=" & pg.name & _
                    "  L=" & f2(pg.Left) & " T=" & f2(pg.Top) & _
                    " W=" & f2(pg.Width) & " H=" & f2(pg.Height) & _
                    "  Vis=" & pg.Visible
#End If
        If HasControls(pg) Then DumpControlsRecursive pg, depth + 1
    Next
    Exit Sub

EH:
#If APP_DEBUG Then
    Debug.Print Ind(depth) & "[PAGES ERR] " & Err.Number & " " & Err.Description
#End If
End Sub

Private Function HasControls(ByVal o As Object) As Boolean
    On Error GoTo EH
    Dim n As Long
    n = o.Controls.Count
    HasControls = (n >= 0)
    Exit Function
EH:
    HasControls = False
End Function

Private Function SafeName(ByVal o As Object) As String
    On Error GoTo EH
    SafeName = o.name
    Exit Function
EH:
    SafeName = "<?>"
End Function

Private Function SafeProp(ByVal o As Object, ByVal propName As String) As Variant
    On Error GoTo EH
    SafeProp = CallByName(o, propName, VbGet)
    Exit Function
EH:
    SafeProp = "n/a"
End Function

Private Function Ind(ByVal depth As Long) As String
    Ind = String$(depth * 2, " ")
End Function

Private Function f2(ByVal v As Variant) As String
    On Error GoTo EH
    f2 = Format$(CDbl(v), "0.00")
    Exit Function
EH:
    f2 = CStr(v)
End Function

Private Function NZ(ByVal v As Variant) As Variant
    If IsEmpty(v) Then
        NZ = "empty"
    Else
        NZ = v
    End If
End Function





Public Sub DumpUFTree_ToFile(ByVal uf As Object)

MsgBox "このDumpは危険版です。modUFDumpSafe の DumpFrmEvalTree_ToFile_Safe を使ってください。", vbExclamation
Exit Sub


    On Error GoTo EH

    Dim p As String
    p = Environ$("TEMP") & "\frmEval_tree_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff

    Print #ff, String(90, "=")
    Print #ff, "[UF TREE] " & TypeName(uf) & " Name=" & uf.name
    Print #ff, "  UF: W=" & f2(uf.Width) & " H=" & f2(uf.Height) & _
               " InW=" & f2(uf.InsideWidth) & " InH=" & f2(uf.InsideHeight)

    DumpControlsRecursive_ToFile uf, 0, ff

    Close #ff
    Debug.Print "[SAVED] " & p
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Debug.Print "[DumpUFTree_ToFile][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Sub DumpControlsRecursive_ToFile(ByVal parent As Object, ByVal depth As Long, ByVal ff As Integer)
    On Error GoTo EH

    Dim c As Object
    For Each c In parent.Controls

        Print #ff, OneLine(c, depth)

        If TypeName(c) = "MultiPage" Then
            DumpPages_ToFile c, depth + 1, ff
        ElseIf HasControls(c) Then
            DumpControlsRecursive_ToFile c, depth + 1, ff
        End If
    Next
    Exit Sub

EH:
    Print #ff, Ind(depth) & "[RECURSE ERR] " & TypeName(parent) & " " & Err.Number & " " & Err.Description
End Sub


Private Sub DumpPages_ToFile(ByVal mp As Object, ByVal depth As Long, ByVal ff As Integer)
    On Error GoTo EH
    Dim i As Long, pg As Object

    For i = 0 To mp.Pages.Count - 1
        Set pg = mp.Pages(i)

        Print #ff, Ind(depth) & "* Page(" & i & ") Name=" & pg.name & _
                   " Caption=" & pg.caption

        ' ★ ここが重要：Page を起点に再帰
        DumpControlsRecursive_ToFile pg, depth + 1, ff
    Next
    Exit Sub

EH:
    Print #ff, Ind(depth) & "[PAGES ERR] " & Err.Number & " " & Err.Description
End Sub


Private Function OneLine(ByVal c As Object, ByVal depth As Long) As String
    On Error GoTo EH

    Dim s As String
    s = Ind(depth) & "- " & TypeName(c) & " Name=" & SafeName(c) & _
        " L=" & f2(SafeProp(c, "Left")) & _
        " T=" & f2(SafeProp(c, "Top")) & _
        " W=" & f2(SafeProp(c, "Width")) & _
        " H=" & f2(SafeProp(c, "Height"))

    If TypeName(c) = "MultiPage" Or TypeName(c) = "Page" Or TypeName(c) = "Frame" Then
        s = s & " InH=" & f2(SafeProp(c, "InsideHeight")) & " InW=" & f2(SafeProp(c, "InsideWidth"))
    End If

    OneLine = s
    Exit Function

EH:
    OneLine = Ind(depth) & "- " & TypeName(c) & " (line err " & Err.Number & ")"
End Function



Public Sub DumpTreeFile_Head(ByVal n As Long)
    Dim f As String
    f = GetLatestTreeFile()
    If f = "" Then
        Debug.Print "[ERR] tree file not found"
        Exit Sub
    End If

    Dim ff As Integer: ff = FreeFile
    Open f For Input As #ff

    Dim i As Long
    Dim line As String
    Debug.Print "[HEAD]"; f
    For i = 1 To n
        If EOF(ff) Then Exit For
        Line Input #ff, line
        Debug.Print line
    Next

    Close #ff
End Sub

Private Function GetLatestTreeFile() As String
    Dim p As String: p = Environ$("TEMP") & "\"
    Dim f As String: f = Dir$(p & "frmEval_tree_*.txt")
    Dim latest As String

    Do While f <> ""
        latest = p & f
        f = Dir$
    Loop

    GetLatestTreeFile = latest
End Function




Public Sub Diag_MultiPagePages(ByVal uf As Object, ByVal mpName As String)
    On Error GoTo EH

    Dim mp As Object
    Set mp = uf.Controls(mpName)

    Debug.Print "[MP]"; mpName; " Type=" & TypeName(mp); " PagesCount=" & mp.Pages.Count
    
    
    
    Dim pg As Object
Dim i As Long: i = 0
For Each pg In mp.Pages
    Debug.Print "  Page(" & i & ") Type=" & TypeName(pg) & _
                " Name=" & SafeProp(pg, "Name") & _
                " Width=" & SafeProp(pg, "Width") & _
                " Height=" & SafeProp(pg, "Height")
    i = i + 1
Next

    Exit Sub
    
    Call Diag_MultiPagePages(frmEval, "MultiPage1")

    

EH:
    Debug.Print "[Diag_MultiPagePages][ERR]"; Err.Number; Err.Description
    

End Sub




Public Sub DumpMP_PageTopControls(ByVal uf As Object, ByVal mpName As String)
    On Error GoTo EH

    Dim mp As Object: Set mp = uf.Controls(mpName)
    Debug.Print "[MP TOP] " & mpName & " Pages=" & mp.Pages.Count

    Dim pg As Object, i As Long
    i = 0
    For Each pg In mp.Pages
        Debug.Print " Page(" & i & ") " & SafeProp(pg, "Name") & "  Ctls=" & PgCtlCount(pg)

        Dim c As Object
        For Each c In pg.Controls
            Debug.Print "   - " & TypeName(c) & " " & SafeName(c)
        Next

        i = i + 1
    Next
    Exit Sub

EH:
    Debug.Print "[DumpMP_PageTopControls][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Function PgCtlCount(ByVal pg As Object) As Long
    On Error GoTo EH
    PgCtlCount = pg.Controls.Count
    Exit Function
EH:
    PgCtlCount = -1
End Function




Public Sub DumpMP_PageCounts(ByVal uf As Object, ByVal mpName As String)
    On Error GoTo EH

    Dim mp As Object: Set mp = uf.Controls(mpName)
    Debug.Print "[MP COUNTS] " & mpName & " Pages=" & mp.Pages.Count

    Dim pg As Object, i As Long
    i = 0
    For Each pg In mp.Pages
        Debug.Print " Page(" & i & ") " & SafeProp(pg, "Name") & " Controls=" & PgCtlCount(pg)
        i = i + 1
    Next
    Exit Sub

EH:
    Debug.Print "[DumpMP_PageCounts][ERR] " & Err.Number & " " & Err.Description
End Sub




Public Sub DumpMP_OnePage_ToFile(ByVal uf As Object, ByVal mpName As String, ByVal pageName As String)
    On Error GoTo EH

    Dim mp As Object: Set mp = uf.Controls(mpName)

    Dim pg As Object
    Set pg = FindPageByName(mp, pageName)
    If pg Is Nothing Then
        Debug.Print "[ERR] page not found: " & pageName
        Exit Sub
    End If

    Dim p As String
    p = Environ$("TEMP") & "\frmEval_" & mpName & "_" & pageName & "_tree_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff

    Print #ff, String(90, "=")
    Print #ff, "[ONE PAGE TREE] " & mpName & " / " & pageName
    Print #ff, " Controls=" & PgCtlCount(pg)

    DumpControlsRecursive_ToFile pg, 0, ff

    Close #ff
    Debug.Print "[SAVED] " & p
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Debug.Print "[DumpMP_OnePage_ToFile][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Function FindPageByName(ByVal mp As Object, ByVal pageName As String) As Object
    On Error GoTo EH
    Dim pg As Object
    For Each pg In mp.Pages
        If StrComp(SafeProp(pg, "Name"), pageName, vbTextCompare) = 0 Then
            Set FindPageByName = pg
            Exit Function
        End If
    Next
    Set FindPageByName = Nothing
    Exit Function
EH:
    Set FindPageByName = Nothing
End Function



Public Sub DumpFile_TailByPath(ByVal path As String, ByVal n As Long)
    On Error GoTo EH

    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff

    Dim buf() As String
    ReDim buf(1 To n)
    Dim idx As Long: idx = 0

    Dim line As String
    Do While Not EOF(ff)
        Line Input #ff, line
        idx = idx + 1
        buf(((idx - 1) Mod n) + 1) = line
    Loop
    Close #ff

    Debug.Print "[TAIL]"; path
    Dim startPos As Long
    If idx < n Then
        startPos = 1
    Else
        startPos = ((idx - 1) Mod n) + 1
    End If

    Dim k As Long, pos As Long, outCount As Long
    outCount = IIf(idx < n, idx, n)

    For k = 0 To outCount - 1
        pos = startPos + k
        If pos > n Then pos = pos - n
        Debug.Print buf(pos)
    Next
    Exit Sub

EH:
    Debug.Print "[DumpFile_TailByPath][ERR] " & Err.Number & " " & Err.Description
End Sub





Public Sub FindControlParents(ByVal uf As Object, ByVal targetName As String)
    On Error GoTo EH
    Debug.Print "[FIND PARENTS] " & targetName
    FindParentsRecursive uf, targetName, TypeName(uf) & ":" & uf.name
    Exit Sub
EH:
    Debug.Print "[FindControlParents][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Sub FindParentsRecursive(ByVal parent As Object, ByVal targetName As String, ByVal path As String)
    On Error GoTo EH

    Dim c As Object
    For Each c In parent.Controls
        If StrComp(SafeName(c), targetName, vbTextCompare) = 0 Then
            Debug.Print " HIT path=" & path & " -> " & TypeName(c) & ":" & SafeName(c)
        End If

        If TypeName(c) = "MultiPage" Then
            Dim pg As Object
            For Each pg In c.Pages
                FindParentsRecursive pg, targetName, path & " -> MultiPage:" & SafeName(c) & " -> Page:" & SafeProp(pg, "Name")
            Next
        ElseIf HasControls(c) Then
            FindParentsRecursive c, targetName, path & " -> " & TypeName(c) & ":" & SafeName(c)
        End If
    Next
    Exit Sub

EH:
    Debug.Print "[FindParentsRecursive][ERR] " & Err.Number & " " & Err.Description & " path=" & path
End Sub




Public Sub DumpMP_OnePage_TreeByParent_ToFile(ByVal uf As Object, ByVal mpName As String, ByVal pageName As String)
    On Error GoTo EH

    Dim mp As Object: Set mp = uf.Controls(mpName)
    Dim pg As Object: Set pg = FindPageByName(mp, pageName)
    If pg Is Nothing Then
        Debug.Print "[ERR] page not found: " & pageName
        Exit Sub
    End If

    Dim p As String
    p = Environ$("TEMP") & "\frmEval_" & mpName & "_" & pageName & "_treeByParent_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff

    Print #ff, String(90, "=")
    Print #ff, "[TREE BY PARENT] " & mpName & " / " & pageName

    ' 直下（Parentがpgのもの）だけを列挙して、そこから再帰
    DumpChildrenByParent_ToFile pg, pg, 0, ff

    Close #ff
    Debug.Print "[SAVED] " & p
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Debug.Print "[DumpMP_OnePage_TreeByParent_ToFile][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Sub DumpChildrenByParent_ToFile(ByVal root As Object, ByVal parent As Object, ByVal depth As Long, ByVal ff As Integer)
    On Error GoTo EH

    Dim c As Object
    For Each c In root.Controls
        If HasParent(c) Then
            If c.parent Is parent Then
                Print #ff, Ind(depth) & "- " & TypeName(c) & " " & SafeName(c) & _
                           " L=" & f2(SafeProp(c, "Left")) & _
                           " T=" & f2(SafeProp(c, "Top")) & _
                           " W=" & f2(SafeProp(c, "Width")) & _
                           " H=" & f2(SafeProp(c, "Height"))
                ' 次の階層へ
                If CanHaveChildren(c) Then
                    DumpChildrenByParent_ToFile root, c, depth + 1, ff
                End If
            End If
        End If
    Next
    Exit Sub

EH:
    Print #ff, Ind(depth) & "[ERR] " & Err.Number & " " & Err.Description
End Sub

Private Function HasParent(ByVal o As Object) As Boolean
    On Error GoTo EH
    Dim tmp As Object
    Set tmp = o.parent
    HasParent = True
    Exit Function
EH:
    HasParent = False
End Function

Private Function CanHaveChildren(ByVal o As Object) As Boolean
    Dim t As String: t = TypeName(o)
    CanHaveChildren = (t = "Frame" Or t = "MultiPage" Or t = "Page")
End Function




Public Sub DumpTreeFile_HeadByPath(ByVal path As String, ByVal n As Long)
    On Error GoTo EH

    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff

    Dim i As Long, line As String
    Debug.Print "[HEAD]"; path
    For i = 1 To n
        If EOF(ff) Then Exit For
        Line Input #ff, line
        Debug.Print line
    Next

    Close #ff
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Debug.Print "[DumpTreeFile_HeadByPath][ERR] " & Err.Number & " " & Err.Description
End Sub




Public Sub Diag_mpPhys_PageCounts()
    Dim mp As Object
    Set mp = frmEval.Controls("mpPhys")
    Debug.Print "[MP COUNTS] mpPhys Pages=" & mp.Pages.Count

    Dim pg As Object, i As Long
    i = 0
    For Each pg In mp.Pages
        Debug.Print " Page(" & i & ") " & SafeProp(pg, "Name") & " Controls=" & PgCtlCount(pg)
        i = i + 1
    Next
End Sub



Public Sub Diag_frKyo_ComboIndex()
    Dim f As Object, i As Long, c As Object

    Set f = frmEval.Controls("MultiPage1").Pages(3).Controls("Frame4") _
                .Controls("mpADL").Pages(2).Controls("frKyo")

    For i = 0 To f.Controls.Count - 1
        Set c = f.Controls(i)
        If TypeName(c) = "ComboBox" Then
            Debug.Print i, "[" & c.name & "]", c.Left, c.Top
        End If
    Next i
End Sub


Public Sub Fix_frKyo_AnonymousCombos()
    Dim f As Object
    Set f = frmEval.Controls("MultiPage1").Pages(3) _
                .Controls("Frame4").Controls("mpADL") _
                .Pages(2).Controls("frKyo")

    f.Controls(7).name = "cmbKyo_StandUp"
    f.Controls(9).name = "cmbKyo_StandHold"

    Debug.Print "[FIX] frKyo unnamed ComboBoxes renamed"
End Sub


Public Sub Test_Dump_mpPhys_Page8()
    DumpTreeByParent_ToFile frmEval.Controls("MultiPage1").Pages(2).Controls("Frame3").Controls("mpPhys").Pages(0)
End Sub




Public Sub DumpTreeByParent_ToFile(ByVal root As Object)
Debug.Print "[DUMP START] root=" & TypeName(root)


    On Error GoTo EH
    
    Static sSeq As Long
    sSeq = sSeq + 1


    Dim p As String
    p = Environ$("TEMP") & "\treeByParent_" & TypeName(root) & "_" & GetNameSafe(root) & "_" & GetCaptionSafe(root) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff

    Print #ff, String(90, "=")
    Print #ff, "[TREE BY PARENT] " & TypeName(root) & " Name=" & GetNameSafe(root)

    If TypeName(root) = "MultiPage" Then
    DumpPages_ToFile root, 0, ff
Else
    DumpControlsRecursive_ToFile root, 0, ff
End If


    Close #ff
    Debug.Print "[SAVED] " & p
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Debug.Print "[DumpTreeByParent_ToFile][ERR] " & Err.Number & " " & Err.Description
End Sub

Private Function GetNameSafe(ByVal o As Object) As String
    On Error Resume Next
    GetNameSafe = o.name
    If Len(GetNameSafe) = 0 Then GetNameSafe = "(no-name)"
End Function




Public Sub Test_TypeName_mpPhys_Page8()
    Dim o As Object
    Set o = frmEval.Controls("MultiPage1").Pages(2).Controls("Frame3").Controls("mpPhys").Pages(0)

    Debug.Print "TypeName(o)=", TypeName(o)
    Debug.Print "o.Name=", o.name

    DumpTreeByParent_ToFile o
End Sub




'=== 全レイアウト構造を一括ダンプ（Immediateはこれを呼ぶだけ）===
Public Sub Dump_AllLayout_Snapshot()
    On Error GoTo EH

    Dim mp1 As Object, pg As Object
    Dim mpPhys As Object, mpADL As Object
    Dim mp2 As Object, mp3 As Object

    '--- 1) ルート MultiPage1 (8ページ) ---
    Set mp1 = frmEval.Controls("MultiPage1")

    Dim i As Long
    For i = 0 To mp1.Pages.Count - 1
        Set pg = mp1.Pages(i)
        DumpTreeByParent_ToFile pg
    Next i

    '--- 2) Page3 -> Frame3 -> mpPhys (6ページ) ---
    Set mpPhys = mp1.Pages(2).Controls("Frame3").Controls("mpPhys")
    For i = 0 To mpPhys.Pages.Count - 1
        DumpTreeByParent_ToFile mpPhys.Pages(i)
    Next i

    '--- 3) Page4 -> Frame4 -> mpADL (3ページ) ---
    Set mpADL = mp1.Pages(3).Controls("Frame4").Controls("mpADL")
    For i = 0 To mpADL.Pages.Count - 1
        DumpTreeByParent_ToFile mpADL.Pages(i)
    Next i

    '--- 4) Page6 -> Frame6 -> MultiPage2 (3ページ) ---
    Set mp2 = mp1.Pages(5).Controls("Frame6").Controls("MultiPage2")
    For i = 0 To mp2.Pages.Count - 1
        DumpTreeByParent_ToFile mp2.Pages(i)
    Next i

    '--- 5) MultiPage2 の Page9 -> Frame26 -> MultiPage3 (2ページ) ---
    Set mp3 = mp2.Pages(1).Controls("Frame26").Controls("MultiPage3")
    For i = 0 To mp3.Pages.Count - 1
        DumpTreeByParent_ToFile mp3.Pages(i)
    Next i

    Debug.Print "[DONE] Dump_AllLayout_Snapshot"
    Exit Sub

EH:
    Debug.Print "[Dump_AllLayout_Snapshot][ERR] " & Err.Number & " " & Err.Description
End Sub



Private Function GetCaptionSafe(ByVal o As Object) As String
    On Error GoTo EH
    Dim s As String
    s = CStr(CallByName(o, "Caption", VbGet))
    s = Replace(s, "\", "＼")
    s = Replace(s, "/", "／")
    s = Replace(s, ":", "：")
    s = Replace(s, "*", "＊")
    s = Replace(s, "?", "？")
   s = Replace(s, Chr(34), ChrW(&H201D))
    s = Replace(s, "<", "＜")
    s = Replace(s, ">", "＞")
    s = Replace(s, "|", "｜")
    GetCaptionSafe = s
    Exit Function
EH:
    GetCaptionSafe = ""
End Function

