Attribute VB_Name = "modScan"
Option Explicit

' === 譁・ｭ怜・荳ｭ縺ｮ蜃ｺ迴ｾ蝗樊焚繧呈焚縺医ｋ・亥､ｧ譁・ｭ怜ｰ乗枚蟄励ｒ辟｡隕厄ｼ・===
Private Function CountOccur(ByVal haystack As String, ByVal needle As String) As Long
    Dim p As Long, n As Long, s As String, t As String
    s = LCase$(haystack): t = LCase$(needle)
    If Len(t) = 0 Then Exit Function
    p = 1
    Do
        p = InStr(p, s, t, vbBinaryCompare)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + Len(t)
    Loop
    CountOccur = n
End Function

' === 繝励Ο繧ｸ繧ｧ繧ｯ繝亥・菴薙ｒ襍ｰ譟ｻ縺励※繝槭ャ繝励ｒ菴懊ｋ・井ｿ｡鬆ｼ繧｢繧ｯ繧ｻ繧ｹ縺悟ｿ・ｦ・ｼ・===
Public Sub MakeProjectMap()
    ' 蠢・ｦ∬ｨｭ螳・ [繝輔ぃ繧､繝ｫ]竊端繧ｪ繝励す繝ｧ繝ｳ]竊端繧ｻ繧ｭ繝･繝ｪ繝・ぅ 繧ｻ繝ｳ繧ｿ繝ｼ]竊・
    '  [VBA 繝励Ο繧ｸ繧ｧ繧ｯ繝・繧ｪ繝悶ず繧ｧ繧ｯ繝・繝｢繝・Ν縺ｸ縺ｮ繧｢繧ｯ繧ｻ繧ｹ繧剃ｿ｡鬆ｼ縺吶ｋ] 縺ｫ繝√ぉ繝・け
    On Error GoTo EH

    Dim wb As Workbook, sh As Worksheet
    Set wb = ThisWorkbook

    ' 繧ｷ繝ｼ繝域ｺ門ｙ
    Const MAP_NAME As String = "PROJECT_MAP"
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(MAP_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set sh = wb.Worksheets.Add
    sh.name = MAP_NAME

    ' 隕句・縺・
    Dim h As Variant
    h = Array("Component", "Type", "Lines", _
              "mpADL", "EnsureBI_IADL", "BuildKyoOnADL", "RemoveAllMpADL", _
              "CAP_BI", "CAP_IADL", "CAP_KYO", "hostMove", "nextTop")
    Dim i As Long
    For i = LBound(h) To UBound(h)
        sh.Cells(1, i + 1).value = h(i)
    Next

    ' 蜿ら・縺ｯ late binding・・xtensibility 蜿ら・縺ｪ縺励〒蜍輔￥・・
    Dim vbProj As Object: Set vbProj = wb.VBProject
    Dim comp As Object, cm As Object
    Dim r As Long: r = 2

    ' 蝙句ｮ壽焚・亥盾辣ｧ縺ｪ縺怜ｯｾ蠢懶ｼ・
    Const ctStdModule As Long = 1
    Const ctClassMod  As Long = 2
    Const ctMSForm    As Long = 3
    Const ctDocument  As Long = 100

    For Each comp In vbProj.VBComponents
        Dim kind As String, code As String, nLines As Long

        Select Case CLng(comp.Type)
            Case ctStdModule: kind = "StdModule"
            Case ctClassMod: kind = "Class"
            Case ctMSForm: kind = "UserForm"
            Case ctDocument: kind = "Document"
            Case Else: kind = CStr(comp.Type)
        End Select

        Set cm = comp.CodeModule
        nLines = cm.CountOfLines
        If nLines > 0 Then
            code = cm.lines(1, nLines)
        Else
            code = ""
        End If

        sh.Cells(r, 1).value = comp.name
        sh.Cells(r, 2).value = kind
        sh.Cells(r, 3).value = nLines

        ' 繧ｭ繝ｼ隱槭・蜃ｺ迴ｾ蝗樊焚
        sh.Cells(r, 4).value = CountOccur(code, "mpADL")
        sh.Cells(r, 5).value = CountOccur(code, "EnsureBI_IADL")
        sh.Cells(r, 6).value = CountOccur(code, "BuildKyoOnADL")
        sh.Cells(r, 7).value = CountOccur(code, "RemoveAllMpADL")
        sh.Cells(r, 8).value = CountOccur(code, "CAP_BI")
        sh.Cells(r, 9).value = CountOccur(code, "CAP_IADL")
        sh.Cells(r, 10).value = CountOccur(code, "CAP_KYO")
        sh.Cells(r, 11).value = CountOccur(code, "hostMove")
        sh.Cells(r, 12).value = CountOccur(code, "nextTop")

        r = r + 1
    Next

    ' 菴楢｣・
    With sh
        .rows(1).Font.Bold = True
        .Columns.AutoFit
    End With

    MsgBox "PROJECT_MAP 繧剃ｽ懈・縺励∪縺励◆縲・, vbInformation
    Exit Sub
EH:
    MsgBox "MakeProjectMap 繧ｨ繝ｩ繝ｼ: " & Err.Description, vbExclamation
End Sub

' === 縺ｩ縺薙〒 mpADL 繧堤函謌舌＠縺ｦ縺・ｋ縺九・隧ｳ邏ｰ荳隕ｧ・郁｡檎分蜿ｷ莉倥″・・===
Public Sub FindMpADLCreates()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim vbProj As Object: Set vbProj = wb.VBProject
    Dim comp As Object, cm As Object
    Debug.Print String(60, "-")
    Debug.Print "[SCAN] mpADL 繧・Add/Set 縺励※縺・ｋ陦後・荳隕ｧ"
    For Each comp In vbProj.VBComponents
        Set cm = comp.CodeModule
        Dim n As Long: n = cm.CountOfLines
        Dim i As Long
        For i = 1 To n
            Dim ln As String: ln = cm.lines(i, 1)
            ' 逕滓・/莉｣蜈･縺｣縺ｽ縺・｡後ｒ諡ｾ縺・ｼ医＊縺｣縺上ｊ・・
            If InStr(1, LCase$(ln), "set mpadl") > 0 Or _
               InStr(1, LCase$(ln), "controls.add(""forms.multipage.1""") > 0 Then
                Debug.Print comp.name & ":" & i & "  " & Trim$(ln)
            End If
        Next
    Next
    Debug.Print String(60, "-")
    MsgBox "Immediate 繧ｦ繧｣繝ｳ繝峨え・・trl+G・峨↓蜃ｺ蜉帙＠縺ｾ縺励◆縲・, vbInformation
    Exit Sub
EH:
    MsgBox "FindMpADLCreates 繧ｨ繝ｩ繝ｼ: " & Err.Description, vbExclamation
End Sub


