Attribute VB_Name = "modScan"
Option Explicit

' === 文字列中の出現回数を数える（大文字小文字を無視） ===
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

' === プロジェクト全体を走査してマップを作る（信頼アクセスが必要） ===
Public Sub MakeProjectMap()
    ' 必要設定: [ファイル]→[オプション]→[セキュリティ センター]→
    '  [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] にチェック
    On Error GoTo EH

    Dim wb As Workbook, sh As Worksheet
    Set wb = ThisWorkbook

    ' シート準備
    Const MAP_NAME As String = "PROJECT_MAP"
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(MAP_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set sh = wb.Worksheets.Add
    sh.name = MAP_NAME

    ' 見出し
    Dim h As Variant
    h = Array("Component", "Type", "Lines", _
              "mpADL", "EnsureBI_IADL", "BuildKyoOnADL", "RemoveAllMpADL", _
              "CAP_BI", "CAP_IADL", "CAP_KYO", "hostMove", "nextTop")
    Dim i As Long
    For i = LBound(h) To UBound(h)
        sh.Cells(1, i + 1).value = h(i)
    Next

    ' 参照は late binding（Extensibility 参照なしで動く）
    Dim vbProj As Object: Set vbProj = wb.VBProject
    Dim comp As Object, cm As Object
    Dim r As Long: r = 2

    ' 型定数（参照なし対応）
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

        ' キー語の出現回数
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

    ' 体裁
    With sh
        .rows(1).Font.Bold = True
        .Columns.AutoFit
    End With

    MsgBox "PROJECT_MAP を作成しました。", vbInformation
    Exit Sub
EH:
    MsgBox "MakeProjectMap エラー: " & Err.Description, vbExclamation
End Sub

' === どこで mpADL を生成しているかの詳細一覧（行番号付き） ===
Public Sub FindMpADLCreates()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim vbProj As Object: Set vbProj = wb.VBProject
    Dim comp As Object, cm As Object
    Debug.Print String(60, "-")
    Debug.Print "[SCAN] mpADL を Add/Set している行の一覧"
    For Each comp In vbProj.VBComponents
        Set cm = comp.CodeModule
        Dim n As Long: n = cm.CountOfLines
        Dim i As Long
        For i = 1 To n
            Dim ln As String: ln = cm.lines(i, 1)
            ' 生成/代入っぽい行を拾う（ざっくり）
            If InStr(1, LCase$(ln), "set mpadl") > 0 Or _
               InStr(1, LCase$(ln), "controls.add(""forms.multipage.1""") > 0 Then
                Debug.Print comp.name & ":" & i & "  " & Trim$(ln)
            End If
        Next
    Next
    Debug.Print String(60, "-")
    MsgBox "Immediate ウィンドウ（Ctrl+G）に出力しました。", vbInformation
    Exit Sub
EH:
    MsgBox "FindMpADLCreates エラー: " & Err.Description, vbExclamation
End Sub


