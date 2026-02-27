Attribute VB_Name = "modEvalSupplementIO"
Option Explicit
Public gArchiveDeleteBasicID As String



'--- ヘッダ行から列番号を探す（見つからなければ 0）
Private Function HeaderCol(ByVal hdrRow As Range, ByVal candidates As Variant) As Long
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim c As Range
        For Each c In hdrRow.Cells
            If Trim$(CStr(c.value)) = CStr(candidates(i)) Then
                HeaderCol = c.Column - hdrRow.Cells(1, 1).Column + 1 'CurrentRegion内の相対列
                Exit Function
            End If
        Next c
    Next i
    HeaderCol = 0
End Function

'--- アーカイブブックに同名シートがなければ作る
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        On Error Resume Next
        ws.name = sheetName
        On Error GoTo 0
    End If
    Set GetOrCreateSheet = ws
End Function



Public Sub DeleteClientFrom_EvalData_ByName()
    Dim nm As String
    nm = Trim$(InputBox("削除したい利用者の氏名（完全一致）", "EvalData 削除", ""))
    If nm = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim nameCol As Long: nameCol = 89 'CK

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, nameCol).End(xlUp).row

    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, nameCol).value) = nm Then
            ws.rows(r).Delete
            cnt = cnt + 1
        End If
    Next r

    MsgBox "EvalData: " & cnt & " 行を削除しました。", vbInformation
End Sub



' 同姓同名時：削除ボタン運用では EvalData から最新のID(空でない)を自動補助取得するため、通常はID入力を求めない

' EvalData 専用：氏名（CK=89列）一致の行を「別ブックへ退避」→「元から削除」
Public Sub ArchiveAndDelete_EvalData_ByName()

    Dim nm As String
    nm = Trim$(InputBox("EvalData：退避→削除したい氏名（完全一致）" & vbCrLf & "※同姓同名がいる場合は次にIDを聞きます", "EvalData 利用終了者アーカイブ", ""))
    If nm = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Const NAME_COL As Long = 89 'CK（氏名）

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, NAME_COL).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "EvalData にデータ行がありません。", vbExclamation
        Exit Sub
    End If
    
    
    Const ID_COL As Long = 82 'Basic.ID


'--- 同姓同名チェック：複数ヒットならIDを追加で聞く ---
Dim hitCount As Long: hitCount = 0
Dim r2 As Long
For r2 = lastRow To 2 Step -1
    If CStr(ws.Cells(r2, NAME_COL).value) = nm Then hitCount = hitCount + 1
Next r2

Dim pid As String: pid = ""
If hitCount >= 2 Then
    If Len(gArchiveDeleteBasicID) > 0 Then pid = gArchiveDeleteBasicID
        If pid = "" Then
    End If
    
    
    
    If pid = "" Then
        pid = Trim$(InputBox("同姓同名が見つかりました。削除（退避）したいIDを入力してください。", "IDで特定", ""))
        If pid = "" Then
            MsgBox "ID未入力のため中止しました。", vbExclamation
            Exit Sub
        End If
    End If
End If


    

    Dim ans As VbMsgBoxResult
    ans = MsgBox("EvalData の氏名=" & nm & " をアーカイブへ退避し、元データから削除します。実行しますか？", _
                 vbYesNo + vbQuestion, "最終確認")
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim wbArc As Workbook: Set wbArc = Workbooks.Add(xlWBATWorksheet)
    Dim wsA As Worksheet: Set wsA = wbArc.Worksheets(1)
    wsA.name = "EvalData"

    ' ヘッダ退避（A1:FW1 をそのまま）
    ws.Range("A1:FW1").Copy Destination:=wsA.Range("A1")

    Dim moved As Long: moved = 0
    Dim r As Long

    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, NAME_COL).value) = nm And (pid = "" Or CStr(ws.Cells(r, ID_COL).value) = pid) Then
            ' 行退避（A:FW の行）
            Dim nextA As Long
            nextA = wsA.Cells(wsA.rows.Count, 1).End(xlUp).row + 1
            ws.Range("A" & r & ":FW" & r).Copy Destination:=wsA.Range("A" & nextA)

            ' 元から削除
            ws.rows(r).Delete
            moved = moved + 1
        End If
    Next r

    ' アーカイブ保存
    Dim arcPath As String, arcFile As String
    arcPath = ThisWorkbook.path
    If arcPath = "" Then arcPath = Environ$("TEMP")
    arcFile = arcPath & Application.PathSeparator & "EvalData_Archive_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    wbArc.SaveAs Filename:=arcFile, FileFormat:=xlOpenXMLWorkbook 'xlsx

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "完了：EvalData から " & moved & " 行を退避→削除しました。" & vbCrLf & _
           "アーカイブ：" & arcFile, vbInformation
           
           
           gArchiveDeleteBasicID = ""

End Sub




Public Sub AddHeaderArchiveDeleteButton()
    Dim f As Object: Set f = frmEval
    Dim hdr As MSForms.Frame: Set hdr = f.Controls("frHeader")

    Dim btn As MSForms.CommandButton

    On Error Resume Next
    Set btn = hdr.Controls("cmdArchiveDelete")
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = hdr.Controls.Add("Forms.CommandButton.1", "cmdArchiveDelete", True)
        btn.caption = "終了者削除"
        btn.Width = 90
        btn.Height = hdr.Controls("txtHdrPID").Height
        btn.Top = hdr.Controls("txtHdrPID").Top
        btn.Left = 8
    End If
End Sub





Public Function GetLatestID_ForName(ByVal ws As Worksheet, ByVal nm As String, ByVal nameCol As Long, ByVal idCol As Long) As String
    Dim lastRow As Long, r As Long
    Dim idS As String
    
    lastRow = ws.Cells(ws.rows.Count, nameCol).End(xlUp).row
    
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, nameCol).value) = nm Then
            idS = Trim$(CStr(ws.Cells(r, idCol).value))
            If idS <> "" Then
                GetLatestID_ForName = idS
                Exit Function
            End If
        End If
    Next r
End Function





