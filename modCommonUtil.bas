Attribute VB_Name = "modCommonUtil"
Option Explicit

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long

Public Sub Tmp_ShowExcelPID()
    Dim pid As Long
    GetWindowThreadProcessId Application.hwnd, pid
    Debug.Print "[PID] ExcelPID=" & pid & "  Hwnd=" & Application.hwnd
End Sub

Public Sub Cleanup_DuplicateROMHeaders_KeepRightmost()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("EvalData")
    Dim lastCol As Long, c As Long, h As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary") ' key: header, value: Collection of columns

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 1) ROM_* の位置を全走査
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If Len(h) > 0 Then
            If LCase$(Left$(h, 4)) = "rom_" Then
                If Not map.exists(h) Then Set map(h) = New Collection
                map(h).Add c
            End If
        End If
    Next c

    ' 2) 右端(最大列)だけ残し、他は削除対象として収集
    Dim toDel As New Collection
    Dim k As Variant, i As Long, keepCol As Long
    For Each k In map.keys
        If map(k).Count > 1 Then
            keepCol = -1
            For i = 1 To map(k).Count
                If map(k)(i) > keepCol Then keepCol = map(k)(i)
            Next i
            For i = 1 To map(k).Count
                If map(k)(i) <> keepCol Then toDel.Add CLng(map(k)(i))
            Next i
        End If
    Next k

    ' 3) 降順で削除
    Dim arr() As Long, n As Long
    n = toDel.Count
    If n > 0 Then
        ReDim arr(1 To n)
        For i = 1 To n
            arr(i) = toDel(i)
        Next i
        ' 降順ソート
        Dim j As Long, tmp As Long
        For i = 1 To n - 1
            For j = i + 1 To n
                If arr(i) < arr(j) Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            Next j
        Next i
        ' 削除実行
        For i = 1 To n
            Debug.Print "[DUP-CLEAN] delete col", arr(i), "(" & ws.Cells(1, arr(i)).value & ")"
            ws.Columns(arr(i)).Delete
        Next i
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' 見出しが完全一致する「一番右の列番号」を返すユーティリティ
Public Function HeaderCol_Compat_Rightmost(ByVal name As String, ByVal ws As Worksheet) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = lastCol To 1 Step -1
        If StrComp(CStr(ws.Cells(1, c).value), name, vbTextCompare) = 0 Then
            HeaderCol_Compat_Rightmost = c
            Exit Function
        End If
    Next c
End Function

' IO連結文字列 "key=value|key=value|..." から、指定keyの value を返す簡易パーサ
Public Function GetIOValue(ByVal ioStr As String, ByVal key As String) As String
    Dim token As Variant, klen As Long
    klen = Len(key) + 1 ' "key=" の長さ
    For Each token In Split(ioStr, "|")
        If Left$(token, klen) = key & "=" Then
            GetIOValue = Mid$(token, klen + 1) ' "=" の後ろ
            Exit Function
        End If
    Next token
End Function

' "key=value|key:...|" の混在を想定。
' 指定 key の右側（= または : の後ろ）を“そのまま”返す。
Public Function GetIOChunk(ByVal ioStr As String, ByVal key As String) As String
    Dim token As Variant, t As String, k1 As String, k2 As String, p As Long
    k1 = key & "=": k2 = key & ":"
    For Each token In Split(ioStr, "|")
        t = Trim$(CStr(token))
        If Left$(t, Len(k1)) = k1 Then
            GetIOChunk = Mid$(t, Len(k1) + 1)
            Exit Function
        ElseIf Left$(t, Len(k2)) = k2 Then
            GetIOChunk = Mid$(t, Len(k2) + 1)
            Exit Function
        End If
    Next token
End Function

' 例: chunk="R=,L=消失" -> GetIOSubVal(chunk,"R")="" / GetIOSubVal(chunk,"L")="消失"
Public Function GetIOSubVal(ByVal chunk As String, ByVal subkey As String) As String
    Dim parts As Variant, p As Variant, k As String, pos As Long, tail As String, nextPos As Long
    k = subkey & "="
    ' カンマ区切りで走査（R=... , L=... , など）
    parts = Split(chunk, ",")
    For Each p In parts
        p = Trim$(CStr(p))
        If Left$(p, Len(k)) = k Then
            tail = Mid$(p, Len(k) + 1)
            ' もし "R=xxx L=yyy" みたいにカンマ無しで続く場合に備え、次のスペースまでを値とみなす
            nextPos = InStr(1, tail, " ", vbBinaryCompare)
            If nextPos > 0 Then
                GetIOSubVal = Left$(tail, nextPos - 1)
            Else
                GetIOSubVal = tail
            End If
            Exit Function
        End If
    Next p
End Function

Public Sub App_Main()

    ActiveWindow.Zoom = 100
    Application.WindowState = xlMaximized

    Load frmEval

    With frmEval
        .StartUpPosition = 0
        .Left = 0
        .Top = 0

        ' 画面に収まるように上限
        If .Height > Application.UsableHeight + 156 Then .Height = Application.UsableHeight + 156

    End With

    ' ★① まず表示（ここで InsideHeight が確定）
    frmEval.Show vbModeless

        Dim yBtn As Single
    yBtn = frmEval.InsideHeight - frmEval.Controls("btnCloseCtl").Height - 12

    frmEval.Controls("btnCloseCtl").Top = yBtn
    frmEval.Controls("cmdSaveGlobal").Top = yBtn
    frmEval.Controls("cmdClearGlobal").Top = yBtn

    frmEval.Controls("mpPhys").Height = yBtn - frmEval.Controls("mpPhys").Top - 12
    
    Debug.Print "[post-mpPhys] yBtn=" & yBtn & " mpPhysB=" & (frmEval.Controls("mpPhys").Top + frmEval.Controls("mpPhys").Height) & " InsideH=" & frmEval.InsideHeight

    
    
    If frmEval.Height > Application.UsableHeight - 40 Then frmEval.Height = Application.UsableHeight - 40

Call frmEval.AdjustBottomButtons


End Sub

'=== Basic.* と旧列（氏名/評価日 など）を 1 行分だけ同期する =====================
Public Sub SyncBasicInfoColumns(ws As Worksheet, ByVal r As Long)
    Dim headersBasic As Variant
    Dim headersLegacy As Variant
    Dim i As Long
    Dim cB As Long, cL As Long
    Dim vB As Variant, vL As Variant

    ' Basic.* 系を「正」とみなすが、
    ' 片方しか入っていない場合は、入っている方からもう片方へコピーする。
     headersBasic = Array("Basic.EvalDate", "Basic.Name", "Basic.Age", "Basic.Evaluator")
     headersLegacy = Array("評価日", "氏名", "年齢", "評価者")


    For i = LBound(headersBasic) To UBound(headersBasic)
        cB = modEvalIOEntry.FindColByHeaderExact(ws, headersBasic(i))
        cL = modEvalIOEntry.FindColByHeaderExact(ws, headersLegacy(i))


        ' どちらかの列が存在していれば同期対象
        If cB > 0 Or cL > 0 Then
            If cB > 0 Then
                vB = ws.Cells(r, cB).value
            Else
                vB = vbNullString
            End If

            If cL > 0 Then
                vL = ws.Cells(r, cL).value
            Else
                vL = vbNullString
            End If

            ' 優先度：
            ' 1) Basic 側に値があって旧列が空 → Basic → 旧列へコピー
            ' 2) Basic 側が空で旧列に値 → 旧列 → Basic へコピー
            If cB > 0 And Len(vB) > 0 And cL > 0 And Len(vL) = 0 Then
                ws.Cells(r, cL).value = vB
            ElseIf cL > 0 And Len(vL) > 0 And cB > 0 And Len(vB) = 0 Then
                ws.Cells(r, cB).value = vL
            End If
        End If
    Next i
End Sub
'====================================================================


Public Function ControlExists(parent As Object, ctrlName As String) As Boolean
    Dim c As Object
    For Each c In parent.Controls
        If c.name = ctrlName Then
            ControlExists = True
            Exit Function
        End If
    Next
    ControlExists = False
End Function

Public Sub Tighten_DailyLog_Boxes()
    Dim uf As Object: Set uf = frmEval

    Dim mp As Object: Set mp = uf.Controls("MultiPage1")
    Dim pg As Object: Set pg = mp.Pages(7) ' 日々の記録

    Dim f As MSForms.Frame: Set f = pg.Controls("fraDailyLog")
    Dim note As MSForms.TextBox: Set note = pg.Controls("txtDailyNote")
    Dim lst As MSForms.ListBox: Set lst = pg.Controls("lstDailyLogList")

    Const gap As Single = 24
    Const NOTE_H As Single = 180 ' ←ここだけで調整（現状290.4→220）

    ' 記録内容：高さを詰める（スクロールは維持）
    note.multiline = True
    note.ScrollBars = fmScrollBarsVertical
    note.Height = NOTE_H
    
    ' 一覧：上へ詰めて、下端は今のまま（=高さが増える）
    Dim bottomKeep As Single
    bottomKeep = f.Height - 12


    ' ラベルを「一覧の直上」に置く
    Dim lbl As MSForms.label
    Set lbl = pg.Controls("lblDailyHistory")

    Const LBL_GAP As Single = 6

    lbl.Top = note.Top + note.Height + gap
    lst.Top = lbl.Top + lbl.Height + LBL_GAP
    Const LIST_MAX_H As Single = 140   ' ← 好みで調整
    lst.Height = Application.Min(LIST_MAX_H, Application.Max(60, bottomKeep - lst.Top))


    ' ListBoxは溢れたら自動でスクロールが出る（常時表示は仕様上できない）
    lst.IntegralHeight = False


End Sub

Public Function HasControls(ByVal o As Object) As Boolean
    On Error GoTo EH
    Dim n As Long
    n = o.Controls.Count
    HasControls = (n >= 0)
    Exit Function
EH:
    HasControls = False
End Function

Public Sub Verify_POST_TagUniqueness()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim dup As Long: dup = 0

    Dim p As Object, f2 As Object, f35 As Object, f36 As Object
    Set p = frmEval.Controls("MultiPage1").Pages("Page2")
    Set f2 = p.Controls("Frame2")
    Set f35 = f2.Controls("Frame35")
    Set f36 = f2.Controls("Frame36")

    dup = dup + CountDupTagsInFrame(seen, f35)
    dup = dup + CountDupTagsInFrame(seen, f36)

    Debug.Print "[VERIFY POST TAG UNIQUE] DUP=" & dup & " UNIQUE=" & seen.Count
End Sub

Private Function CountDupTagsInFrame(ByVal seen As Object, ByVal fr As Object) As Long
    Dim c As Object, t As String
    Dim dup As Long: dup = 0

    For Each c In fr.Controls
        If c.parent Is fr Then
            If TypeName(c) = "CheckBox" Or TypeName(c) = "ComboBox" Or TypeName(c) = "TextBox" Or TypeName(c) = "OptionButton" Then
                t = ""
                On Error Resume Next
                t = CStr(c.tag)
                On Error GoTo 0

                If Len(t) > 0 Then
                    If seen.exists(t) Then
                        dup = dup + 1
                        Debug.Print "DUP TAG=" & t
                    Else
                        seen.Add t, True
                    End If
                Else
                    Debug.Print "EMPTY TAG at " & TypeName(c) & " L=" & c.Left & " T=" & c.Top
                End If
            End If
        End If
    Next c

    CountDupTagsInFrame = dup
End Function

Public Sub Build_POST_Narrative()
    Dim posture As String, contr As String, note As String, noteC As String

    posture = CollectTrueTags("Posture.")
    contr = CollectTrueTags("Contracture.")

    note = GetTagText("Posture.備考")
    noteC = GetTagText("Contracture.備考")

    If posture = "" Then posture = "姿勢：特記なし" Else posture = "姿勢：" & posture
    If contr = "" Then contr = "拘縮：特記なし" Else contr = "拘縮：" & contr

    Debug.Print posture
    Debug.Print contr
    If note <> "" Then Debug.Print "姿勢備考：" & note
    If noteC <> "" Then Debug.Print "拘縮備考：" & noteC
End Sub

Private Function CollectTrueTags(ByVal prefix As String) As String
    Dim p As Object, f2 As Object, f35 As Object, f36 As Object
    Set p = frmEval.Controls("MultiPage1").Pages("Page2")
    Set f2 = p.Controls("Frame2")
    Set f35 = f2.Controls("Frame35")
    Set f36 = f2.Controls("Frame36")

    Dim res As String
    res = res & CollectInFrame(f35, prefix)
    res = res & CollectInFrame(f36, prefix)

    If Len(res) > 0 Then
        If Right$(res, 1) = "、" Then res = Left$(res, Len(res) - 1)
    End If
    CollectTrueTags = res
End Function

Private Function CollectInFrame(ByVal fr As Object, ByVal prefix As String) As String
    Dim c As Object, t As String, s As String
    For Each c In fr.Controls
        If c.parent Is fr Then
            If TypeName(c) = "CheckBox" Then
                t = CStr(c.tag)
                If Len(t) > 0 And Left$(t, Len(prefix)) = prefix Then
                    If c.value = True Then s = s & Replace$(t, prefix, "") & "、"
                End If
            End If
        End If
    Next
    CollectInFrame = s
End Function

Private Function GetTagText(ByVal tagName As String) As String
    Dim p As Object, f2 As Object, fr As Object, c As Object
    Set p = frmEval.Controls("MultiPage1").Pages("Page2")
    Set f2 = p.Controls("Frame2")

    ' Frame35
    Set fr = f2.Controls("Frame35")
    For Each c In fr.Controls
        If c.parent Is fr Then
            If TypeName(c) = "TextBox" And CStr(c.tag) = tagName Then
                GetTagText = Replace$(CStr(c.Text), vbCrLf, " ")
                Exit Function
            End If
        End If
    Next

    ' Frame36
    Set fr = f2.Controls("Frame36")
    For Each c In fr.Controls
        If c.parent Is fr Then
            If TypeName(c) = "TextBox" And CStr(c.tag) = tagName Then
                GetTagText = Replace$(CStr(c.Text), vbCrLf, " ")
                Exit Function
            End If
        End If
    Next
End Function


Public Function AsLongArray(ByVal v As Variant) As Variant
    If IsEmpty(v) Then Exit Function
    If IsArray(v) Then AsLongArray = v: Exit Function
    If IsObject(v) Then
        Dim i As Long, arr() As Long
        ReDim arr(1 To v.Count)
        For i = 1 To v.Count
            arr(i) = CLng(v(i))
        Next
        AsLongArray = arr
    End If
End Function

Public Function NormalizeName(ByVal s As String) As String
    s = Replace$(CStr(s), " ", "")   ' 半角スペース除去
    s = Replace$(s, "　", "")        ' 全角スペース除去
    NormalizeName = s
End Function



' 氏名を印刷用に伏字化する
' 4?5文字：2・4文字目を〇（ユーザー指定）
Public Function MaskNameForPrint(ByVal s As String, Optional ByVal maskChar As String = "〇") As String
    Dim n As Long, i As Long
    Dim out As String, ch As String

    s = Trim$(s)
    n = Len(s)

    If n <= 0 Then
        MaskNameForPrint = ""
        Exit Function
    End If

    For i = 1 To n
        ch = Mid$(s, i, 1)

        Select Case n
            Case 1
                out = out & ch
            Case 2
                If i = 2 Then out = out & maskChar Else out = out & ch
            Case 3
                If i = 2 Then out = out & maskChar Else out = out & ch
            Case 4, 5
                If (i = 2) Or (i = 4) Then out = out & maskChar Else out = out & ch
            Case Else
                ' 6文字以上：偶数位置を伏字（2,4,6,...）
                If (i Mod 2 = 0) Then out = out & maskChar Else out = out & ch
        End Select
    Next i

    MaskNameForPrint = out
End Function



' 月単位のブックに、利用者シートを追加していく
' ymKey 例: "2026-02" など（月単位で一意になる文字列）
Public Sub ExportMonitoring_ToMonthlyWorkbook(ByVal dailyDate As Date, ByVal clientName As String, ByVal bodyText As String)
    
    If Len(Trim$(clientName)) = 0 Then
        clientName = frmEval.Controls("frHeader").Controls("txtHdrName").Text
    End If

    
    
    
    Dim ymKey As String: ymKey = Format(dailyDate, "yyyy-mm")
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim saveName As String
    Dim savePath As String


    saveName = "Monitoring_" & ymKey & ".xlsx"

    Dim baseFolder As String
Dim monthFolder As String

baseFolder = ThisWorkbook.path & "\Monitoring"
monthFolder = baseFolder & "\" & ymKey

If Dir$(baseFolder, vbDirectory) = "" Then
    MkDir baseFolder
End If

If Dir$(monthFolder, vbDirectory) = "" Then
    MkDir monthFolder
End If

savePath = monthFolder & "\" & saveName



If Len(Dir$(savePath)) > 0 Then
    Set wbNew = Workbooks.Open(savePath)
Else
    ThisWorkbook.Worksheets("Monitoring").Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
End If


    ' 既に同名シートがあれば削除して作り直し（上書き運用）
    If Len(Trim$(clientName)) > 0 Then
    On Error Resume Next
    Application.DisplayAlerts = False
    wbNew.Worksheets(clientName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End If


    ' テンプレ Monitoring をこの月ブックへコピー
    ThisWorkbook.Worksheets("Monitoring").Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    Set wsNew = wbNew.Worksheets(wbNew.Worksheets.Count)
    
    
    Dim c As Long, firstCol As Long, lastCol As Long
With ThisWorkbook.Worksheets("Monitoring").UsedRange
    firstCol = .Column
    lastCol = .Column + .Columns.Count - 1
End With

With ThisWorkbook.Worksheets("Monitoring")
    For c = firstCol To lastCol
        wsNew.Columns(c).ColumnWidth = .Columns(c).ColumnWidth
    Next c
End With





    ' シート名は実名でOK
    wsNew.name = Left$(clientName, 31)

    ' 印刷物の表示名だけ伏字（シート名は実名）
    wsNew.Range("C7").value = MaskNameForPrint(clientName)

    ' 本文
    wsNew.Range("A32:J37").Merge
    Dim p As Long
Dim s As String

s = bodyText
p = InStr(1, s, "■ コメント・考察", vbTextCompare)

If p > 0 Then
    s = Mid$(s, p + Len("■ コメント・考察"))
    s = Trim$(s)
Else
    s = bodyText ' 見つからない時は保険で全文
End If


Dim p1 As Long, p2 As Long
Dim tok As String, special As String

tok = "■ この月に記録された特記事項"
p1 = InStr(1, bodyText, tok, vbTextCompare)

If p1 > 0 Then
    special = Mid$(bodyText, p1 + Len(tok))
    ' 次の見出しで止める（候補）
    p2 = InStr(1, special, "■ コメント・考察", vbTextCompare)
    If p2 = 0 Then p2 = InStr(1, special, "■ 本文", vbTextCompare)
    If p2 > 0 Then special = Left$(special, p2 - 1)
    special = Trim$(special)
Else
    special = ""
End If



Dim tmplNoSpecial As String
tmplNoSpecial = "この月は特記事項となる記録はありませんでした。" & vbCrLf & _
                "体調面に大きな変動はなく、日々のリハビリにも安定して取り組まれていました。" & vbCrLf & _
                "今後も現在の状態を維持できるよう、引き続き経過を観察していきます。"

If Len(Trim$(special)) = 0 Or InStr(special, "特記事項となる記録はありません") > 0 Then
    
    ' 特記事項なし
    wsNew.Range("A24").value = tmplNoSpecial
    wsNew.Range("A31:J37").ClearContents

Else
    
    ' 特記事項あり
    wsNew.Range("A24").value = special
    wsNew.Range("A31").value = s

End If



wsNew.Range("A24").WrapText = True
With wsNew.Range("A24:J29")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With






With wsNew.Range("A31:J37")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With


    'wsNew.Range("A32").value = bodyText
    
    wsNew.Range("A24:J29").Merge
    With wsNew.Range("A31: J37 ")
      .Font.name = "ＭＳ Ｐゴシック"
      .Font.Size = 11
    End With
    
    
   Application.PrintCommunication = False

 With wsNew.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .Zoom = 100

 End With
Application.PrintCommunication = True


    

    
    wsNew.Range("A31").WrapText = True
    
    wsNew.ResetAllPageBreaks

    Dim r As Long, lastRow As Long
With ThisWorkbook.Worksheets("Monitoring").UsedRange
    lastRow = .row + .rows.Count - 1
End With

With ThisWorkbook.Worksheets("Monitoring")
    For r = 1 To lastRow
        wsNew.rows(r).RowHeight = .rows(r).RowHeight
    Next r
End With

With wsNew.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .Zoom = False
End With




    wbNew.Save
wbNew.Close SaveChanges:=True



End Sub

