Attribute VB_Name = "modEvalEntry"
' 参照設定：Microsoft Visual Basic for Applications Extensibility 5.3（必要なら）
' ファイル→オプション→マクロの設定で「VBAプロジェクト オブジェクトモデルへの信頼を許可」にチェック

Option Explicit


' 必須の定義（プロジェクトに合わせて調整）
Public Const FORM_MAIN As String = "frmEval"
Public Const HOST_MOVE_NAME As String = "hostMove"   ' 日常生活動作のコンテナFrameのName
Public Const MP_ADL_NAME As String = "mpADL"
Public Const CAP_BI As String = "バーサルインデックス"
Public Const CAP_IADL As String = "IADL"
Public Const CAP_KYO As String = "起居動作"



Public Sub ProjectMap_ToSheet()
    Dim wb As Workbook
Dim ws As Worksheet
    On Error Resume Next
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("PROJECT_MAP")
    If Not ws Is Nothing Then Application.DisplayAlerts = False: ws.Delete: Application.DisplayAlerts = True
    Set ws = wb.Worksheets.Add
    ws.name = "PROJECT_MAP"
    On Error GoTo 0

    Dim r As Long: r = 1
    ws.Cells(r, 1).value = "Project Map (" & Format(Now, "yyyy-mm-dd hh:nn:ss") & ")": r = r + 2

    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ws.Cells(r, 1).value = TypeName(comp)
        ws.Cells(r, 2).value = comp.name
        r = r + 1

        'UserFormならコントロールを列挙（デザイン時）
        If comp.Type = 3 Then ' vbext_ct_MSForm
            On Error Resume Next
            Dim d As Object, i As Long
            Set d = comp.Designer
            If Not d Is Nothing Then
                For i = 0 To d.Controls.Count - 1
                    ws.Cells(r, 2).value = "Ctrl"
                    ws.Cells(r, 3).value = TypeName(d.Controls(i))
                    ws.Cells(r, 4).value = d.Controls(i).name
                    On Error Resume Next
                    ws.Cells(r, 5).value = d.Controls(i).caption
                    On Error GoTo 0
                    r = r + 1
                Next
            End If
            On Error GoTo 0
        End If

        r = r + 1
    Next

    ws.Columns.AutoFit
    MsgBox "PROJECT_MAP シートを作成しました。", vbInformation
End Sub








' 標準モジュール（Option Explicit のままでOK）

Public Sub Validate_App()
    Dim ok As Boolean: ok = True

    ' 1) フォームを一時生成
    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add(FORM_MAIN)
    On Error GoTo 0
    If frm Is Nothing Then
        MsgBox "UserForm '" & FORM_MAIN & "' が見つかりません", vbCritical
        Exit Sub
    End If

    ' 2) hostMove を取得（Frame）
    Dim c As Object
    Dim hostMove As Object: Set hostMove = Nothing
    For Each c In frm.Controls
        If TypeName(c) = "Frame" Then
            If StrComp(c.name, HOST_MOVE_NAME, vbTextCompare) = 0 Then
                Set hostMove = c
                Exit For
            End If
        End If
    Next
    If hostMove Is Nothing Then
        MsgBox "Frame '" & HOST_MOVE_NAME & "' が " & FORM_MAIN & " に見つかりません", vbCritical
        Unload frm: Exit Sub
    End If

    ' 3) mpADL を取得（MultiPage）
    Dim mp As Object: Set mp = Nothing
    For Each c In hostMove.Controls
        If TypeName(c) = "MultiPage" Then
            If StrComp(c.name, MP_ADL_NAME, vbTextCompare) = 0 Then
                Set mp = c
                Exit For
            End If
        End If
    Next

    ' 4) チェック
    If mp Is Nothing Then
        ok = False
        Debug.Print "[NG] mpADL がありません"
    Else
        If mp.Pages.Count < 3 Then ok = False: Debug.Print "[NG] mpADL Pages.Count < 3"
        If mp.Pages.Count >= 1 Then If mp.Pages(0).caption <> CAP_BI Then ok = False: Debug.Print "[NG] Page0 Caption≠" & CAP_BI
        If mp.Pages.Count >= 2 Then If mp.Pages(1).caption <> CAP_IADL Then ok = False: Debug.Print "[NG] Page1 Caption≠" & CAP_IADL
        If mp.Pages.Count >= 3 Then If mp.Pages(2).caption <> CAP_KYO Then ok = False: Debug.Print "[NG] Page2 Caption≠" & CAP_KYO
    End If

    MsgBox IIf(ok, "Validate OK：構成は想定どおりです。", _
                    "Validate NG：Immediate ウィンドウ（Ctrl+G）を確認してください。"), _
          IIf(ok, vbInformation, vbExclamation)

    Unload frm
End Sub


