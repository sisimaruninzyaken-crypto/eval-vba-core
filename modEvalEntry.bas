Attribute VB_Name = "modEvalEntry"
' 蜿ら・險ｭ螳夲ｼ哺icrosoft Visual Basic for Applications Extensibility 5.3・亥ｿ・ｦ√↑繧会ｼ・
' 繝輔ぃ繧､繝ｫ竊偵が繝励す繝ｧ繝ｳ竊偵・繧ｯ繝ｭ縺ｮ險ｭ螳壹〒縲祁BA繝励Ο繧ｸ繧ｧ繧ｯ繝・繧ｪ繝悶ず繧ｧ繧ｯ繝医Δ繝・Ν縺ｸ縺ｮ菫｡鬆ｼ繧定ｨｱ蜿ｯ縲阪↓繝√ぉ繝・け

Option Explicit


' 蠢・医・螳夂ｾｩ・医・繝ｭ繧ｸ繧ｧ繧ｯ繝医↓蜷医ｏ縺帙※隱ｿ謨ｴ・・
Public Const FORM_MAIN As String = "frmEval"
Public Const HOST_MOVE_NAME As String = "hostMove"   ' 譌･蟶ｸ逕滓ｴｻ蜍穂ｽ懊・繧ｳ繝ｳ繝・リFrame縺ｮName
Public Const MP_ADL_NAME As String = "mpADL"
Public Const CAP_BI As String = "繝舌・繧ｵ繝ｫ繧､繝ｳ繝・ャ繧ｯ繧ｹ"
Public Const CAP_IADL As String = "IADL"
Public Const CAP_KYO As String = "襍ｷ螻・虚菴・



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

        'UserForm縺ｪ繧峨さ繝ｳ繝医Ο繝ｼ繝ｫ繧貞・謖呻ｼ医ョ繧ｶ繧､繝ｳ譎ゑｼ・
        If comp.Type = 3 Then ' vbext_ct_MSForm
            On Error Resume Next
            Dim d As Object, i As Long
            Set d = comp.Designer
            If Not d Is Nothing Then
                For i = 0 To d.controls.count - 1
                    ws.Cells(r, 2).value = "Ctrl"
                    ws.Cells(r, 3).value = TypeName(d.controls(i))
                    ws.Cells(r, 4).value = d.controls(i).name
                    On Error Resume Next
                    ws.Cells(r, 5).value = d.controls(i).caption
                    On Error GoTo 0
                    r = r + 1
                Next
            End If
            On Error GoTo 0
        End If

        r = r + 1
    Next

    ws.Columns.AutoFit
    MsgBox "PROJECT_MAP 繧ｷ繝ｼ繝医ｒ菴懈・縺励∪縺励◆縲・, vbInformation
End Sub








' 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ・・ption Explicit 縺ｮ縺ｾ縺ｾ縺ｧOK・・

Public Sub Validate_App()
    Dim ok As Boolean: ok = True

    ' 1) 繝輔か繝ｼ繝繧剃ｸ譎ら函謌・
    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add(FORM_MAIN)
    On Error GoTo 0
    If frm Is Nothing Then
        MsgBox "UserForm '" & FORM_MAIN & "' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ", vbCritical
        Exit Sub
    End If

    ' 2) hostMove 繧貞叙蠕暦ｼ・rame・・
    Dim c As Object
    Dim hostMove As Object: Set hostMove = Nothing
    For Each c In frm.controls
        If TypeName(c) = "Frame" Then
            If StrComp(c.name, HOST_MOVE_NAME, vbTextCompare) = 0 Then
                Set hostMove = c
                Exit For
            End If
        End If
    Next
    If hostMove Is Nothing Then
        MsgBox "Frame '" & HOST_MOVE_NAME & "' 縺・" & FORM_MAIN & " 縺ｫ隕九▽縺九ｊ縺ｾ縺帙ｓ", vbCritical
        Unload frm: Exit Sub
    End If

    ' 3) mpADL 繧貞叙蠕暦ｼ・ultiPage・・
    Dim mp As Object: Set mp = Nothing
    For Each c In hostMove.controls
        If TypeName(c) = "MultiPage" Then
            If StrComp(c.name, MP_ADL_NAME, vbTextCompare) = 0 Then
                Set mp = c
                Exit For
            End If
        End If
    Next

    ' 4) 繝√ぉ繝・け
    If mp Is Nothing Then
        ok = False
        Debug.Print "[NG] mpADL 縺後≠繧翫∪縺帙ｓ"
    Else
        If mp.Pages.count < 3 Then ok = False: Debug.Print "[NG] mpADL Pages.Count < 3"
        If mp.Pages.count >= 1 Then If mp.Pages(0).caption <> CAP_BI Then ok = False: Debug.Print "[NG] Page0 Caption竕" & CAP_BI
        If mp.Pages.count >= 2 Then If mp.Pages(1).caption <> CAP_IADL Then ok = False: Debug.Print "[NG] Page1 Caption竕" & CAP_IADL
        If mp.Pages.count >= 3 Then If mp.Pages(2).caption <> CAP_KYO Then ok = False: Debug.Print "[NG] Page2 Caption竕" & CAP_KYO
    End If

    MsgBox IIf(ok, "Validate OK・壽ｧ区・縺ｯ諠ｳ螳壹←縺翫ｊ縺ｧ縺吶・, _
                    "Validate NG・唔mmediate 繧ｦ繧｣繝ｳ繝峨え・・trl+G・峨ｒ遒ｺ隱阪＠縺ｦ縺上□縺輔＞縲・), _
          IIf(ok, vbInformation, vbExclamation)

    Unload frm
End Sub


