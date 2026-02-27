Attribute VB_Name = "modPostureIO"

' ===== modPostureIO.bas（複数項目＋備考 版）=====
Option Explicit

Public Sub SavePostureToSheet(ws As Worksheet, ByVal r As Long, owner As Object)
    Dim caps As Variant, i As Long, cap As String, col As Long
    caps = PostureCaptions()
    For i = LBound(caps) To UBound(caps)
        cap = CStr(caps(i))
        col = EnsureHeaderCol_Posture(ws, "姿勢_" & cap)
        ws.Cells(r, col).value = GetCheckByCaption(owner, cap)
        Debug.Print "[SAVE][Posture]", cap, "=", ws.Cells(r, col).value
    Next

   ' 骨盤傾斜（コンボ）?→ 枠「姿勢評価」内のラベル「骨盤傾斜」に紐づくComboを拾う
Dim cPel As Long, pel As String
cPel = EnsureHeaderCol_Posture(ws, "姿勢_骨盤傾斜")
pel = GetComboInFrameByLabelCaption_(owner, "姿勢評価", "骨盤傾斜")
ws.Cells(r, cPel).value = pel
Debug.Print "[SAVE][Posture] 骨盤傾斜 =", pel

' 上段 備考（姿勢の備考）?→ 枠「姿勢評価」内のラベル「備考」に紐づくTextBoxを拾う
Dim cNote As Long, noteVal As String
cNote = EnsureHeaderCol_Posture(ws, "姿勢_備考")
noteVal = GetTextInFrameByLabelCaption_(owner, "姿勢評価", "備考")
ws.Cells(r, cNote).value = noteVal
Debug.Print "[SAVE][Posture] 備考 =", noteVal


   
   ' ?? 関節拘縮：頸部（左右なし）
Dim colNeck As Long
colNeck = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_頸部")
ws.Cells(r, colNeck).value = GetCheckInFrameByCaptionLike_(owner, "関節拘縮", "頸部")
Debug.Print "[SAVE][Posture] 拘縮_頸部 =", ws.Cells(r, colNeck).value

   

    ' ?? 関節拘縮：6部位（右/左） ??
Dim colJ As Long, joints As Variant, jr As String

joints = Array("肩関節", "肘関節", "手関節", "股関節", "膝関節", "足関節")

For i = LBound(joints) To UBound(joints)
    jr = CStr(joints(i))
    colJ = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_" & Replace(jr, "関節", "") & "_右")
    ws.Cells(r, colJ).value = GetKoushuku_OnRow_(owner, jr, "右")

    colJ = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_" & Replace(jr, "関節", "") & "_左")
    ws.Cells(r, colJ).value = GetKoushuku_OnRow_(owner, jr, "左")
Next


    ' 備考（姿勢ブロック）
    Dim cKNote As Long, kNote As String
cKNote = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_備考")
kNote = GetTextInFrameByLabelCaption_(owner, "関節拘縮", "備考")
ws.Cells(r, cKNote).value = kNote
Debug.Print "[SAVE][Posture] 拘縮_備考 =", kNote

End Sub



Public Sub LoadPostureFromSheet(ws As Worksheet, ByVal r As Long, owner As Object)
    
    Debug.Print "[POSTURE][ENTER] r=" & r
    
    Dim caps As Variant, i As Long, cap As String, col As Long, v As Variant
    caps = PostureCaptions()
    For i = LBound(caps) To UBound(caps)
        cap = CStr(caps(i))
        col = EnsureHeaderCol_Posture(ws, "姿勢_" & cap)
        v = ws.Cells(r, col).value
        



        SetCheckByCaption owner, cap, CBool(v)
        
    Next

    ' 骨盤傾斜（コンボ）
Dim cPel As Long, vPel As Variant
cPel = EnsureHeaderCol_Posture(ws, "姿勢_骨盤傾斜")
vPel = ws.Cells(r, cPel).value
SetComboInFrameByLabelCaption_ owner, "姿勢評価", "骨盤傾斜", CStr(vPel)

' 上段 備考（姿勢の備考）
Dim cNote As Long, vNote As Variant
cNote = EnsureHeaderCol_Posture(ws, "姿勢_備考")
vNote = ws.Cells(r, cNote).value
SetTextInFrameByLabelCaption_ owner, "姿勢評価", "備考", CStr(vNote)


    
    ' ?? 関節拘縮：頸部（左右なし）
Dim colNeck As Long, vNeck As Variant
colNeck = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_頸部")
vNeck = ws.Cells(r, colNeck).value
SetCheckInFrameByCaptionLike_ owner, "関節拘縮", "頸部", CBool(vNeck)


    

    ' ?? 関節拘縮：6部位（右/左） ??
Dim colJ As Long, joints As Variant, jr As String

joints = Array("肩関節", "肘関節", "手関節", "股関節", "膝関節", "足関節")

For i = LBound(joints) To UBound(joints)
    jr = CStr(joints(i))

    colJ = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_" & Replace(jr, "関節", "") & "_右")
    v = ws.Cells(r, colJ).value
    SetKoushuku_OnRow_ owner, jr, "右", CBool(v)

    colJ = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_" & Replace(jr, "関節", "") & "_左")
    v = ws.Cells(r, colJ).value
    SetKoushuku_OnRow_ owner, jr, "左", CBool(v)
Next i

    ' 備考（姿勢ブロック）
    Dim cKNote As Long, vKNote As Variant
cKNote = EnsureHeaderCol_Posture(ws, "姿勢_拘縮_備考")
vKNote = ws.Cells(r, cKNote).value
SetTextInFrameByLabelCaption_ owner, "関節拘縮", "備考", CStr(vKNote)


Debug.Print "[POSTURE][EXIT] r=" & r

End Sub



' ====== ここからヘルパ ======

' 扱うチェック項目名（キャプション）をここに追加していく

Private Function PostureCaptions() As Variant
    PostureCaptions = Array("頭部前方突出", "円背", "側弯", "体幹回旋", "反張膝")
End Function


' --- CheckBox（Caption一致）取得/設定 ---
Private Function GetCheckByCaption(owner As Object, ByVal cap As String) As Boolean
    Dim chk As Object
    Set chk = FindCheckByCaptionLike_(owner, cap)
    If Not chk Is Nothing Then GetCheckByCaption = (chk.value <> 0)
End Function

Private Sub SetCheckByCaption(owner As Object, ByVal cap As String, ByVal v As Boolean)
    Dim chk As Object
    Set chk = FindCheckByCaptionLike_(owner, cap)
    If Not chk Is Nothing Then chk.value = v
End Sub

' --- Captionでチェックボックスを探す（入れ子再帰・部分一致OK）---
Private Function FindCheckByCaptionLike_(container As Object, ByVal needle As String) As Object
    Dim c As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(needle), vbTextCompare) > 0 Then
                Set FindCheckByCaptionLike_ = c: Exit Function
            End If
        End If
        If HasControls__(c) Then
            Set FindCheckByCaptionLike_ = FindCheckByCaptionLike_(c, needle)
            If Not FindCheckByCaptionLike_ Is Nothing Then Exit Function
        End If
    Next
End Function

' --- 備考：ラベル「備考」と同じ親内のTextBoxを拾う ---
Private Function GetTextByLabelCaption(owner As Object, ByVal labelCap As String) As String
    Dim tb As Object: Set tb = FindTextBoxNearLabel_(owner, labelCap)
    If Not tb Is Nothing Then GetTextByLabelCaption = CStr(tb.Text)
End Function

Private Sub SetTextByLabelCaption(owner As Object, ByVal labelCap As String, ByVal s As String)
    Dim tb As Object: Set tb = FindTextBoxNearLabel_(owner, labelCap)
    If Not tb Is Nothing Then tb.Text = s
End Sub

' ラベルのCaption一致→同じ親（Frameなど）内のTextBoxを返す（最初の1つ）
Private Function FindTextBoxNearLabel_(container As Object, ByVal labelCap As String) As Object
    Dim c As Object, inner As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                ' 同じ親内のTextBoxを返す
                For Each inner In c.parent.Controls
                    If TypeName(inner) = "TextBox" Then Set FindTextBoxNearLabel_ = inner: Exit Function
                Next
            End If
        End If
        If HasControls__(c) Then
            Set FindTextBoxNearLabel_ = FindTextBoxNearLabel_(c, labelCap)
            If Not FindTextBoxNearLabel_ Is Nothing Then Exit Function
        End If
    Next
End Function

' 子を持つかの判定（安全版）
Private Function HasControls__(obj As Object) As Boolean
    On Error Resume Next
    HasControls__ = (obj.Controls.Count >= 0)
End Function

' --- 姿勢用ローカル: 見出し列を確保して列番号を返す ---
Private Function EnsureHeaderCol_Posture(ws As Worksheet, ByVal header As String) As Long
    Dim c As Range
    Set c = Nothing
    On Error Resume Next
    Set c = ws.rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If Not c Is Nothing Then EnsureHeaderCol_Posture = c.Column: Exit Function

    Dim lastCol As Long
    If Application.WorksheetFunction.CountA(ws.rows(1)) = 0 Then
        lastCol = 0
    Else
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If
    ws.Cells(1, lastCol + 1).value = header
    EnsureHeaderCol_Posture = lastCol + 1
End Function


Private Function GetComboByLabelCaption(owner As Object, ByVal labelCap As String) As String
    Dim cmb As Object: Set cmb = FindFirstByTypeNearLabel_(owner, labelCap, "ComboBox")
    If Not cmb Is Nothing Then GetComboByLabelCaption = CStr(cmb.value)
End Function

Private Sub SetComboByLabelCaption(owner As Object, ByVal labelCap As String, ByVal s As String)
    Dim cmb As Object: Set cmb = FindFirstByTypeNearLabel_(owner, labelCap, "ComboBox")
    If Not cmb Is Nothing Then cmb.value = s
End Sub

Private Function FindFirstByTypeNearLabel_(container As Object, ByVal labelCap As String, ByVal wantType As String) As Object
    Dim c As Object, inner As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "Label" And InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
            For Each inner In c.parent.Controls
                If TypeName(inner) = wantType Then Set FindFirstByTypeNearLabel_ = inner: Exit Function
            Next
        End If
        If HasControls__(c) Then
            Set FindFirstByTypeNearLabel_ = FindFirstByTypeNearLabel_(c, labelCap, wantType)
            If Not FindFirstByTypeNearLabel_ Is Nothing Then Exit Function
        End If
    Next
End Function



' ==== 関節拘縮（行ラベルの「右／左」チェック）ヘルパ ====

' ラベルCaption＝部位名（例「肩関節」）と同じ行の「右／左」CheckBoxを探す（Topが近いものを採用）
Private Function FindSideCheck_OnSameRow_(owner As Object, ByVal rowLabel As String, ByVal sideCaption As String) As Object
    Dim lbl As Object, p As Object, c As Object
    Dim best As Object, bestDy As Single, dy As Single, tol As Single
    tol = 18 ' 同じ行とみなすTop距離（フォームのスケールにより調整可）

    Set lbl = FindLabelByCaptionDeep_(owner, rowLabel)
    If lbl Is Nothing Then Exit Function

    Set p = lbl.parent
    bestDy = 1E+20
    For Each c In p.Controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(sideCaption), vbTextCompare) > 0 Then
                dy = Abs(CSng(c.Top) - CSng(lbl.Top))
                If dy < bestDy And dy <= tol Then
                    Set best = c
                    bestDy = dy
                End If
            End If
        End If
    Next
    Set FindSideCheck_OnSameRow_ = best
End Function

' ラベル（部位名）を深く探す
Private Function FindLabelByCaptionDeep_(container As Object, ByVal cap As String) As Object
    Dim c As Object, r As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(cap), vbTextCompare) > 0 Then
                Set FindLabelByCaptionDeep_ = c: Exit Function
            End If
        End If
        If HasControls__(c) Then
            Set r = FindLabelByCaptionDeep_(c, cap)
            If Not r Is Nothing Then Set FindLabelByCaptionDeep_ = r: Exit Function
        End If
    Next
End Function

' 右左の値取得・設定
Private Function GetKoushuku_OnRow_(owner As Object, ByVal rowLabel As String, ByVal side As String) As Boolean
    Dim chk As Object
    Set chk = FindSideCheck_OnSameRow_(owner, rowLabel, side)
    If Not chk Is Nothing Then GetKoushuku_OnRow_ = (chk.value <> 0)
End Function

Private Sub SetKoushuku_OnRow_(owner As Object, ByVal rowLabel As String, ByVal side As String, ByVal v As Boolean)
    Dim chk As Object
    Set chk = FindSideCheck_OnSameRow_(owner, rowLabel, side)
    If Not chk Is Nothing Then chk.value = v
End Sub



' === Frame（枠）をCaptionで特定して中のコントロールを扱う ===

' 枠Captionに部分一致するFrameを深く探す
Private Function FindFrameByCaptionDeep_(container As Object, ByVal capLike As String) As Object
    Dim c As Object, r As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "Frame" Then
            If InStr(1, Trim$(c.caption), Trim$(capLike), vbTextCompare) > 0 Then
                Set FindFrameByCaptionDeep_ = c: Exit Function
            End If
        End If
        If HasControls__(c) Then
            Set r = FindFrameByCaptionDeep_(c, capLike)
            If Not r Is Nothing Then Set FindFrameByCaptionDeep_ = r: Exit Function
        End If
    Next
End Function

' 枠内でCaption一致のCheckBoxを探して取得/設定
Private Function GetCheckInFrameByCaptionLike_(owner As Object, ByVal frameCap As String, ByVal chkCap As String) As Boolean
    Dim fr As Object, c As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function
    For Each c In fr.Controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(chkCap), vbTextCompare) > 0 Then
                GetCheckInFrameByCaptionLike_ = (c.value <> 0)
                Exit Function
            End If
        End If
    Next
End Function

Private Sub SetCheckInFrameByCaptionLike_(owner As Object, ByVal frameCap As String, ByVal chkCap As String, ByVal v As Boolean)
    Dim fr As Object, c As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Sub
    For Each c In fr.Controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(chkCap), vbTextCompare) > 0 Then
                c.value = v: Exit Sub
            End If
        End If
    Next
End Sub

' 枠内でラベルCaption一致→同じ枠のTextBoxを拾う
Private Function GetTextInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String) As String
    Dim fr As Object, c As Object, inner As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function
    For Each c In fr.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                For Each inner In fr.Controls
                    If TypeName(inner) = "TextBox" Then GetTextInFrameByLabelCaption_ = CStr(inner.Text): Exit Function
                Next
            End If
        End If
    Next
End Function

Private Sub SetTextInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String, ByVal s As String)
    Dim fr As Object, c As Object, inner As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Sub
    For Each c In fr.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                For Each inner In fr.Controls
                    If TypeName(inner) = "TextBox" Then inner.Text = s: Exit Sub
                Next
            End If
        End If
    Next
End Sub


' 枠内でラベルCaption一致 → 同じ枠で一番近いComboBoxを取得/設定
Private Function GetComboInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String) As String
    Dim fr As Object, c As Object, best As Object, inner As Object
    Dim targetLbl As Object, bestDx As Single, dx As Single
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function

    ' ラベルを見つける
    For Each c In fr.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                Set targetLbl = c: Exit For
            End If
        End If
    Next
    If targetLbl Is Nothing Then Exit Function

    ' ラベルと同じ親内で、横方向に一番近いComboBoxを選ぶ
    bestDx = 1E+20
    For Each inner In fr.Controls
        If TypeName(inner) = "ComboBox" Then
            dx = Abs(CSng(inner.Top) - CSng(targetLbl.Top)) + Abs(CSng(inner.Left) - CSng(targetLbl.Left))
            If dx < bestDx Then Set best = inner: bestDx = dx
        End If
    Next
    If Not best Is Nothing Then GetComboInFrameByLabelCaption_ = CStr(best.value)
End Function

Private Sub SetComboInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String, ByVal s As String)
    Dim fr As Object, c As Object, best As Object, inner As Object
    Dim targetLbl As Object, bestDx As Single, dx As Single
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Sub

    For Each c In fr.Controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                Set targetLbl = c: Exit For
            End If
        End If
    Next
    If targetLbl Is Nothing Then Exit Sub

    bestDx = 1E+20
    For Each inner In fr.Controls
        If TypeName(inner) = "ComboBox" Then
            dx = Abs(CSng(inner.Top) - CSng(targetLbl.Top)) + Abs(CSng(inner.Left) - CSng(targetLbl.Left))
            If dx < bestDx Then Set best = inner: bestDx = dx
        End If
    Next
    If Not best Is Nothing Then best.value = s
End Sub







' --- Captionでチェックボックスを探す（部分一致OK） ---
Public Function FindCheckByCaptionLike(container As Object, ByVal needle As String) As MSForms.CheckBox
    Dim c As Object
    On Error Resume Next
    For Each c In container.Controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(needle)) > 0 Then
                Set FindCheckByCaptionLike = c
                Exit Function
            End If
        End If
        ' FrameやPageなど入れ子も辿る
        If HasControls_(c) Then
            Set FindCheckByCaptionLike = FindCheckByCaptionLike(c, needle)
            If Not FindCheckByCaptionLike Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function HasControls_(obj As Object) As Boolean
    On Error Resume Next
    HasControls_ = (obj.Controls.Count >= 0)
End Function

