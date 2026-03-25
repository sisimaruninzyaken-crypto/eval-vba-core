Attribute VB_Name = "modPostureIO"

' ===== modPostureIO.bas・郁､・焚鬆・岼・句ｙ閠・迚茨ｼ・====
Option Explicit

Public Sub SavePostureToSheet(ws As Worksheet, ByVal r As Long, owner As Object)
    Dim caps As Variant, i As Long, cap As String, col As Long
    caps = PostureCaptions()
    For i = LBound(caps) To UBound(caps)
        cap = CStr(caps(i))
        col = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_" & cap)
        ws.Cells(r, col).value = GetCheckByCaption(owner, cap)
        Debug.Print "[SAVE][Posture]", cap, "=", ws.Cells(r, col).value
    Next

   ' 鬪ｨ逶､蛯ｾ譁懶ｼ医さ繝ｳ繝懶ｼ・竊・譫縲悟ｧｿ蜍｢隧穂ｾ｡縲榊・縺ｮ繝ｩ繝吶Ν縲碁ｪｨ逶､蛯ｾ譁懊阪↓邏舌▼縺修ombo繧呈鏡縺・
Dim cPel As Long, pel As String
cPel = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_鬪ｨ逶､蛯ｾ譁・)
pel = GetComboInFrameByLabelCaption_(owner, "蟋ｿ蜍｢隧穂ｾ｡", "鬪ｨ逶､蛯ｾ譁・)
ws.Cells(r, cPel).value = pel
Debug.Print "[SAVE][Posture] 鬪ｨ逶､蛯ｾ譁・=", pel

' 荳頑ｮｵ 蛯呵・ｼ亥ｧｿ蜍｢縺ｮ蛯呵・ｼ・竊・譫縲悟ｧｿ蜍｢隧穂ｾ｡縲榊・縺ｮ繝ｩ繝吶Ν縲悟ｙ閠・阪↓邏舌▼縺週extBox繧呈鏡縺・
Dim cNote As Long, noteVal As String
cNote = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_蛯呵・)
noteVal = GetTextInFrameByLabelCaption_(owner, "蟋ｿ蜍｢隧穂ｾ｡", "蛯呵・)
ws.Cells(r, cNote).value = noteVal
Debug.Print "[SAVE][Posture] 蛯呵・=", noteVal


   
   ' ?? 髢｢遽諡倡ｸｮ・夐ｸ驛ｨ・亥ｷｦ蜿ｳ縺ｪ縺暦ｼ・
Dim colNeck As Long
colNeck = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_鬆ｸ驛ｨ")
ws.Cells(r, colNeck).value = GetCheckInFrameByCaptionLike_(owner, "髢｢遽諡倡ｸｮ", "鬆ｸ驛ｨ")
Debug.Print "[SAVE][Posture] 諡倡ｸｮ_鬆ｸ驛ｨ =", ws.Cells(r, colNeck).value

   

    ' ?? 髢｢遽諡倡ｸｮ・・驛ｨ菴搾ｼ亥承/蟾ｦ・・??
Dim colJ As Long, joints As Variant, jr As String

joints = Array("閧ｩ髢｢遽", "閧倬未遽", "謇矩未遽", "閧｡髢｢遽", "閹晞未遽", "雜ｳ髢｢遽")

For i = LBound(joints) To UBound(joints)
    jr = CStr(joints(i))
    colJ = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_" & Replace(jr, "髢｢遽", "") & "_蜿ｳ")
    ws.Cells(r, colJ).value = GetKoushuku_OnRow_(owner, jr, "蜿ｳ")

    colJ = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_" & Replace(jr, "髢｢遽", "") & "_蟾ｦ")
    ws.Cells(r, colJ).value = GetKoushuku_OnRow_(owner, jr, "蟾ｦ")
Next


    ' 蛯呵・ｼ亥ｧｿ蜍｢繝悶Ο繝・け・・
    Dim cKNote As Long, kNote As String
cKNote = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_蛯呵・)
kNote = GetTextInFrameByLabelCaption_(owner, "髢｢遽諡倡ｸｮ", "蛯呵・)
ws.Cells(r, cKNote).value = kNote
Debug.Print "[SAVE][Posture] 諡倡ｸｮ_蛯呵・=", kNote

End Sub



Public Sub LoadPostureFromSheet(ws As Worksheet, ByVal r As Long, owner As Object)
    
    Debug.Print "[POSTURE][ENTER] r=" & r
    
    Dim caps As Variant, i As Long, cap As String, col As Long, v As Variant
    caps = PostureCaptions()
    For i = LBound(caps) To UBound(caps)
        cap = CStr(caps(i))
        col = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_" & cap)
        v = ws.Cells(r, col).value
        



        SetCheckByCaption owner, cap, CBool(v)
        
    Next

    ' 鬪ｨ逶､蛯ｾ譁懶ｼ医さ繝ｳ繝懶ｼ・
Dim cPel As Long, vPel As Variant
cPel = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_鬪ｨ逶､蛯ｾ譁・)
vPel = ws.Cells(r, cPel).value
SetComboInFrameByLabelCaption_ owner, "蟋ｿ蜍｢隧穂ｾ｡", "鬪ｨ逶､蛯ｾ譁・, CStr(vPel)

' 荳頑ｮｵ 蛯呵・ｼ亥ｧｿ蜍｢縺ｮ蛯呵・ｼ・
Dim cNote As Long, vNote As Variant
cNote = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_蛯呵・)
vNote = ws.Cells(r, cNote).value
SetTextInFrameByLabelCaption_ owner, "蟋ｿ蜍｢隧穂ｾ｡", "蛯呵・, CStr(vNote)


    
    ' ?? 髢｢遽諡倡ｸｮ・夐ｸ驛ｨ・亥ｷｦ蜿ｳ縺ｪ縺暦ｼ・
Dim colNeck As Long, vNeck As Variant
colNeck = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_鬆ｸ驛ｨ")
vNeck = ws.Cells(r, colNeck).value
SetCheckInFrameByCaptionLike_ owner, "髢｢遽諡倡ｸｮ", "鬆ｸ驛ｨ", CBool(vNeck)


    

    ' ?? 髢｢遽諡倡ｸｮ・・驛ｨ菴搾ｼ亥承/蟾ｦ・・??
Dim colJ As Long, joints As Variant, jr As String

joints = Array("閧ｩ髢｢遽", "閧倬未遽", "謇矩未遽", "閧｡髢｢遽", "閹晞未遽", "雜ｳ髢｢遽")

For i = LBound(joints) To UBound(joints)
    jr = CStr(joints(i))

    colJ = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_" & Replace(jr, "髢｢遽", "") & "_蜿ｳ")
    v = ws.Cells(r, colJ).value
    SetKoushuku_OnRow_ owner, jr, "蜿ｳ", CBool(v)

    colJ = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_" & Replace(jr, "髢｢遽", "") & "_蟾ｦ")
    v = ws.Cells(r, colJ).value
    SetKoushuku_OnRow_ owner, jr, "蟾ｦ", CBool(v)
Next i

    ' 蛯呵・ｼ亥ｧｿ蜍｢繝悶Ο繝・け・・
    Dim cKNote As Long, vKNote As Variant
cKNote = EnsureHeaderCol_Posture(ws, "蟋ｿ蜍｢_諡倡ｸｮ_蛯呵・)
vKNote = ws.Cells(r, cKNote).value
SetTextInFrameByLabelCaption_ owner, "髢｢遽諡倡ｸｮ", "蛯呵・, CStr(vKNote)


Debug.Print "[POSTURE][EXIT] r=" & r

End Sub



' ====== 縺薙％縺九ｉ繝倥Ν繝・======

' 謇ｱ縺・メ繧ｧ繝・け鬆・岼蜷搾ｼ医く繝｣繝励す繝ｧ繝ｳ・峨ｒ縺薙％縺ｫ霑ｽ蜉縺励※縺・￥

Private Function PostureCaptions() As Variant
    PostureCaptions = Array("鬆ｭ驛ｨ蜑肴婿遯∝・", "蜀・レ", "蛛ｴ蠑ｯ", "菴灘ｹｹ蝗樊雷", "蜿榊ｼｵ閹・)
End Function


' --- CheckBox・・aption荳閾ｴ・牙叙蠕・險ｭ螳・---
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

' --- Caption縺ｧ繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ繧呈爾縺呻ｼ亥・繧悟ｭ仙・蟶ｰ繝ｻ驛ｨ蛻・ｸ閾ｴOK・・--
Private Function FindCheckByCaptionLike_(container As Object, ByVal needle As String) As Object
    Dim c As Object
    On Error Resume Next
    For Each c In container.controls
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

' --- 蛯呵・ｼ壹Λ繝吶Ν縲悟ｙ閠・阪→蜷後§隕ｪ蜀・・TextBox繧呈鏡縺・---
Private Function GetTextByLabelCaption(owner As Object, ByVal labelCap As String) As String
    Dim tb As Object: Set tb = FindTextBoxNearLabel_(owner, labelCap)
    If Not tb Is Nothing Then GetTextByLabelCaption = CStr(tb.text)
End Function

Private Sub SetTextByLabelCaption(owner As Object, ByVal labelCap As String, ByVal s As String)
    Dim tb As Object: Set tb = FindTextBoxNearLabel_(owner, labelCap)
    If Not tb Is Nothing Then tb.text = s
End Sub

' 繝ｩ繝吶Ν縺ｮCaption荳閾ｴ竊貞酔縺倩ｦｪ・・rame縺ｪ縺ｩ・牙・縺ｮTextBox繧定ｿ斐☆・域怙蛻昴・1縺､・・
Private Function FindTextBoxNearLabel_(container As Object, ByVal labelCap As String) As Object
    Dim c As Object, inner As Object
    On Error Resume Next
    For Each c In container.controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                ' 蜷後§隕ｪ蜀・・TextBox繧定ｿ斐☆
                For Each inner In c.parent.controls
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

' 蟄舌ｒ謖√▽縺九・蛻､螳夲ｼ亥ｮ牙・迚茨ｼ・
Private Function HasControls__(obj As Object) As Boolean
    On Error Resume Next
    HasControls__ = (obj.controls.count >= 0)
End Function

' --- 蟋ｿ蜍｢逕ｨ繝ｭ繝ｼ繧ｫ繝ｫ: 隕句・縺怜・繧堤｢ｺ菫昴＠縺ｦ蛻礼分蜿ｷ繧定ｿ斐☆ ---
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
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
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
    For Each c In container.controls
        If TypeName(c) = "Label" And InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
            For Each inner In c.parent.controls
                If TypeName(inner) = wantType Then Set FindFirstByTypeNearLabel_ = inner: Exit Function
            Next
        End If
        If HasControls__(c) Then
            Set FindFirstByTypeNearLabel_ = FindFirstByTypeNearLabel_(c, labelCap, wantType)
            If Not FindFirstByTypeNearLabel_ Is Nothing Then Exit Function
        End If
    Next
End Function



' ==== 髢｢遽諡倡ｸｮ・郁｡後Λ繝吶Ν縺ｮ縲悟承・丞ｷｦ縲阪メ繧ｧ繝・け・峨・繝ｫ繝・====

' 繝ｩ繝吶ΝCaption・晞Κ菴榊錐・井ｾ九瑚か髢｢遽縲搾ｼ峨→蜷後§陦後・縲悟承・丞ｷｦ縲垢heckBox繧呈爾縺呻ｼ・op縺瑚ｿ代＞繧ゅ・繧呈治逕ｨ・・
Private Function FindSideCheck_OnSameRow_(owner As Object, ByVal rowLabel As String, ByVal sideCaption As String) As Object
    Dim lbl As Object, p As Object, c As Object
    Dim best As Object, bestDy As Single, dy As Single, tol As Single
    tol = 18 ' 蜷後§陦後→縺ｿ縺ｪ縺儺op霍晞屬・医ヵ繧ｩ繝ｼ繝縺ｮ繧ｹ繧ｱ繝ｼ繝ｫ縺ｫ繧医ｊ隱ｿ謨ｴ蜿ｯ・・

    Set lbl = FindLabelByCaptionDeep_(owner, rowLabel)
    If lbl Is Nothing Then Exit Function

    Set p = lbl.parent
    bestDy = 1E+20
    For Each c In p.controls
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

' 繝ｩ繝吶Ν・磯Κ菴榊錐・峨ｒ豺ｱ縺乗爾縺・
Private Function FindLabelByCaptionDeep_(container As Object, ByVal cap As String) As Object
    Dim c As Object, r As Object
    On Error Resume Next
    For Each c In container.controls
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

' 蜿ｳ蟾ｦ縺ｮ蛟､蜿門ｾ励・險ｭ螳・
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



' === Frame・域棧・峨ｒCaption縺ｧ迚ｹ螳壹＠縺ｦ荳ｭ縺ｮ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ繧呈桶縺・===

' 譫Caption縺ｫ驛ｨ蛻・ｸ閾ｴ縺吶ｋFrame繧呈ｷｱ縺乗爾縺・
Private Function FindFrameByCaptionDeep_(container As Object, ByVal capLike As String) As Object
    Dim c As Object, r As Object
    On Error Resume Next
    For Each c In container.controls
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

' 譫蜀・〒Caption荳閾ｴ縺ｮCheckBox繧呈爾縺励※蜿門ｾ・險ｭ螳・
Private Function GetCheckInFrameByCaptionLike_(owner As Object, ByVal frameCap As String, ByVal chkCap As String) As Boolean
    Dim fr As Object, c As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function
    For Each c In fr.controls
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
    For Each c In fr.controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(chkCap), vbTextCompare) > 0 Then
                c.value = v: Exit Sub
            End If
        End If
    Next
End Sub

' 譫蜀・〒繝ｩ繝吶ΝCaption荳閾ｴ竊貞酔縺俶棧縺ｮTextBox繧呈鏡縺・
Private Function GetTextInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String) As String
    Dim fr As Object, c As Object, inner As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function
    For Each c In fr.controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                For Each inner In fr.controls
                    If TypeName(inner) = "TextBox" Then GetTextInFrameByLabelCaption_ = CStr(inner.text): Exit Function
                Next
            End If
        End If
    Next
End Function

Private Sub SetTextInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String, ByVal s As String)
    Dim fr As Object, c As Object, inner As Object
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Sub
    For Each c In fr.controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                For Each inner In fr.controls
                    If TypeName(inner) = "TextBox" Then inner.text = s: Exit Sub
                Next
            End If
        End If
    Next
End Sub


' 譫蜀・〒繝ｩ繝吶ΝCaption荳閾ｴ 竊・蜷後§譫縺ｧ荳逡ｪ霑代＞ComboBox繧貞叙蠕・險ｭ螳・
Private Function GetComboInFrameByLabelCaption_(owner As Object, ByVal frameCap As String, ByVal labelCap As String) As String
    Dim fr As Object, c As Object, best As Object, inner As Object
    Dim targetLbl As Object, bestDx As Single, dx As Single
    Set fr = FindFrameByCaptionDeep_(owner, frameCap)
    If fr Is Nothing Then Exit Function

    ' 繝ｩ繝吶Ν繧定ｦ九▽縺代ｋ
    For Each c In fr.controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                Set targetLbl = c: Exit For
            End If
        End If
    Next
    If targetLbl Is Nothing Then Exit Function

    ' 繝ｩ繝吶Ν縺ｨ蜷後§隕ｪ蜀・〒縲∵ｨｪ譁ｹ蜷代↓荳逡ｪ霑代＞ComboBox繧帝∈縺ｶ
    bestDx = 1E+20
    For Each inner In fr.controls
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

    For Each c In fr.controls
        If TypeName(c) = "Label" Then
            If InStr(1, Trim$(c.caption), Trim$(labelCap), vbTextCompare) > 0 Then
                Set targetLbl = c: Exit For
            End If
        End If
    Next
    If targetLbl Is Nothing Then Exit Sub

    bestDx = 1E+20
    For Each inner In fr.controls
        If TypeName(inner) = "ComboBox" Then
            dx = Abs(CSng(inner.Top) - CSng(targetLbl.Top)) + Abs(CSng(inner.Left) - CSng(targetLbl.Left))
            If dx < bestDx Then Set best = inner: bestDx = dx
        End If
    Next
    If Not best Is Nothing Then best.value = s
End Sub







' --- Caption縺ｧ繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ繧呈爾縺呻ｼ磯Κ蛻・ｸ閾ｴOK・・---
Public Function FindCheckByCaptionLike(container As Object, ByVal needle As String) As MSForms.CheckBox
    Dim c As Object
    On Error Resume Next
    For Each c In container.controls
        If TypeName(c) = "CheckBox" Then
            If InStr(1, Trim$(c.caption), Trim$(needle)) > 0 Then
                Set FindCheckByCaptionLike = c
                Exit Function
            End If
        End If
        ' Frame繧Пage縺ｪ縺ｩ蜈･繧悟ｭ舌ｂ霎ｿ繧・
        If HasControls_(c) Then
            Set FindCheckByCaptionLike = FindCheckByCaptionLike(c, needle)
            If Not FindCheckByCaptionLike Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function HasControls_(obj As Object) As Boolean
    On Error Resume Next
    HasControls_ = (obj.controls.count >= 0)
End Function

