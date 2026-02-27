Attribute VB_Name = "modWinHost"
'=== modWinHost (32/64 両対応) ===
#If VBA7 Then
    ' --- 32/64 を分岐して Ptr 系 API をエイリアス ---
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        ' 32bit Office では *Ptr が無いので Get/SetWindowLongA にエイリアス
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If

    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#Else
    ' 古い VBA 環境（推奨外だが互換用）
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_THICKFRAME  As Long = &H40000

Private Const SWP_NOMOVE       As Long = &H2
Private Const SWP_NOSIZE       As Long = &H1
Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Public Function GetFormHwnd(ByVal caption As String) As LongPtr
    GetFormHwnd = FindWindowA("ThunderDFrame", caption)
End Function

Public Sub EnableFormSystemButtons(ByVal hwnd As LongPtr, _
                                   Optional allowMin As Boolean = True, _
                                   Optional allowMax As Boolean = True)
    Dim s As LongPtr
    s = GetWindowLongPtr(hwnd, GWL_STYLE)
    If allowMin Then s = s Or WS_MINIMIZEBOX
    If allowMax Then s = s Or WS_MAXIMIZEBOX
    s = s Or WS_THICKFRAME

    Call SetWindowLongPtr(hwnd, GWL_STYLE, s)
    DrawMenuBar hwnd
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub


