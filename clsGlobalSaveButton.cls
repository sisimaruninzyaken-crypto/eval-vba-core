VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGlobalSaveButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public Event Clicked()

Private Sub btn_Click()
    ' 下段の保存ボタンが押されたら、自分のイベントを発火させるだけ
    RaiseEvent Clicked
End Sub

