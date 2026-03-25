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
    ' 荳区ｮｵ縺ｮ菫晏ｭ倥・繧ｿ繝ｳ縺梧款縺輔ｌ縺溘ｉ縲∬・蛻・・繧､繝吶Φ繝医ｒ逋ｺ轣ｫ縺輔○繧九□縺・
    RaiseEvent Clicked
End Sub

