Attribute VB_Name = "modExport"
Option Explicit

Public Sub Z_ExpAll()
    Dim comp As Object
    Dim exportPath As String

    exportPath = ThisWorkbook.path & "\_vba_export\"
    MsgBox "START: " & exportPath

    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1: comp.Export exportPath & comp.name & ".bas"
            Case 2: comp.Export exportPath & comp.name & ".cls"
            Case 3: comp.Export exportPath & comp.name & ".frm"
        End Select
    Next comp

    MsgBox "DONE"
End Sub
