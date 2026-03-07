Attribute VB_Name = "modExport"
Option Explicit

Public Sub ExportVbaProjectToDesktop()
    Const EXPORT_DIR As String = "C:\Users\jimu\Desktop\_vba_export"

    Dim comp As Object
    Dim outPath As String
    Dim successCount As Long

    EnsureDirectoryExists EXPORT_DIR

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' Standard Module
                outPath = EXPORT_DIR & "\" & comp.name & ".bas"
                DeleteIfExists outPath
                comp.Export outPath
                If FileExists(outPath) Then successCount = successCount + 1

            Case 2 ' Class Module
                outPath = EXPORT_DIR & "\" & comp.name & ".cls"
                DeleteIfExists outPath
                comp.Export outPath
                If FileExists(outPath) Then successCount = successCount + 1

            Case 3 ' UserForm (.frm + .frx)
                outPath = EXPORT_DIR & "\" & comp.name & ".frm"
                DeleteIfExists outPath
                DeleteIfExists EXPORT_DIR & "\" & comp.name & ".frx"
                comp.Export outPath
                If FileExists(outPath) Then successCount = successCount + 1

            Case Else
                ' Exclude Document components and all non-target types
        End Select
    Next comp

    MsgBox "エクスポート成功件数: " & successCount, vbInformation
End Sub

Private Sub EnsureDirectoryExists(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

Private Sub DeleteIfExists(ByVal filePath As String)
    If FileExists(filePath) Then
        Kill filePath
    End If
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir(filePath, vbNormal)) > 0)
End Function
