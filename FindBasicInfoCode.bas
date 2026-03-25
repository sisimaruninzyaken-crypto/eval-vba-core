Attribute VB_Name = "FindBasicInfoCode"
' プロジェクト内の全モジュールからキーワードを横断検索
' ※「ファイル > オプション > セキュリティ センター > セキュリティ センターの設定 > マクロの設定」
'    で「VBAプロジェクト オブジェクトモデルへの信頼を付与」にチェックが必要です。
Public Sub FindBasicInfoCode()
    Dim targets As Variant
    targets = Array("BasicInfo", "基本情報", "SaveBasicInfo", "LoadBasicInfo", "EnsureHeaderCol_BasicInfo", "評価日", "氏名")

    On Error GoTo TrustErr
    Dim vbProj As Object, vbComp As Object, codeMod As Object
    Set vbProj = Application.VBE.ActiveVBProject

    Dim t As Variant, i As Long, lastLine As Long, lineText As String
    Debug.Print "---- BasicInfo 検索 ----"
    For Each vbComp In vbProj.VBComponents
        Set codeMod = vbComp.CodeModule
        If Not codeMod Is Nothing Then
            lastLine = codeMod.CountOfLines
            For i = 1 To lastLine
                lineText = codeMod.lines(i, 1)
                For Each t In targets
                    If InStr(1, lineText, CStr(t), vbTextCompare) > 0 Then
                        Debug.Print vbComp.name & "  L" & i & "  : " & Trim$(lineText)
                        Exit For
                    End If
                Next t
            Next i
        End If
    Next vbComp
    Debug.Print "---- 検索完了 ----"
    Exit Sub

TrustErr:
    Debug.Print "[ERROR] プロジェクトにアクセスできません。信頼設定を有効にしてください。"
End Sub

