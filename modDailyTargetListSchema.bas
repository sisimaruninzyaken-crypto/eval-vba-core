Attribute VB_Name = "modDailyTargetListSchema"

Option Explicit

'=== Daily Log: 対象者一覧 ListBox の共有スキーマ ===
' UI (frmEval) と保存ロジック (modEvalIOEntry) の双方から参照する。

Public Const DAILY_TARGET_COL_NAME As Long = 0
Public Const DAILY_TARGET_COL_PID As Long = 1
Public Const DAILY_TARGET_COL_EXCLUDE As Long = 2
Public Const DAILY_TARGET_COL_CATEGORY As Long = 3
Public Const DAILY_TARGET_COL_COUNT As Long = 4

Public Const DAILY_TARGET_CATEGORY_NORMAL As String = "通常"
Public Const DAILY_TARGET_CATEGORY_ADDED As String = "追加"
Public Const DAILY_TARGET_EXCLUDE_MARK As String = "除外"

Public Function IsDailyTargetExcludeMarker(ByVal marker As String) As Boolean
    IsDailyTargetExcludeMarker = (StrComp(Trim$(marker), DAILY_TARGET_EXCLUDE_MARK, vbTextCompare) = 0)
End Function

