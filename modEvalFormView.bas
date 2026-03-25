Attribute VB_Name = "modEvalFormView"
Option Explicit
Public Const CAP_ADL As String = "日常生活動作"
Public Const CAP_BI As String = "バーサルインデックス"
Public Const CAP_IADL As String = "IADL"
Public Const CAP_KYO As String = "起居動作"

Public Const TAG_BI_PREFIX As String = "BI."
Public Const TAG_IADL_PREFIX As String = "IADL."
Public Const TAG_POSTURE_PREFIX As String = "POSTURE|"

' === 身体機能評価タブ用（追記） ===
Public Const CAP_FUNC              As String = "身体機能評価"
Public Const CAP_FUNC_ROM          As String = "ROM（主要関節）"
Public Const CAP_FUNC_MMT          As String = "筋力（MMT）"
Public Const CAP_FUNC_SENS_REF     As String = "感覚（表在・深部）／筋緊張・反射（痙縮含む）/変形・疼痛（部位／NRS）"
Public Const CAP_FUNC_NOTE         As String = "備考"

Public Const TAG_FUNC_PREFIX       As String = "PHYS"      ' 保存用カテゴリ接頭辞
Public Const HOST_BODY_NAME        As String = "hostBody"  ' frmEval内のフレーム名（既存そのまま）
Public Const MP_PHYS_NAME          As String = "mpPhys"    ' 身体機能評価用MultiPage名


Public Const CAP_FUNC_PARALYSIS As String = "麻痺"



