Attribute VB_Name = "modEvalIOEntry"


Option Explicit

Public Const EVAL_SHEET_NAME As String = "EvalData"
Private Const EVAL_INDEX_SHEET_NAME As String = "EvalIndex"
Private Const CLIENT_MASTER_SHEET_NAME As String = "ClientMaster"
Private Const EVAL_HISTORY_SHEET_PREFIX As String = "EV_"
Private Const HDR_ROWNO As String = "RowNo"
Private Const HDR_USER_ID As String = "UserID"
Private Const HDR_NAME As String = "Name"
Private Const HDR_KANA As String = "Kana"
Private Const HDR_SHEET As String = "SheetName"
Private Const HDR_FIRST_EVAL As String = "FirstEvalDate"
Private Const HDR_LATEST_EVAL As String = "LatestEvalDate"
Private Const HDR_RECORD_COUNT As String = "RecordCount"
Private Const HDR_BIRTH_DATE As String = "BirthDate"
Private Const HDR_GENDER As String = "Gender"
Private Const HDR_CARE_LEVEL As String = "CareLevel"
Private Const HDR_CREATED_DATE As String = "CreatedDate"
Public mDailyLogManual As Boolean    ' 譌･縲・・險倬鹸縺ｮ謇句虚菫晏ｭ倥ヵ繝ｩ繧ｰ



' === 陬懷勧蜈ｷ/繝ｪ繧ｹ繧ｯ 繝輔Ξ繝ｼ繝蜷搾ｼ亥崋螳夂畑・・===
Private Const FRM_AIDS As String = "Frame33"
Private Const FRM_RISK As String = "Frame34"
Private Const IO_TRACE As Boolean = False
Private Const MAIN_SAVE_MIN_FILLED_FIELDS As Long = 10
Private Const MAIN_SAVE_FEW_INPUT_MESSAGE As String = "蜈･蜉幃・岼縺悟ｰ代↑縺・憾諷九〒縺吶・ & vbCrLf & _
    "譌｢蟄倥ョ繝ｼ繧ｿ繧剃ｸ頑嶌縺阪☆繧九→蜈・↓謌ｻ縺帙↑縺・庄閭ｽ諤ｧ縺後≠繧翫∪縺吶・ & vbCrLf & _
    "譛ｬ蠖薙↓菫晏ｭ倥＠縺ｾ縺吶°・・
Private Const MAIN_SAVE_MIN_CHANGE_COUNT As Long = 3
Private Const MAIN_SAVE_FEW_CHANGE_MESSAGE As String = "螟画峩鬆・岼縺後⊇縺ｨ繧薙←縺ゅｊ縺ｾ縺帙ｓ縲・ & vbCrLf & _
    "隱､縺｣縺ｦ菫晏ｭ倥＠繧医≧縺ｨ縺励※縺・↑縺・°遒ｺ隱阪＠縺ｦ縺上□縺輔＞縲・ & vbCrLf & _
    "譛ｬ蠖薙↓菫晏ｭ倥＠縺ｾ縺吶°・・
Private Const HDR_HOMEENV_CHECKS As String = "Basic.HomeEnv.Checks"
Private Const HDR_HOMEENV_NOTE As String = "Basic.HomeEnv.Note"
Private Const HDR_RISK_CHECKS As String = "Basic.Risk.Checks"
Private Const HDR_AIDS_CHECKS As String = "Basic.Aids.Checks"
Private Const HISTORY_LOAD_DEBUG As Boolean = True


Public Sub LoadEvaluation_CurrentRow()
    MsgBox "縺薙・蜈･蜿｣縺ｯ蟒・ｭ｢縺励∪縺励◆縲りｪｭ縺ｿ霎ｼ縺ｿ縺ｯ縲悟錐蜑坂・逶ｴ霑大呵｣懊°繧蛾∈謚槭阪↓邨ｱ荳縺励※縺・∪縺吶・, vbInformation
End Sub

' 笘・％縺薙ｒ騾驕ｿ蜷阪↓縺励※蠢・★髢峨§繧・
Private Sub LoadEvaluation_fromLastRow_OBSOLETE()
End Sub




Private Sub IO_T(ParamArray a())
    If Not IO_TRACE Then Exit Sub
    Dim i As Long, s As String
    For i = LBound(a) To UBound(a): s = s & IIf(i > 0, " ", "") & CStr(a(i)): Next
    Debug.Print Format(Now, "hh:nn:ss"), s
End Sub

Private Sub IO_SafeRunSave(procName As String, ws As Worksheet, r As Long, owner As Object)
    On Error GoTo EH
    IO_T "[RUN] target", procName

    
    IO_T "[SAVE] call", procName
    Application.Run procName, ws, r, owner
    IO_T "[SAVE] ok", procName
    Exit Sub
EH:
    IO_T "[SAVE] NG", procName, "Err", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub IO_SafeRunLoad(procName As String, ws As Worksheet, r As Long, owner As Object)
    On Error GoTo EH
    IO_T "[LOAD] call", procName
    Application.Run procName, ws, r, owner
    IO_T "[LOAD] ok", procName
    Exit Sub
EH:
    IO_T "[LOAD] NG", procName, "Err", Err.Number, Err.Description
    Err.Clear
End Sub


Private Sub t(ParamArray a())
    If Not IO_TRACE Then Exit Sub
    Dim i As Long, s As String
    For i = LBound(a) To UBound(a)
        s = s & IIf(i > 0, " ", "") & CStr(a(i))
    Next
    Debug.Print Format(Now, "hh:nn:ss"), s
End Sub



' 笘・ompat・壽立蜈･蜿｣縲ょ・驛ｨ逧・↓縺ｯ SaveEvaluation_Append_From 縺ｫ蟋碑ｭｲ縺吶ｋ縲・
' 縲縺ｩ縺薙°縺ｮ繝懊ち繝ｳ繧・商縺・・繧ｯ繝ｭ縺後∪縺 SaveEvaluation_Append 繧呈欠縺励※縺・※繧ゅ・
' 縲譛邨ら噪縺ｪ菫晏ｭ倥Ν繝ｼ繝医・ SaveEvaluation_Append_From 縺ｫ荳譛ｬ蛹悶＆繧後ｋ縲・
Public Sub SaveEvaluation_Append()
    EnsureFormLoaded                ' frmEval 縺後Ο繝ｼ繝峨＆繧後※縺・↑縺代ｌ縺ｰ繝ｭ繝ｼ繝・
    SaveEvaluation_Append_From frmEval
End Sub


' 笘・OBSOLETE] 逶ｴ謗･蜻ｼ縺ｰ縺ｪ縺・りｪｭ縺ｿ霎ｼ縺ｿ縺ｯ LoadEvaluation_ByName_From 縺ｫ荳譛ｬ蛹悶・
Private Sub LoadEvaluation_LastRow_OBSOLETE(owner As Object)

    MsgBox "縺薙・蜈･蜿｣縺ｯ蟒・ｭ｢縺励∪縺励◆縲りｪｭ縺ｿ霎ｼ縺ｿ縺ｯ縲主錐蜑坂・逶ｴ霑大呵｣懊°繧蛾∈謚槭上↓邨ｱ荳縺励※縺・∪縺吶・, vbInformation
End Sub


Private Sub SaveEvaluation_CurrentRow_OBSOLETE()
    MsgBox "縺薙・蜈･蜿｣縺ｯ蟒・ｭ｢縺励∪縺励◆縲ゆｿ晏ｭ倥・縲手ｿｽ蜉菫晏ｭ假ｼ・ppend・峨上↓邨ｱ荳縺励※縺・∪縺吶・, vbInformation
End Sub
Private Sub LoadEvaluation_CurrentRow_OBSOLETE()
    ' OBSOLETE: this procedure must not be used.
    Debug.Assert False
    Exit Sub
End Sub

'======================== 螳滉ｽ難ｼ壼・驛ｨ縺ｾ縺ｨ繧√※蜻ｼ縺ｶ ========================

' ===== 縺吶∋縺ｦ菫晏ｭ・=====
Public Sub SaveAllSectionsToSheet(ws As Worksheet, r As Long, owner As Object)


   ' 菫晏ｭ倥ワ繝厄ｼ哘valData 1 陦悟・縺ｫ縺ｾ縺ｨ繧√※譖ｸ縺崎ｾｼ繧
' 菫晏ｭ倬・・繧､繝｡繝ｼ繧ｸ・・
'   1) 蝓ｺ譛ｬ諠・ｱ・・asic・・
'   2) 鮗ｻ逞ｺ / ROM / 蟋ｿ蜍｢
'   3) MMT / 諢溯ｦ・/ 繝医・繝ｳ繝ｻ蜿榊ｰ・
'   4) 逍ｼ逞幢ｼ・ain IO・・
'   5) 繝・せ繝医・隧穂ｾ｡・・0m / TUG / 謠｡蜉・/ 5蝗樒ｫ九■ / 繧ｻ繝溘ち繝ｳ繝・Β・・
'   6) 陬懷勧蜈ｷ / 繝ｪ繧ｹ繧ｯ・医メ繧ｧ繝・け鄒､・・
'   7) ADL・・O_ADL・・

   
   

    ' 蝓ｺ譛ｬ諠・ｱ・医％縺ｮ繝｢繧ｸ繝･繝ｼ繝ｫ蜀・・螳溯｣・ｼ・
    Call SaveBasicInfoToSheet_FromMe(ws, r, owner)



    ' 鮗ｻ逞ｺ / ROM・域里縺ｫOK・・
    IO_SafeRunSave "SaveParalysisToSheet", ws, r, owner
    IO_SafeRunSave "SaveROMToSheet", ws, r, owner
    IO_SafeRunSave "SavePostureToSheet", ws, r, owner
    


    ' 蠢・ｦ√↓縺ｪ縺｣縺溘ｉ鬆・ｬ｡ON
    IO_SafeRunSave "SaveMMTToSheet", ws, r, owner
    IO_SafeRunSave "SaveSensoryToSheet", ws, r, owner
     'Call Mirror_SensoryIO(ws, r)    'Legacy莠呈鋤・夂樟陦御ｻ墓ｧ倥〒縺ｯ譛ｪ菴ｿ逕ｨ縺ｮ縺溘ａ蛛懈ｭ｢
    IO_SafeRunSave "modToneReflexIO.SaveToneReflexToSheet", ws, r, owner
  

    Call SavePainToSheet(ws, r, owner)
     Call Save_TestEvalToSheet(ws, r, owner)
     Call Save_WalkIndepToSheet(ws, r, owner)  '笘・ｭｩ陦瑚・遶句ｺｦ IO_WalkIndep 菫晏ｭ・
     Call Save_WalkAbnToSheet(ws, r, owner)    '笘・焚蟶ｸ豁ｩ陦・IO_WalkAbn 菫晏ｭ・
     Call Save_WalkRLAToSheet(ws, r, owner)    '笘・LA IO_WalkRLA 菫晏ｭ・



Call Save_ADL_AtRow(ws, r)




End Sub

' ===== 縺吶∋縺ｦ隱ｭ霎ｼ =====
'====================================================================
' [HUB] 隧穂ｾ｡隱ｭ縺ｿ霎ｼ縺ｿ繝上ヶ
'  - 蜻ｼ縺ｳ蜃ｺ縺怜・・哭oadEvaluation_ByName_From・域ｭ｣隕丞・蜿｣・峨↑縺ｩ
'  - 蠖ｹ蜑ｲ・・
'       1) 蜷榊燕縺九ｉ縲梧怙譁ｰ陦後阪↓ r 繧貞ｷｮ縺玲崛縺医ｋ・・indLatestRowByName・・
'       2) BasicInfo / ROM / 蟋ｿ蜍｢ / MMT / 諢溯ｦ壹・繝医・繝ｳ / 逍ｼ逞・/
'          繝・せ繝郁ｩ穂ｾ｡ / 豁ｩ陦・/ 隱咲衍繝ｻ邊ｾ逾・縺ｪ縺ｩ蜷・そ繧ｯ繧ｷ繝ｧ繝ｳ縺ｮ
'          Load*FromSheet 繧偵∪縺ｨ繧√※蜻ｼ縺ｳ蜃ｺ縺・
'  - 豕ｨ諢擾ｼ・
'       * 莉悶Δ繧ｸ繝･繝ｼ繝ｫ縺九ｉ縺薙％繧堤峩謗･蜻ｼ縺ｶ縺ｮ縺ｯ讌ｵ蜉幃∩縺代ｋ
'         ・郁ｪｭ縺ｿ霎ｼ縺ｿ莉墓ｧ倥・荳蜈・ｮ｡逅・・縺溘ａ・・
'       * 蜷・そ繧ｯ繧ｷ繝ｧ繝ｳ縺ｮ UI 繝ｬ繧､繧｢繧ｦ繝郁ｪｿ謨ｴ縺ｯ縺薙％縺ｧ縺ｯ陦後ｏ縺ｪ縺・
'====================================================================
Public Sub LoadAllSectionsFromSheet(ws As Worksheet, r As Long, owner As Object)

    Dim nm As String
    Dim rLatest As Long

    ' 笘・酔縺伜錐蜑阪↑繧峨√◎縺ｮ莠ｺ縺ｮ縲梧怙譁ｰ陦後阪↓隱ｭ縺ｿ霎ｼ縺ｿ陦後ｒ蟾ｮ縺玲崛縺医ｋ
         nm = Trim$(owner.txtName.text)

    ' 笘・ヵ繧ｩ繝ｼ繝蛛ｴ縺檎ｩｺ縺ｪ繧峨√す繝ｼ繝医・豌丞錐繧ｻ繝ｫ縺九ｉ諡ｾ縺・
    If Len(nm) = 0 Then
        Dim cName As Long
        cName = FindHeaderCol(ws, "Basic.Name")
        If cName = 0 Then cName = FindHeaderCol(ws, "豌丞錐")
        If cName = 0 Then cName = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
        If cName = 0 Then cName = FindHeaderCol(ws, "蜷榊燕")


        If cName > 0 Then
            nm = Trim$(CStr(ws.Cells(r, cName).value))
        End If
    End If
    
    

    ' 笘・・蜿｣縺ｧ r 縺梧欠螳壹＆繧後※縺・ｋ蝣ｴ蜷医・蟆企㍾縺吶ｋ・医％縺薙〒荳頑嶌縺阪＠縺ｪ縺・ｼ・
If r < 2 And Len(nm) > 0 Then
    rLatest = FindLatestRowByName(ws, nm)
    If rLatest > 0 Then r = rLatest
End If




   ' 鮗ｻ逞ｺ / ROM / 蟋ｿ蜍｢縺ｮ隱ｭ霎ｼ縺ｯ LoadBasicInfoFromSheet_FromMe 蜀・〒
    ' chkLoadParalysis / chkLoadROM / chkLoadPosture 縺ｫ蠢懊§縺ｦ螳滓命
    
    Call LoadBasicInfoFromSheet_FromMe(ws, r, owner)
    IO_SafeRunLoad "Load_ADL_FromRow", ws, r, owner
   


    
    'Call LoadParalysisFromSheet(ws, r, owner)
    'Call LoadROMFromSheet(ws, r, owner)
    Call LoadSensoryFromSheet(ws, r, owner)
    'Call LoadPostureFromSheet(ws, r, owner)
    
   
    Call Load_TestEvalFromSheet(ws, r, owner)
    Call Load_WalkIndepFromSheet(ws, r, owner)
    Call Load_WalkAbnFromSheet(ws, r, owner)
    Call Load_WalkRLAFromSheet(ws, r, owner)   '笘・LA隱ｭ縺ｿ霎ｼ縺ｿ

    'Call MMT.LoadMMTFromSheet(ws, r, owner)
    Call modToneReflexIO.LoadToneReflexFromSheet(ws, r, owner)


   

    IO_SafeRunLoad "LoadPainFromSheet", ws, r, owner
    
    ' 陬懷勧蜈ｷ
Dim cA As Long
cA = FindHeaderCol(ws, "陬懷勧蜈ｷ")
If cA > 0 Then
    DeserializeChecks owner, "Frame33", CStr(ws.Cells(r, cA).value), True   ' 陬懷勧蜈ｷ
End If

' 繝ｪ繧ｹ繧ｯ
Dim cR As Long
cR = FindHeaderCol(ws, "繝ｪ繧ｹ繧ｯ")
If cR > 0 Then
    DeserializeChecks owner, "Frame34", CStr(ws.Cells(r, cR).value), False  ' 繝ｪ繧ｹ繧ｯ
End If
    
        Call Load_CognitionMental_FromRow(ws, r, owner)
        'Load_DailyLog_Latest_FromForm owner
        
End Sub


'====================================================================
' [ENTRY] 隧穂ｾ｡隱ｭ縺ｿ霎ｼ縺ｿ縺ｮ豁｣隕丞・蜿｣
'  - UI 蛛ｴ・・rmEval 繧・ｻ悶ヵ繧ｩ繝ｼ繝・峨・蜴溷援縺薙％縺縺代ｒ蜻ｼ縺ｳ蜃ｺ縺・
'  - 蜷榊燕・・xtName・峨°繧・EvalData 荳翫・譛譁ｰ陦後ｒ迚ｹ螳壹＠縲・
'    LoadAllSectionsFromSheet 縺ｫ蟋碑ｭｲ縺吶ｋ
'  - LoadAllSectionsFromSheet / 蜷・そ繧ｯ繧ｷ繝ｧ繝ｳ縺ｮ Load*FromSheet 縺ｯ
'    莉悶Δ繧ｸ繝･繝ｼ繝ｫ縺九ｉ逶ｴ謗･蜻ｼ縺ｰ縺ｪ縺・％縺ｨ・郁ｪｭ縺ｿ霎ｼ縺ｿ莉墓ｧ倥・蛻・｣る亟豁｢・・
'====================================================================
Public Sub LoadEvaluation_ByName_From(owner As Object)



    EnsureFormLoaded
    Dim wsTarget As Worksheet
    Dim resolveMessage As String
    If ResolveUserHistorySheet(owner, False, wsTarget, resolveMessage) Then
        Dim validRow As Long
        Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
        Dim nameVal As String: nameVal = Trim$(owner.txtName.text)
        Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
        HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                               "resolvedSheet=" & HistoryLoadDebug_SheetName(wsTarget), _
                               "nameVal=" & HistoryLoadDebug_Quote(nameVal), _
                               "idVal=" & HistoryLoadDebug_Quote(idVal), _
                               "kanaVal=" & HistoryLoadDebug_Quote(kanaVal)

        If Len(idVal) > 0 Then
            validRow = FindLatestValidEvalRowByIdentity(wsTarget, nameVal, idVal, kanaVal)
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "identityLookupCalled=True", _
                                   "identityRow=" & CStr(validRow)
        End If
        If validRow = 0 Then
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "fallbackFindLatestRowByName=True"
            validRow = FindLatestRowByName(wsTarget, nameVal)
        Else
            HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                                   "fallbackFindLatestRowByName=False"
        End If
        HistoryLoadDebug_Print "[LoadEvaluation_ByName_From]", _
                               "finalValidRow=" & CStr(validRow)
        
        If validRow = 0 Then
             HistoryLoadDebug_ScanWorkbookForName nameVal, wsTarget
            MsgBox "蟇ｾ雎｡縺ｮ隧穂ｾ｡螻･豁ｴ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbInformation
            Exit Sub
        End If
        LoadAllSectionsFromSheet wsTarget, validRow, owner
        Exit Sub

    End If

    If Len(resolveMessage) > 0 Then
        MsgBox resolveMessage, vbExclamation
    End If
    ' 笘・％縺薙∪縺ｧ

End Sub


' 荳九°繧蛾■縺｣縺ｦ豌丞錐荳閾ｴ縺ｮ譛譁ｰ陦後ｒ霑斐☆・郁ｦ句・縺励・縲梧ｰ丞錐縲阪悟茜逕ｨ閠・錐縲阪悟錐蜑阪阪ｒ鬆・↓謗｢縺呻ｼ・
Public Function FindLatestRowByName(ws As Worksheet, nameText As String) As Long

    Dim c As Long
    c = FindHeaderCol(ws, "Basic.Name")
    If c = 0 Then c = FindHeaderCol(ws, "豌丞錐")
    If c = 0 Then c = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
    If c = 0 Then c = FindHeaderCol(ws, "蜷榊燕")
    If c = 0 Then c = FindHeaderCol(ws, "Name")
    If c = 0 Then
        HistoryLoadDebug_Print "[FindLatestRowByName]", _
                               "sheet=" & HistoryLoadDebug_SheetName(ws), _
                               "targetName=" & HistoryLoadDebug_Quote(nameText), _
                               "nameHeaderMissing=True"
        Exit Function
    End If

    HistoryLoadDebug_Print "[FindLatestRowByName]", _
                           "sheet=" & HistoryLoadDebug_SheetName(ws), _
                           "targetName=" & HistoryLoadDebug_Quote(nameText), _
                           "nameHeader=" & HistoryLoadDebug_Quote(CStr(ws.Cells(1, c).value)), _
                           "nameCol=" & CStr(c)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, c).End(xlUp).row
    Dim r As Long
    Dim rowName As String
    Dim normalizedTarget As String
    Dim normalizedRow As String

    normalizedTarget = NormalizeName(nameText)
    HistoryLoadDebug_Print "[FindLatestRowByName]", "lastRow=" & CStr(lastRow)
    For r = lastRow To 2 Step -1      ' 1陦檎岼縺ｯ隕句・縺玲Φ螳・
        rowName = CStr(ws.Cells(r, c).value)
        normalizedRow = NormalizeName(rowName)
        HistoryLoadDebug_Print "[FindLatestRowByName][SCAN]", _
                               "row=" & CStr(r), _
                               "raw=" & HistoryLoadDebug_Quote(rowName), _
                               "normalized=" & HistoryLoadDebug_Quote(normalizedRow), _
                               "matched=" & CStr(normalizedRow = normalizedTarget)
        If normalizedRow = normalizedTarget Then
            HistoryLoadDebug_Print "[FindLatestRowByName]", _
                                   "matchedRow=" & CStr(r)
            FindLatestRowByName = r
            Exit Function
        End If
    Next r

    HistoryLoadDebug_Print "[FindLatestRowByName]", "matchedRow=0"
End Function



Public Function CountRowsByName(ws As Worksheet, nameText As String) As Long
    Dim c As Long
    c = FindHeaderCol(ws, "豌丞錐")
    If c = 0 Then c = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
    If c = 0 Then c = FindHeaderCol(ws, "蜷榊燕")
    If c = 0 Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, c).End(xlUp).row

    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, c).value), nameText, vbTextCompare) = 0 Then
            CountRowsByName = CountRowsByName + 1
        End If
    Next r
End Function



Public Function FindLatestRowByNameAndID( _
        ws As Worksheet, _
        nameText As String, _
        idVal As String) As Long

    Dim cName As Long, cID As Long
    cName = FindHeaderCol(ws, "豌丞錐")
    If cName = 0 Then cName = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
    If cName = 0 Then cName = FindHeaderCol(ws, "蜷榊燕")
    If cName = 0 Then Exit Function

    cID = FindColByHeaderExact(ws, "Basic.ID")
    If cID = 0 Then cID = FindColByHeaderExact(ws, "ID")
    If cID = 0 Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, cName).End(xlUp).row

    ' 荳九°繧画爾縺呻ｼ晄怙譁ｰ蜆ｪ蜈・
    For r = lastRow To 2 Step -1
        If StrComp(CStr(ws.Cells(r, cName).value), nameText, vbTextCompare) = 0 Then
            If StrComp(CStr(ws.Cells(r, cID).value), idVal, vbTextCompare) = 0 Then
                FindLatestRowByNameAndID = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function FindEvalIndexRowsByUserID(ByVal indexWs As Worksheet, ByVal userID As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    Dim r As Long

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(indexWs.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then c.Add r
    Next r

    Set FindEvalIndexRowsByUserID = c
End Function

Private Function BuildDuplicateUserIDMessage(ByVal indexWs As Worksheet, ByVal userID As String, ByVal rowsByID As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long

    For i = 1 To rowsByID.count
        rowNo = CLng(rowsByID(i))
        lines = lines & _
            "- Name: " & Trim$(CStr(indexWs.Cells(rowNo, 2).value)) & _
            " / Kana: " & Trim$(CStr(indexWs.Cells(rowNo, 3).value)) & _
            " / Sheet: " & Trim$(CStr(indexWs.Cells(rowNo, 4).value))
        If i < rowsByID.count Then lines = lines & vbCrLf
    Next i

    BuildDuplicateUserIDMessage = _
       "EvalIndex蜀・〒蜷御ｸID縺瑚､・焚蟄伜惠縺励※縺・∪縺吶・ & vbCrLf & _
       "ID: " & userID & vbCrLf & vbCrLf & lines
End Function

Private Function BuildUserIdentityMismatchMessage(ByVal userID As String, _
                                                  ByVal inputName As String, _
                                                  ByVal indexName As String, _
                                                  ByVal inputKana As String, _
                                                  ByVal indexKana As String) As String
    Dim lines As String

    lines = lines & "ID荳堺ｸ閾ｴ繧ｨ繝ｩ繝ｼ" & vbCrLf
    lines = lines & "ID: " & userID & vbCrLf
    lines = lines & "蜈･蜉帶ｰ丞錐: " & inputName & vbCrLf
    lines = lines & "逋ｻ骭ｲ豌丞錐: " & indexName

    If Len(Trim$(inputKana)) > 0 Or Len(Trim$(indexKana)) > 0 Then
        lines = lines & vbCrLf & "蜈･蜉帙き繝・ " & inputKana & vbCrLf & "逋ｻ骭ｲ繧ｫ繝・ " & indexKana
    End If

    BuildUserIdentityMismatchMessage = lines
End Function

Private Function IsSameKanaIfAvailable(ByVal leftKana As String, ByVal rightKana As String) As Boolean
    leftKana = Trim$(leftKana)
    rightKana = Trim$(rightKana)

    If Len(leftKana) = 0 Or Len(rightKana) = 0 Then
        IsSameKanaIfAvailable = True
    Else
        IsSameKanaIfAvailable = (StrComp(leftKana, rightKana, vbTextCompare) = 0)
    End If
End Function

Private Function FindLatestValidEvalRowByIdentity(ByVal ws As Worksheet, _
                                                  ByVal nameText As String, _
                                                  ByVal idVal As String, _
                                                  Optional ByVal kanaText As String = "") As Long
    Dim cEval As Long: cEval = FindColByHeaderExact(ws, "Basic.EvalDate")
    Dim cID As Long: cID = FindColByHeaderExact(ws, "Basic.ID")
    Dim cName As Long
    Dim cKana As Long: cKana = FindColByHeaderExact(ws, "Basic.NameKana")
    Dim lastRow As Long
    Dim r As Long
    Dim d As Date
    Dim bestDate As Date
    Dim bestRow As Long
    Dim rowName As String
    Dim rowKana As String

    If cEval = 0 Then Exit Function

    If cID = 0 Then cID = FindColByHeaderExact(ws, "ID")
    If cID = 0 Then Exit Function

    cName = FindColByHeaderExact(ws, "Basic.Name")
      If cName = 0 Then cName = FindHeaderCol(ws, "豌丞錐")
      If cName = 0 Then cName = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
      If cName = 0 Then cName = FindHeaderCol(ws, "Name")
    If cName = 0 Then Exit Function

    lastRow = LastDataRow(ws)

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, cID).value)), Trim$(idVal), vbTextCompare) <> 0 Then GoTo NextRow

        rowName = Trim$(CStr(ws.Cells(r, cName).value))
        If NormalizeName(rowName) <> NormalizeName(nameText) Then GoTo NextRow

        If Len(Trim$(kanaText)) > 0 And cKana > 0 Then
            rowKana = Trim$(CStr(ws.Cells(r, cKana).value))
            If Len(rowKana) > 0 Then
                If StrComp(rowKana, Trim$(kanaText), vbTextCompare) <> 0 Then GoTo NextRow
            End If
        End If

        If Not TryParseEvalDate(ws.Cells(r, cEval).value, d) Then GoTo NextRow

        If bestRow = 0 Then
            bestRow = r
            bestDate = d
        ElseIf d > bestDate Then
            bestRow = r
            bestDate = d
        ElseIf d = bestDate Then
            If r > bestRow Then bestRow = r
        End If
NextRow:
    Next r

    FindLatestValidEvalRowByIdentity = bestRow
End Function


'======================== 陬懷勧・壹ヵ繧ｩ繝ｼ繝・上す繝ｼ繝茨ｼ剰｡・========================

Private Sub EnsureFormLoaded()
    On Error Resume Next
    Dim t$: t = frmEval.caption            ' 蜿ら・縺ｧ縺阪ｌ縺ｰ繝ｭ繝ｼ繝画ｸ医∩
    If Err.Number <> 0 Then Load frmEval
    On Error GoTo 0
    If frmEval.Visible = False Then frmEval.Show vbModeless   ' 繝｢繝・Ν繝ｬ繧ｹ縺ｧ謫堺ｽ懷庄
End Sub

Private Function EnsureEvalSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureEvalSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureEvalSheet Is Nothing Then
        Set EnsureEvalSheet = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        On Error Resume Next
        EnsureEvalSheet.name = sheetName   ' 譌｢蟄伜錐縺ｪ繧右xcel縺瑚・蜍輔Μ繝阪・繝
        On Error GoTo 0
    End If
End Function



Private Function LastDataRow(ws As Worksheet) As Long
    On Error Resume Next
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    LastDataRow = IIf(f Is Nothing, 1, f.row)
End Function

Private Function NextAppendRow(ws As Worksheet) As Long
    Dim lr As Long: lr = LastDataRow(ws)
    NextAppendRow = IIf(lr < 2, 2, lr + 1)
End Function

'====================================================================
' [ENTRY] 隧穂ｾ｡菫晏ｭ倥・豁｣隕丞・蜿｣
'  - UI 蛛ｴ・・rmEval 繧・ｻ悶ヵ繧ｩ繝ｼ繝・峨・蜴溷援縺薙％縺縺代ｒ蜻ｼ縺ｳ蜃ｺ縺・
'  - 陦後・豎ｺ螳夲ｼ・ppend 陦鯉ｼ峨・縺薙・荳ｭ縺ｧ NextAppendRow 縺ｫ繧医ｊ荳蜈・ｮ｡逅・
'  - SaveAllSectionsToSheet / SaveBasicInfoToSheet_FromMe 遲峨・荳倶ｽ埼未謨ｰ繧・
'    逶ｴ謗･莉悶Δ繧ｸ繝･繝ｼ繝ｫ縺九ｉ蜻ｼ縺ｰ縺ｪ縺・％縺ｨ・医せ繧ｭ繝ｼ繝槫､画峩譎ゅ・貍上ｌ髦ｲ豁｢・・
'====================================================================



Public Sub SaveEvaluation_Append_From(owner As Object)
    Dim wsUser As Worksheet
    Dim resolveMessage As String


    If ResolveUserHistorySheet(owner, True, wsUser, resolveMessage) Then
        EnsureHistorySheetInitialized wsUser
        EnsureClientMasterEntry owner
        
        Dim patientName As String
        patientName = Trim$(GetCtlTextGeneric(owner, "txtName"))
        If Len(patientName) = 0 Then
              MsgBox "謔｣閠・錐繧貞・蜉帙＠縺ｦ縺九ｉ菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲・, vbExclamation
              Exit Sub
        End If
        
        Dim warnMessage As String
        warnMessage = GetSparseMainSaveWarningMessage(wsUser, patientName, owner)
        If Len(warnMessage) > 0 Then
            If MsgBox(warnMessage, vbExclamation + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        Dim appendRow As Long
        appendRow = NextAppendRow(wsUser)
        
        wsUser.Cells(appendRow, EnsureHeader(wsUser, HDR_ROWNO)).value = appendRow - 1
        SaveAllSectionsToSheet wsUser, appendRow, owner
        Save_CognitionMental_AtRow wsUser, appendRow, owner
        MirrorBasicRow wsUser, appendRow
        
        Dim idxRow As Long
        idxRow = FindEvalIndexRowBySheetName(EnsureEvalIndexSheet(), wsUser.name)
        If idxRow > 0 Then
            UpdateEvalIndexMetadata owner, idxRow, wsUser.name
            UpdateEvalIndexStats idxRow, wsUser
        End If
        Exit Sub
    End If



    If Len(resolveMessage) > 0 Then
        MsgBox resolveMessage, vbExclamation
    Else
        MsgBox "菫晏ｭ伜・繧ｷ繝ｼ繝医′隕九▽縺九ｉ縺ｪ縺・◆繧√∽ｿ晏ｭ倥ｒ荳ｭ譁ｭ縺励∪縺吶・, vbExclamation
    End If
    
End Sub

Private Function ClientMasterHeaders() As Variant
    ClientMasterHeaders = Array(HDR_USER_ID, HDR_NAME, HDR_KANA, HDR_BIRTH_DATE, HDR_GENDER, HDR_CARE_LEVEL, HDR_CREATED_DATE)
End Function

Private Function EnsureClientMasterSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CLIENT_MASTER_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        ws.name = CLIENT_MASTER_SHEET_NAME
    End If

    Dim headers As Variant: headers = ClientMasterHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = CStr(headers(i))
    Next i

    Set EnsureClientMasterSheet = ws
End Function

Private Function FindClientMasterRowsByName(ByVal ws As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(NormalizeName(CStr(ws.Cells(r, 2).value)), NormalizeName(nameText), vbTextCompare) = 0 Then c.Add r
    Next r
    Set FindClientMasterRowsByName = c
End Function

Private Function FindClientMasterRowByUserID(ByVal ws As Worksheet, ByVal userID As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then
            FindClientMasterRowByUserID = r
            Exit Function
        End If
    Next r
End Function

Private Function FindClientMasterRow(ByVal ws As Worksheet, ByVal userID As String, ByVal nameText As String, ByRef shouldSkip As Boolean) As Long
    Dim rowsByName As Collection

    If Len(Trim$(userID)) > 0 Then
        FindClientMasterRow = FindClientMasterRowByUserID(ws, userID)
        Exit Function
    End If

    If Len(Trim$(nameText)) = 0 Then Exit Function

    Set rowsByName = FindClientMasterRowsByName(ws, nameText)
    If rowsByName.count = 1 Then
        FindClientMasterRow = CLng(rowsByName(1))
    ElseIf rowsByName.count > 1 Then
        shouldSkip = True
    End If
End Function

Private Function TryGetBirthDateForClientMaster(ByVal owner As Object, ByRef outDateText As String) As Boolean
    On Error GoTo EH

    Dim rawBirth As String
    rawBirth = Trim$(GetCtlTextGeneric(owner, "txtBirth"))
    If Len(rawBirth) = 0 Then Exit Function

    Dim dtBirth As Date
    If CallByName(owner, "TryGetBirthDateForStorage", VbMethod, rawBirth, dtBirth) Then
        outDateText = Format$(dtBirth, "yyyy/mm/dd")
        TryGetBirthDateForClientMaster = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Private Sub EnsureClientMasterEntry(ByVal owner As Object)
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = EnsureClientMasterSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim nameVal As String: nameVal = Trim$(GetCtlTextGeneric(owner, "txtName"))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    Dim genderVal As String: genderVal = Trim$(GetCtlTextGeneric(owner, "cboSex"))
    Dim careVal As String: careVal = Trim$(GetCtlTextGeneric(owner, "cboCare"))

    Dim skipRegistration As Boolean
    Dim hitRow As Long
    hitRow = FindClientMasterRow(ws, idVal, nameVal, skipRegistration)
    If hitRow > 0 Then Exit Sub
    If skipRegistration Then Exit Sub
    If Len(nameVal) = 0 Then Exit Sub

    Dim birthText As String
    Call TryGetBirthDateForClientMaster(owner, birthText)

    Dim newRow As Long
    newRow = NextAppendRow(ws)

    ws.Cells(newRow, 1).value = idVal
    ws.Cells(newRow, 2).value = nameVal
    ws.Cells(newRow, 3).value = kanaVal
    ws.Cells(newRow, 4).value = birthText
    ws.Cells(newRow, 5).value = genderVal
    ws.Cells(newRow, 6).value = careVal
    ws.Cells(newRow, 7).value = Format$(Date, "yyyy/mm/dd")
    Exit Sub
EH:
    Err.Clear
End Sub


Private Function GetSparseMainSaveWarningMessage(ws As Worksheet, ByVal patientName As String, owner As Object) As String
    Dim existingRow As Long
    existingRow = ResolveExistingEvalRow(ws, patientName, owner)
    If existingRow <= 0 Then Exit Function

    Dim totalCount As Long, blankCount As Long
    CountMainFormTextInputs owner, totalCount, blankCount
    

    Dim filledCount As Long
    filledCount = CountMainFormFilledFields(owner)

    Dim changeCount As Long
    changeCount = CountMainFormTextboxChanges(ws, existingRow, owner)
    changeCount = changeCount + CountMainFormMajorBlockChanges(ws, existingRow, owner)
    
    If filledCount < MAIN_SAVE_MIN_FILLED_FIELDS Then
        GetSparseMainSaveWarningMessage = MAIN_SAVE_FEW_INPUT_MESSAGE
        Exit Function
    End If

    If changeCount < MAIN_SAVE_MIN_CHANGE_COUNT Then
        GetSparseMainSaveWarningMessage = MAIN_SAVE_FEW_CHANGE_MESSAGE
    End If
End Function

Private Function ResolveExistingEvalRow(ws As Worksheet, ByVal patientName As String, owner As Object) As Long
    ResolveExistingEvalRow = FindLatestRowByName(ws, patientName)

    Dim idVal As String
    idVal = Trim$(GetID_FromBasicInfo(owner))
    If Len(idVal) = 0 Then Exit Function

    Dim rowByID As Long
    rowByID = FindLatestRowByNameAndID(ws, patientName, idVal)
    If rowByID > 0 Then ResolveExistingEvalRow = rowByID
End Function

Private Function CountMainFormTextboxChanges(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Long
    Dim map As Variant
    map = MainSaveTextboxHeaderMap()

    Dim i As Long
    For i = LBound(map) To UBound(map)
        Dim headerName As String
        Dim ctlName As String
        Dim c As Long
        Dim curVal As String
        Dim oldVal As String

        headerName = CStr(map(i)(0))
        ctlName = CStr(map(i)(1))

        c = FindColByHeaderExact(ws, headerName)
        If c = 0 Then GoTo NextItem

        curVal = NormalizeCompareValue(GetCtlTextGeneric(owner, ctlName))
        oldVal = NormalizeCompareValue(CStr(ws.Cells(existingRow, c).value))

        If StrComp(curVal, oldVal, vbBinaryCompare) <> 0 Then
            CountMainFormTextboxChanges = CountMainFormTextboxChanges + 1
        End If
NextItem:
    Next i
End Function

Private Function CountMainFormFilledFields(owner As Object) As Long
    Dim map As Variant
    map = MainSaveTextboxHeaderMap()

    Dim i As Long
    For i = LBound(map) To UBound(map)
        If Len(NormalizeCompareValue(GetCtlTextGeneric(owner, CStr(map(i)(1))))) > 0 Then
            CountMainFormFilledFields = CountMainFormFilledFields + 1
        End If
    Next i
End Function

Private Function CountMainFormMajorBlockChanges(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Long
    If HasMainFormTestEvalChange(ws, existingRow, owner) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If

    If HasMainFormBIChange(ws, existingRow) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If

    If HasMainFormIADLChange(ws, existingRow) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If

    If HasMainFormKyoChange(ws, existingRow) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If

    If HasMainFormROMChange(ws, existingRow, owner) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If

    If HasMainFormMMTChange(ws, existingRow, owner) Then
        CountMainFormMajorBlockChanges = CountMainFormMajorBlockChanges + 1
    End If
End Function

Private Function HasMainFormTestEvalChange(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Boolean
    HasMainFormTestEvalChange = HasSerializedBlockChange( _
        Build_TestEval_IO(owner), _
        ReadStr_Compat("IO_TestEval", existingRow, ws), _
        TestEvalCompareKeys() _
    )
End Function

Private Function HasMainFormBIChange(ws As Worksheet, ByVal existingRow As Long) As Boolean
    HasMainFormBIChange = HasSerializedBlockChange( _
        Build_ADL_IO(), _
        ReadStr_Compat("IO_ADL", existingRow, ws), _
        BICompareKeys() _
    )
End Function

Private Function HasMainFormIADLChange(ws As Worksheet, ByVal existingRow As Long) As Boolean
    HasMainFormIADLChange = HasSerializedBlockChange( _
        Build_ADL_IO(), _
        ReadStr_Compat("IO_ADL", existingRow, ws), _
        IADLCompareKeys() _
    )
End Function

Private Function HasMainFormKyoChange(ws As Worksheet, ByVal existingRow As Long) As Boolean
    HasMainFormKyoChange = HasSerializedBlockChange( _
        Build_ADL_IO(), _
        ReadStr_Compat("IO_ADL", existingRow, ws), _
        KyoCompareKeys() _
    )
End Function

Private Function HasMainFormROMChange(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Boolean
    If HasMainFormROMJointChange(ws, existingRow, owner, "Upper", "Shoulder", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Upper", "Elbow", Array("Flex", "Ext")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Upper", "Forearm", Array("Sup", "Pro")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Upper", "Wrist", Array("Dorsi", "Palmar", "Radial", "Ulnar")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Lower", "Hip", Array("Flex", "Ext", "Abd", "Add", "ER", "IR")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Lower", "Knee", Array("Flex", "Ext")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMJointChange(ws, existingRow, owner, "Lower", "Ankle", Array("Dorsi", "Plantar", "Inv", "Ev")) Then
        HasMainFormROMChange = True
        Exit Function
    End If

    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_Flex", "txtROM_Trunk_Neck_Flex") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_Ext", "txtROM_Trunk_Neck_Ext") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_Rot_R", "txtROM_Trunk_Neck_Rot_R") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_Rot_L", "txtROM_Trunk_Neck_Rot_L") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_LatFlex_R", "txtROM_Trunk_Neck_LatFlex_R") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Neck_LatFlex_L", "txtROM_Trunk_Neck_LatFlex_L") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_Flex", "txtROM_Trunk_Trunk_Flex") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_Ext", "txtROM_Trunk_Trunk_Ext") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_Rot_R", "txtROM_Trunk_Trunk_Rot_R") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_Rot_L", "txtROM_Trunk_Trunk_Rot_L") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_LatFlex_R", "txtROM_Trunk_Trunk_LatFlex_R") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "ROM_Trunk_LatFlex_L", "txtROM_Trunk_Trunk_LatFlex_L") Then HasMainFormROMChange = True: Exit Function
    If HasMainFormROMFieldChange(ws, existingRow, owner, "Thorax_Expansion", "txtROM_Trunk_Thorax_ChestDiff") Then HasMainFormROMChange = True
End Function

Private Function HasMainFormROMJointChange(ws As Worksheet, ByVal existingRow As Long, owner As Object, _
                                           ByVal layer As String, ByVal joint As String, motions As Variant) As Boolean
    Dim motion As Variant
    Dim side As Variant

    For Each motion In motions
        For Each side In Array("R", "L")
            If HasMainFormROMFieldChange(ws, existingRow, owner, _
                                         "ROM_" & layer & "_" & joint & "_" & CStr(motion) & "_" & CStr(side), _
                                         "txtROM_" & layer & "_" & joint & "_" & CStr(motion) & "_" & CStr(side)) Then
                HasMainFormROMJointChange = True
                Exit Function
            End If
        Next side
    Next motion
End Function

Private Function HasMainFormROMFieldChange(ws As Worksheet, ByVal existingRow As Long, owner As Object, _
                                           ByVal headerName As String, ByVal ctlName As String) As Boolean
    Dim curVal As String
    Dim oldVal As String

    curVal = NormalizeCompareValue(GetCtlTextGeneric(owner, ctlName))
    oldVal = NormalizeCompareValue(ReadStr_Compat(headerName, existingRow, ws))

    HasMainFormROMFieldChange = (StrComp(curVal, oldVal, vbBinaryCompare) <> 0)
End Function

Private Function HasMainFormMMTChange(ws As Worksheet, ByVal existingRow As Long, owner As Object) As Boolean
    HasMainFormMMTChange = (StrComp( _
        NormalizeCompareValue(BuildCurrentMMTCompareValue(owner)), _
        NormalizeCompareValue(ReadStr_Compat("IO_MMT", existingRow, ws)), _
        vbBinaryCompare) <> 0)
End Function

Private Function BuildCurrentMMTCompareValue(owner As Object) As String
    Dim pg As Object
    Dim mp As Object
    Dim p As Long
    Dim c As Object
    Dim parts() As String
    Dim n As Long

    Set pg = MMT.GetMMTPage(owner)
    If pg Is Nothing Then Exit Function

    On Error Resume Next
    Set mp = MMT.GetMMTChildTabs(pg)
    On Error GoTo 0
    If mp Is Nothing Then Exit Function

    ReDim parts(0 To 0)
    n = -1

    For p = 0 To mp.Pages.count - 1
        For Each c In mp.Pages(p).controls
            If TypeName(c) = "ComboBox" Then
                Dim nm As String
                Dim side As String

                If Left$(c.name, 5) = "cboR_" Then
                    nm = Mid$(c.name, 6)
                    side = "R"
                ElseIf Left$(c.name, 5) = "cboL_" Then
                    nm = Mid$(c.name, 6)
                    side = "L"
                Else
                    nm = vbNullString
                    side = vbNullString
                End If

                If Len(nm) > 0 And side = "R" Then
                    Dim rVal As String
                    Dim lVal As String

                    rVal = NormalizeCompareValue(CStr(c.value))
                    On Error Resume Next
                    lVal = NormalizeCompareValue(CStr(mp.Pages(p).controls("cboL_" & nm).value))
                    On Error GoTo 0

                    n = n + 1
                    ReDim Preserve parts(0 To n)
                    parts(n) = CStr(p) & "|" & nm & "|" & rVal & "|" & lVal
                End If
            End If
        Next c
    Next p

    If n >= 0 Then BuildCurrentMMTCompareValue = Join(parts, ";")
End Function

Private Function HasSerializedBlockChange(ByVal currentValue As String, ByVal oldValue As String, keys As Variant) As Boolean
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If StrComp( _
            NormalizeCompareValue(IO_GetVal(currentValue, CStr(keys(i)))), _
            NormalizeCompareValue(IO_GetVal(oldValue, CStr(keys(i)))), _
            vbBinaryCompare) <> 0 Then
            HasSerializedBlockChange = True
            Exit Function
        End If
    Next i
End Function

Private Function TestEvalCompareKeys() As Variant
    TestEvalCompareKeys = Array( _
        "Test_10MWalk_sec", _
        "Test_TUG_sec", _
        "Test_Grip_R_kg", _
        "Test_Grip_L_kg", _
        "Test_5xSitStand_sec", _
        "Test_SemiTandem_sec" _
    )
End Function

Private Function BICompareKeys() As Variant
    BICompareKeys = Array( _
        "BITotal", _
        "BI_0", "BI_1", "BI_2", "BI_3", "BI_4", "BI_5", "BI_6", "BI_7", "BI_8", "BI_9", _
        "BI_HomeEnv_0", "BI_HomeEnv_1", "BI_HomeEnv_2", "BI_HomeEnv_3", "BI_HomeEnv_4", "BI_HomeEnv_5", "BI_HomeEnv_6" _
    )
End Function

Private Function IADLCompareKeys() As Variant
    IADLCompareKeys = Array( _
        "IADL_0", "IADL_1", "IADL_2", "IADL_3", "IADL_4", "IADL_5", "IADL_6", "IADL_7", "IADL_8" _
    )
End Function

Private Function KyoCompareKeys() As Variant
    KyoCompareKeys = Array( _
        "Kyo_Roll", "Kyo_SitUp", "Kyo_SitHold", "Kyo_StandUp", "Kyo_StandHold" _
    )
End Function


Private Function MainSaveTextboxHeaderMap() As Variant
    MainSaveTextboxHeaderMap = Array( _
        Array("隧穂ｾ｡譌･", "txtEDate"), _
        Array("蟷ｴ鮨｢", "txtAge"), _
        Array("逕溷ｹｴ譛域律", "txtBirth"), _
        Array("Basic.Name", "txtName"), _
        Array("隧穂ｾ｡閠・, "txtEvaluator"), _
        Array("隧穂ｾ｡閠・・遞ｮ", "txtEvaluatorJob"), _
        Array("逋ｺ逞・律", "txtOnset"), _
        Array("謔｣閠・eeds", "txtNeedsPt"), _
        Array("螳ｶ譌蒐eeds", "txtNeedsFam"), _
        Array("BI.SocialParticipation", "txtLiving"), _
        Array("菴丞ｮ・ｙ閠・, "txtBIHomeEnvNote"), _
        Array("荳ｻ險ｺ譁ｭ", "txtDx"), _
        Array("逶ｴ霑大・髯｢譌･", "txtAdmDate"), _
        Array("逶ｴ霑鷹髯｢譌･", "txtDisDate"), _
        Array("豐ｻ逋らｵ碁℃", "txtTxCourse"), _
        Array("蜷井ｽｵ逍ｾ謔｣", "txtComplications"), _
        Array("IO_Cog_DementiaNote", "txtDementiaNote"), _
        Array("IO_Mental_Note", "txtMentalNote") _
    )
End Function

Private Function NormalizeCompareValue(ByVal v As String) As String
    NormalizeCompareValue = Trim$(Replace(CStr(v), vbCrLf, vbLf))
End Function

Private Function ResolveDailyLogRoot(ByVal owner As Object) As Object
    If owner Is Nothing Then Exit Function
    Set ResolveDailyLogRoot = SafeGetControl(owner, "fraDailyLog")
End Function

Private Function ResolveDailyLogControl(ByVal owner As Object, ByVal controlName As String) As Object
    Dim root As Object
    If owner Is Nothing Then Exit Function
    If LenB(Trim$(controlName)) = 0 Then Exit Function

    Set root = ResolveDailyLogRoot(owner)
    If root Is Nothing Then Exit Function

    Set ResolveDailyLogControl = SafeGetControl(root, controlName)
End Function

Private Sub CountMainFormTextInputs(owner As Object, ByRef totalCount As Long, ByRef blankCount As Long)
    Dim dailyRoot As Object
    On Error Resume Next
    Set dailyRoot = ResolveDailyLogRoot(owner)
    On Error GoTo 0

    CountTextInputsRecursive owner, dailyRoot, totalCount, blankCount
End Sub

Private Sub CountTextInputsRecursive(ByVal container As Object, ByVal excludedRoot As Object, ByRef totalCount As Long, ByRef blankCount As Long)
    Dim ctrl As Object
    For Each ctrl In container.controls
        If Not excludedRoot Is Nothing Then
            If IsDescendantControl(ctrl, excludedRoot) Then GoTo NextControl
        End If

        Select Case TypeName(ctrl)
            Case "TextBox", "ComboBox"
                totalCount = totalCount + 1
                On Error Resume Next
                If Len(Trim$(CStr(ctrl.value))) = 0 Then blankCount = blankCount + 1
                On Error GoTo 0
        End Select

        On Error Resume Next
        Dim childCount As Long
        childCount = ctrl.controls.count
        If Err.Number = 0 And childCount > 0 Then
            On Error GoTo 0
            CountTextInputsRecursive ctrl, excludedRoot, totalCount, blankCount
        Else
            Err.Clear
            On Error GoTo 0
        End If
NextControl:
    Next ctrl
End Sub

Private Function IsDescendantControl(ByVal ctrl As Object, ByVal root As Object) As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = ctrl
    Do While Not p Is Nothing
        If p Is root Then
            IsDescendantControl = True
            Exit Function
        End If
        Set p = p.parent
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop
    On Error GoTo 0
End Function




Private Sub LoadEvaluation_LastRow_From_OBSOLETE(owner As Object)
    MsgBox "縺薙・蜈･蜿｣縺ｯ蟒・ｭ｢縺励∪縺励◆縲りｪｭ縺ｿ霎ｼ縺ｿ縺ｯ縲主錐蜑坂・逶ｴ霑大呵｣懊°繧蛾∈謚槭上↓邨ｱ荳縺励※縺・∪縺吶・, vbInformation
End Sub




' ====== 蝓ｺ譛ｬ諠・ｱ縺ｮ菫晏ｭ・隱ｭ霎ｼ・医％縺ｮ繝｢繧ｸ繝･繝ｼ繝ｫ蜀・ｼ・======

' 隕句・縺励・蛻励ｒ蜿門ｾ暦ｼ育┌縺代ｌ縺ｰ譁ｰ隕丈ｽ懈・・・
Private Function EnsureHeaderCol(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Then
        EnsureHeaderCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + IIf(ws.Cells(1, 1).value <> "", 1, 0)
        If EnsureHeaderCol = 0 Then EnsureHeaderCol = 1
        ws.Cells(1, EnsureHeaderCol).value = header
    Else
        EnsureHeaderCol = f.Column
    End If
End Function

' 隕句・縺励・蛻励ｒ謗｢縺呻ｼ育┌縺代ｌ縺ｰ 0・・
Private Function FindHeaderCol(ws As Worksheet, header As String) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Then FindHeaderCol = 0 Else FindHeaderCol = f.Column
End Function

Private Function FindHeaderColAny(ws As Worksheet, headers As Variant) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        FindHeaderColAny = FindHeaderCol(ws, CStr(headers(i)))
        If FindHeaderColAny > 0 Then Exit Function
    Next i
End Function


Private Sub SetCtlValueSafe(owner As Object, ctlName As String, ByVal v As Variant)
    Dim o As Object
    Set o = FindCtlDeep(owner, ctlName)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.value = v
    On Error GoTo 0
End Sub

Private Function HomeEnvControlNames() As Variant
    HomeEnvControlNames = Array( _
        "chkBIHomeEnv_Entrance", _
        "chkBIHomeEnv_Genkan", _
        "chkBIHomeEnv_IndoorStep", _
        "chkBIHomeEnv_Stairs", _
        "chkBIHomeEnv_Handrail", _
        "chkBIHomeEnv_Slope", _
        "chkBIHomeEnv_NarrowPath" _
    )
End Function

Private Function SerializeNamedChecks(owner As Object, checkNames As Variant) As String
    Dim i As Long, o As Object, s As String
    For i = LBound(checkNames) To UBound(checkNames)
        Set o = FindCtlDeep(owner, CStr(checkNames(i)))
        If Not o Is Nothing Then
            On Error Resume Next
            If CBool(o.value) Then
                If Len(s) > 0 Then s = s & ","
                s = s & CStr(checkNames(i))
            End If
            On Error GoTo 0
        End If
    Next i
    SerializeNamedChecks = s
End Function

Private Sub DeserializeNamedChecks(owner As Object, checkNames As Variant, ByVal csv As String)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim p As Variant, i As Long, o As Object
    If Len(Trim$(csv)) > 0 Then
        For Each p In Split(csv, ",")
            dict(Trim$(CStr(p))) = True
        Next
    End If

    For i = LBound(checkNames) To UBound(checkNames)
        Set o = FindCtlDeep(owner, CStr(checkNames(i)))
        If Not o Is Nothing Then
            On Error Resume Next
            o.value = dict.exists(CStr(checkNames(i)))
            On Error GoTo 0
        End If
    Next i
End Sub


' 豎守畑・壹ユ繧ｭ繧ｹ繝亥､繧貞叙蠕暦ｼ・extBox/ComboBox/Label縺ｪ縺ｩ縺ｫ蟇ｾ蠢懶ｼ・
Private Function GetCtlTextGeneric(owner As Object, ctlName As String) As String
    Dim c As Object
    Set c = FindCtlDeep(owner, ctlName)
    If c Is Nothing Then Exit Function
    
    On Error Resume Next
    GetCtlTextGeneric = CStr(c.value)
End Function

Private Function GetHdrKanaText(owner As Object) As String
    Dim c As Object

    On Error Resume Next
    Set c = owner.controls("frHeader").controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.controls("txtHdrKana")
        On Error GoTo 0
    End If

    If c Is Nothing Then
        GetHdrKanaText = ""
    Else
        On Error Resume Next
        GetHdrKanaText = Trim$(CStr(c.value))
        On Error GoTo 0
    End If
End Function

Private Sub SetHdrKanaText(owner As Object, ByVal v As Variant)
    Dim c As Object

    On Error Resume Next
    Set c = owner.controls("frHeader").controls("txtHdrKana")
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set c = owner.controls("txtHdrKana")
        On Error GoTo 0
    End If

    If c Is Nothing Then Exit Sub

    On Error Resume Next
    c.value = CStr(v)
    On Error GoTo 0
End Sub

' 豎守畑・壹さ繝ｳ繝懊ｒ螳牙・縺ｫ繧ｻ繝・ヨ・医Μ繧ｹ繝医↓縺ゅｋ譎ゅ□縺鷹∈謚橸ｼ・
Private Sub SetComboSafe_Basic(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cb As MSForms.ComboBox
    Dim s As String, i As Long, hit As Long
    s = CStr(v)
    Set cb = FindCtlDeep(owner, ctlName)
    If cb Is Nothing Then Exit Sub
    hit = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i)) = s Then hit = i: Exit For
    Next
    If hit >= 0 Then cb.ListIndex = hit Else cb.ListIndex = -1
End Sub

Private Sub SyncAgeBeforeBasicSave(ByVal owner As Object)
    On Error GoTo EH

    ' frmEval ?J??N贒ｯ\bh??pAXR[vOQ?
    CallByName owner, "SyncAgeFromBirth", VbMethod
    Exit Sub
EH:
    Debug.Print "[Basic] SyncAgeBeforeBasicSave skipped:", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub WriteBirthTextCell(ByVal target As Range, ByVal birthText As String)
    On Error Resume Next
    target.NumberFormat = "@"
    On Error GoTo 0
    target.Value2 = CStr(birthText)
End Sub

Private Function ReadBirthTextCell(ByVal target As Range) As String
    Dim s As String
    s = CStr(target.text)
    If Len(s) = 0 Then s = CStr(target.Value2)
    If Len(s) > 0 And Left$(s, 1) = "'" Then s = Mid$(s, 2)
    ReadBirthTextCell = s
End Function

'====================================================================
' BasicInfo IO 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ・郁ｩ穂ｾ｡譌･繝ｻ豌丞錐繝ｻ蟷ｴ鮨｢繝ｻNeeds 遲会ｼ・
'  - EvalData 荳翫・ Basic.* 邉ｻ繝倥ャ繝縺ｨ縺ｮ蟇ｾ蠢懊ｒ荳蜈・ｮ｡逅・☆繧狗ｪ灘哨
'  - 譁ｰ縺励＞ Basic 鬆・岼繧定ｿｽ蜉縺吶ｋ蝣ｴ蜷医・縲∝次蜑・％縺薙↓繝槭ャ繝斐Φ繧ｰ繧定ｶｳ縺・
'  - 蛻励・蛻･蜷咲ｵｱ蜷医ｄ繧ｹ繧ｭ繝ｼ繝樒ｵｱ荳縺ｯ EnsureHeaderCol_BasicInfo 蛛ｴ縺ｧ陦後≧
'  - 莉悶・繝｢繧ｸ繝･繝ｼ繝ｫ縺九ｉ縺ｯ縲。asic.* 縺ｮ迚ｩ逅・・繧堤峩謗･隗ｦ繧峨★縲・
'    蠢・ｦ√↑繧・GetID_FromBasicInfo / GetBasicInfoFrame 縺ｪ縺ｩ縺ｮ繝倥Ν繝代ｒ邨檎罰縺吶ｋ
'====================================================================




' --- 菫晏ｭ・---
Public Sub SaveBasicInfoToSheet_FromMe(ws As Worksheet, r As Long, owner As Object)
    
    Debug.Print "[Basic] Enter_SaveBasicInfo | ws=" & ws.name & " | r=" & r

    SyncAgeBeforeBasicSave owner
    
    
    '--- 蜊倅ｸ蛟､縺ｮ繝槭ャ繝斐Φ繧ｰ・域怙蠕後・隕∫ｴ縺ｫ _ 繧剃ｻ倥￠縺ｪ縺・ｼ・---
    Dim map As Variant
map = Array( _
    Array("隧穂ｾ｡譌･", "txtEDate"), _
    Array("蟷ｴ鮨｢", "txtAge"), _
    Array("逕溷ｹｴ譛域律", "txtBirth"), _
    Array("諤ｧ蛻･", "cboSex"), _
    Array("Basic.Name", "txtName"), _
    Array("隧穂ｾ｡閠・, "txtEvaluator"), _
    Array("隧穂ｾ｡閠・・遞ｮ", "txtEvaluatorJob"), _
    Array("逋ｺ逞・律", "txtOnset"), _
    Array("謔｣閠・eeds", "txtNeedsPt"), _
    Array("螳ｶ譌蒐eeds", "txtNeedsFam"), _
    Array("逕滓ｴｻ迥ｶ豕・, "txtLiving"), _
    Array("菴丞ｮ・ｙ閠・, "txtBIHomeEnvNote"), _
    Array("荳ｻ險ｺ譁ｭ", "txtDx"), _
    Array("隕∽ｻ玖ｭｷ蠎ｦ", "cboCare"), _
    Array("髫懷ｮｳ鬮倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", "cboElder"), _
    Array("隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", "cboDementia"), _
    Array("逶ｴ霑大・髯｢譌･", "txtAdmDate"), _
    Array("逶ｴ霑鷹髯｢譌･", "txtDisDate"), _
    Array("豐ｻ逋らｵ碁℃", "txtTxCourse"), _
    Array("蜷井ｽｵ逍ｾ謔｣", "txtComplications") _
)

    Call EnsureHeaderCol(ws, "N")

    '--- 譌｢蟄倥・繝ｫ繝ｼ繝暦ｼ壼腰荳蛟､繧呈嶌縺崎ｾｼ縺ｿ ---
    Dim i As Long, head As String, ctl As String, c As Long, v As String
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0)):  ctl = CStr(map(i)(1))
        v = GetCtlTextGeneric(owner, ctl)
        If Len(v) > 0 Then
            c = FindColByHeaderExact(ws, head): If c = 0 Then c = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1: ws.Cells(1, c).value = head
            If ctl = "txtBirth" Then
                WriteBirthTextCell ws.Cells(r, c), v
            Else
                ws.Cells(r, c).value = v
            End If
            Debug.Print "[BASIC][SAVE]", head, "->", v
        End If
    Next i
    
    c = EnsureHeader(ws, "菴丞ｮ・憾豕・)
    ws.Cells(r, c).value = SerializeNamedChecks(owner, HomeEnvControlNames())


    c = EnsureHeader(ws, "Basic.NameKana")
    ws.Cells(r, c).value = GetHdrKanaText(owner)
    Debug.Print "[BASIC][SAVE] Basic.NameKana ->", CStr(ws.Cells(r, c).value)
    
    Dim idVal As String: idVal = GetID_FromBasicInfo(owner)
    If Len(idVal) > 0 Then ws.Cells(r, EnsureHeader(ws, "Basic.ID")).value = idVal
    ws.Cells(r, EnsureHeader(ws, "Basic.EvalDate")).value = GetCtlTextGeneric(owner, "txtEDate")
    

    '--- 縺薙％縺九ｉ霑ｽ險假ｼ壹メ繧ｧ繝・け鄒､縺ｮCSV菫晏ｭ假ｼ郁｣懷勧蜈ｷ・上Μ繧ｹ繧ｯ・俄ｻ繝ｫ繝ｼ繝励・窶懷ｾ後ｍ窶・---
    Dim s As String
    c = EnsureHeader(ws, "陬懷勧蜈ｷ")
s = SerializeChecks(owner, "Frame33", True)
Debug.Print "[BASIC][SAVE] 陬懷勧蜈ｷ ->", s, " @col=", c
ws.Cells(r, c).value = s
c = EnsureHeader(ws, HDR_AIDS_CHECKS)
ws.Cells(r, c).value = s

   c = EnsureHeader(ws, "繝ｪ繧ｹ繧ｯ")
s = SerializeChecks(owner, "Frame34", False)
Debug.Print "[BASIC][SAVE] 繝ｪ繧ｹ繧ｯ ->", s, " @col=", c
ws.Cells(r, c).value = s

c = EnsureHeader(ws, HDR_RISK_CHECKS)
ws.Cells(r, c).value = s

c = EnsureHeader(ws, HDR_HOMEENV_CHECKS)
ws.Cells(r, c).value = SerializeNamedChecks(owner, HomeEnvControlNames())

c = EnsureHeader(ws, HDR_HOMEENV_NOTE)
ws.Cells(r, c).value = GetCtlTextGeneric(owner, "txtBIHomeEnvNote")

    
    
    
    
End Sub




' --- 隱ｭ霎ｼ ---
Public Sub LoadBasicInfoFromSheet_FromMe(ws As Worksheet, ByVal r As Long, owner As Object)

    On Error GoTo EH
    Debug.Print "[TRACE] Enter LoadBasicInfoFromSheet_FromMe r=" & r

    '--- 蜊倅ｸ蛟､縺ｮ繝槭ャ繝斐Φ繧ｰ ---
    Dim map As Variant
map = Array( _
    Array("隧穂ｾ｡譌･", "txtEDate"), _
    Array("蟷ｴ鮨｢", "txtAge"), _
    Array("逕溷ｹｴ譛域律", "txtBirth"), _
    Array("諤ｧ蛻･", "cboSex"), _
    Array("Basic.Name", "txtName"), _
    Array("隧穂ｾ｡閠・, "txtEvaluator"), _
    Array("隧穂ｾ｡閠・・遞ｮ", "txtEvaluatorJob"), _
    Array("逋ｺ逞・律", "txtOnset"), _
    Array("謔｣閠・eeds", "txtNeedsPt"), _
    Array("螳ｶ譌蒐eeds", "txtNeedsFam"), _
    Array("菴丞ｮ・ｙ閠・, "txtBIHomeEnvNote"), _
    Array("逕滓ｴｻ迥ｶ豕・, "txtLiving"), _
    Array("荳ｻ險ｺ譁ｭ", "txtDx"), _
    Array("隕∽ｻ玖ｭｷ蠎ｦ", "cboCare"), _
    Array("髫懷ｮｳ鬮倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", "cboElder"), _
    Array("隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", "cboDementia"), _
    Array("逶ｴ霑大・髯｢譌･", "txtAdmDate"), _
    Array("逶ｴ霑鷹髯｢譌･", "txtDisDate"), _
    Array("豐ｻ逋らｵ碁℃", "txtTxCourse"), _
    Array("蜷井ｽｵ逍ｾ謔｣", "txtComplications") _
)



    '--- 蜊倅ｸ蛟､繧偵ヵ繧ｩ繝ｼ繝縺ｸ隱ｭ霎ｼ ---
    Dim i As Long, head As String, ctl As String, c As Long, v As Variant
    For i = LBound(map) To UBound(map)
        head = CStr(map(i)(0))
        ctl = CStr(map(i)(1))

        c = FindHeaderCol(ws, head)
        If c > 0 Then
            If ctl = "txtBirth" Then
                v = ReadBirthTextCell(ws.Cells(r, c))
            Else
                v = ws.Cells(r, c).value
            End If
            If Left$(ctl, 3) = "cbo" Then
                SetComboSafely owner, ctl, CStr(v)
            Else
                Dim o As Object: Set o = FindCtlDeep(owner, ctl)
                If Not o Is Nothing Then o.value = v
            End If
            
        End If
        Next i

    c = FindHeaderCol(ws, "菴丞ｮ・憾豕・)
    If c > 0 Then DeserializeNamedChecks owner, HomeEnvControlNames(), CStr(ws.Cells(r, c).value)

    c = FindHeaderCol(ws, "Basic.NameKana")
    If c > 0 Then SetHdrKanaText owner, ws.Cells(r, c).value

    c = FindHeaderCol(ws, "Basic.NameKana")
    If c > 0 Then SetHdrKanaText owner, ws.Cells(r, c).value

    '--- 繝√ぉ繝・け鄒､縺ｮ蠕ｩ蜈・ｼ郁｣懷勧蜈ｷ・上Μ繧ｹ繧ｯ・・---
    Dim csv As String

    ' 陬懷勧蜈ｷ
c = FindHeaderCol(ws, "陬懷勧蜈ｷ")
If c > 0 Then
    csv = CStr(ws.Cells(r, c).value)
    DeserializeChecks owner, "Frame33", csv, True
End If

' 繝ｪ繧ｹ繧ｯ
c = FindHeaderCol(ws, "繝ｪ繧ｹ繧ｯ")
If c > 0 Then
    csv = CStr(ws.Cells(r, c).value)
    DeserializeChecks owner, "Frame34", csv, False
End If

c = FindHeaderColAny(ws, Array(HDR_AIDS_CHECKS, "?"))
If c > 0 Then DeserializeChecks owner, "Frame33", CStr(ws.Cells(r, c).value), True

c = FindHeaderColAny(ws, Array(HDR_RISK_CHECKS, "XN"))
If c > 0 Then DeserializeChecks owner, "Frame34", CStr(ws.Cells(r, c).value), False

c = FindHeaderCol(ws, HDR_HOMEENV_CHECKS)
If c > 0 Then DeserializeNamedChecks owner, HomeEnvControlNames(), CStr(ws.Cells(r, c).value)

c = FindHeaderCol(ws, HDR_HOMEENV_NOTE)
If c > 0 Then SetCtlValueSafe owner, "txtBIHomeEnvNote", ws.Cells(r, c).value




If GetBool(owner, "chkLoadParalysis", True) Then Call IO_SafeRunLoad("LoadParalysisFromSheet", ws, r, owner)
If GetBool(owner, "chkLoadROM", True) Then Call IO_SafeRunLoad("LoadROMFromSheet", ws, r, owner)
Debug.Print "[TRACE] About to run POSTURE"
If GetBool(owner, "chkLoadPosture", True) Then Call IO_SafeRunLoad("LoadPostureFromSheet", ws, r, owner)
Debug.Print "[TRACE] Done POSTURE"

Debug.Print "[TRACE] About to run MMT"
If TypeName(owner) = "frmEval" Then
    owner.QueueMMTLoadAfterUI ws, r
Else
    Call MMT.LoadMMTFromSheet(ws, r, owner)
End If
Debug.Print "[TRACE] Done MMT"
ExitHere:
    Exit Sub

EH:
    Debug.Print "[ERR][LoadBasicInfo] Err=" & Err.Number & " Desc=" & Err.Description
    Resume ExitHere


End Sub


' EvalData繧ｷ繝ｼ繝亥叙蠕・
Public Function GetEvalDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise 5, , "EvalData 繧ｷ繝ｼ繝医′縺ゅｊ縺ｾ縺帙ｓ縲・
    Set GetEvalDataSheet = ws
End Function

' 隕句・縺励°繧牙・逡ｪ蜿ｷ・亥ｮ悟・荳閾ｴ・・
Public Function FindColByHeaderExact(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindColByHeaderExact = c
            Exit Function
        End If
    Next c
End Function

' ID陦後ｒ讀懃ｴ｢・育┌縺代ｌ縺ｰ譛ｫ蟆ｾ縺ｫ菴懈・縺励※ID繧貞・繧後ｋ・・
Public Function GetOrCreateRowByID(ByVal ws As Worksheet, ByVal idVal As String) As Long
    Dim idCol As Long: idCol = FindColByHeaderExact(ws, "Basic.ID")
    If idCol = 0 Then
        ' 譌ｧ譚･縺ｮ蜻ｽ蜷阪↑繧峨％縺薙〒菴懊ｋ
        idCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If
    If Len(idVal) = 0 Then Err.Raise 5, , "ID縺檎ｩｺ縺ｧ縺吶・

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, idCol).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, idCol).value) = idVal Then
            GetOrCreateRowByID = r
            Exit Function
        End If
    Next r
    ' 辟｡縺代ｌ縺ｰ譁ｰ隕剰｡・
    r = lastRow + 1
    ws.Cells(r, idCol).value = idVal
    GetOrCreateRowByID = r
End Function





' 繝ｩ繝吶Ν縲栗D縲阪・蜿ｳ縺ｫ縺ゅｋ TextBox 縺九ｉ蛟､繧貞叙蠕暦ｼ医さ繝ｳ繝医Ο繝ｼ繝ｫ蜷阪↓萓晏ｭ倥＠縺ｪ縺・ｼ・
Public Function GetID_FromBasicInfo(ByVal owner As Object) As String
    On Error Resume Next
    GetID_FromBasicInfo = Trim$(CStr(owner.controls("frHeader").controls("txtHdrPID").value))
    On Error GoTo 0
End Function


'================ Basic諠・ｱ縺ｮ蜈ｱ騾壹・繝ｫ繝・==================

Public Function GetBasicInfoFrame(ByVal owner As Object) As Object
    Dim f As MSForms.Frame
    Set f = FindFrameByCaptionDeep_(owner, "蝓ｺ譛ｬ諠・ｱ")
    If Not f Is Nothing Then
        Set GetBasicInfoFrame = f
    Else
        Set GetBasicInfoFrame = owner   ' 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・夂峩謗･繧ｪ繝ｼ繝翫・繧呈ｸ｡縺帙ｋ繧医≧縺ｫ
    End If
End Function

Public Function GetTextByLabelInFrame(ByVal frm As Object, ByVal labelCaption As String) As String
    ' null / 髱曦rame 縺ｧ繧ょｮ牙・縺ｫ謚懊￠繧・
    If frm Is Nothing Then Exit Function
    On Error Resume Next
    Dim HasControls As Boolean
    HasControls = Not (frm.controls Is Nothing)
    On Error GoTo 0
    If Not HasControls Then Exit Function

    ' --- 莉･荳九・莉翫・繝ｭ繧ｸ繝・け縺昴・縺ｾ縺ｾ ---
    Dim lb As Object, ctl As Object
    For Each ctl In frm.controls
        If TypeName(ctl) = "Label" Then
            If InStr(1, CStr(ctl.caption), labelCaption, vbTextCompare) > 0 Then
                Set lb = ctl: Exit For
            End If
        End If
    Next ctl
    If lb Is Nothing Then Exit Function

    Dim best As Object, bestScore As Double
    bestScore = 1E+20
    For Each ctl In frm.controls
        If TypeName(ctl) = "TextBox" Then
            Dim dy As Double: dy = Abs((ctl.Top + ctl.Height / 2) - (lb.Top + lb.Height / 2))
            If dy <= lb.Height Then
                Dim dx As Double: dx = ctl.Left - lb.Left
                If dx > -5 Then
                    Dim sc As Double: sc = dy * 10 + Abs(dx)
                    If sc < bestScore Then Set best = ctl: bestScore = sc
                End If
            End If
        End If
    Next ctl
    If Not best Is Nothing Then GetTextByLabelInFrame = CStr(best.value)
End Function


' Frame 繧・Caption 驛ｨ蛻・ｸ閾ｴ縺ｧ豺ｱ縺募━蜈域爾邏｢・・serForm / Frame / MultiPage 蟇ｾ蠢懶ｼ・
Public Function FindFrameByCaptionDeep_(ByVal owner As Object, ByVal captionLike As String) As MSForms.Frame
    Set FindFrameByCaptionDeep_ = FindFrameByCaptionDeep_Walk(owner, captionLike)
End Function

Private Function FindFrameByCaptionDeep_Walk(ByVal container As Object, ByVal captionLike As String) As MSForms.Frame
    On Error Resume Next

    If TypeName(container) = "MultiPage" Then
        Dim pg As Object
        For Each pg In container.Pages
            Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(pg, captionLike)
            If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function
        Next pg
    End If

    Dim tmp As Object: Set tmp = container.controls
    If Err.Number <> 0 Then Err.Clear: Exit Function

    Dim ctl As Object
    For Each ctl In container.controls
        Select Case TypeName(ctl)
            Case "Frame"
                If InStr(1, CStr(ctl.caption), captionLike, vbTextCompare) > 0 Then
                    Set FindFrameByCaptionDeep_Walk = ctl: Exit Function
                End If
                Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function

            Case "MultiPage"
                Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function

            Case Else
                Err.Clear
                Set tmp = ctl.controls
                If Err.Number = 0 Then
                    Set FindFrameByCaptionDeep_Walk = FindFrameByCaptionDeep_Walk(ctl, captionLike)
                    If Not FindFrameByCaptionDeep_Walk Is Nothing Then Exit Function
                Else
                    Err.Clear
                End If
        End Select
    Next ctl
End Function
'================ 縺薙％縺ｾ縺ｧ雋ｼ繧・==================









' ==== BasicInfo 縺ｮ蛻怜錐繧・Basic.* 縺ｫ邨ｱ荳縺励∽ｸ崎ｶｳ縺ｯ菴懊ｋ・亥ｮ牙・繝槭・繧ｸ莉倥″・・====
Public Sub EnsureHeaderCol_BasicInfo(ByVal ws As Worksheet)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    ' --- 蜊倬・岼・井ｸｻ縺ｫ繝・く繧ｹ繝・繧ｳ繝ｳ繝懶ｼ・---
    d("BasicInfo_ID") = "Basic.ID":                  d("ID") = "Basic.ID": d("Pid") = "Basic.ID"
    d("BasicInfo_豌丞錐") = "Basic.Name":              d("豌丞錐") = "Basic.Name": d("Name") = "Basic.Name"
    d("BasicInfo_隧穂ｾ｡譌･") = "Basic.EvalDate":        d("隧穂ｾ｡譌･") = "Basic.EvalDate": d("EvalDate") = "Basic.EvalDate"
    d("BasicInfo_隧穂ｾ｡閠・) = "Basic.Evaluator":       d("隧穂ｾ｡閠・) = "Basic.Evaluator"
    d("BasicInfo_蟷ｴ鮨｢") = "Basic.Age":               d("蟷ｴ鮨｢") = "Basic.Age": d("Age") = "Basic.Age"
    d("BasicInfo_隧穂ｾ｡閠・・遞ｮ") = "Basic.EvaluatorJob": d("隧穂ｾ｡閠・・遞ｮ") = "Basic.EvaluatorJob": d("EvaluatorJob") = "Basic.EvaluatorJob"
    d("BasicInfo_諤ｧ蛻･") = "Basic.Sex":               d("諤ｧ蛻･") = "Basic.Sex": d("Sex") = "Basic.Sex"
    d("BasicInfo_荳ｻ險ｺ譁ｭ") = "Basic.PrimaryDx":       d("荳ｻ險ｺ譁ｭ") = "Basic.PrimaryDx": d("荳ｻ逞・錐") = "Basic.PrimaryDx"
    d("BasicInfo_逋ｺ逞・律") = "Basic.OnsetDate":       d("逋ｺ逞・律") = "Basic.OnsetDate"
    d("BasicInfo_隕∽ｻ玖ｭｷ蠎ｦ") = "Basic.CareLevel":     d("隕∽ｻ玖ｭｷ蠎ｦ") = "Basic.CareLevel"
    d("BasicInfo_隱咲衍逞・・遶句ｺｦ") = "Basic.DementiaADL"
    d("BasicInfo_隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ") = "Basic.DementiaADL"
    d("隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ") = "Basic.DementiaADL"
    d("BasicInfo_BI.SocialParticipation") = "BI.SocialParticipation":    d("BasicInfo_逕滓ｴｻ迥ｶ豕・) = "BI.SocialParticipation":    d("逕滓ｴｻ迥ｶ豕・) = "BI.SocialParticipation"
    AddAlias d, "Basic.LifeStatus", "BI.SocialParticipation"
    d("BasicInfo_謔｣閠・eeds") = "Basic.Needs.Patient": d("謔｣閠・eeds") = "Basic.Needs.Patient"
    d("BasicInfo_螳ｶ譌蒐eeds") = "Basic.Needs.Family":  d("螳ｶ譌蒐eeds") = "Basic.Needs.Family"

    ' --- 陬懷勧蜈ｷ・医メ繧ｧ繝・け・俄・ Basic.Aids.* 縺ｸ ---
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_譚・, "Basic.Aids.譚・
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_豁ｩ陦悟勣", "Basic.Aids.豁ｩ陦悟勣"
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_遏ｭ荳玖い陬・・", "Basic.Aids.遏ｭ荳玖い陬・・"
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_謇九☆繧・, "Basic.Aids.謇九☆繧・
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_繧ｷ繝ｫ繝舌・繧ｫ繝ｼ", "Basic.Aids.繧ｷ繝ｫ繝舌・繧ｫ繝ｼ"
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_霆翫＞縺・, "Basic.Aids.霆翫＞縺・: AddAlias d, "BasicInfo_陬懷勧蜈ｷ_霆頑､・ｭ・, "Basic.Aids.霆翫＞縺・
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_莉句勧繝吶Ν繝・, "Basic.Aids.莉句勧繝吶Ν繝・
    AddAlias d, "BasicInfo_陬懷勧蜈ｷ_繧ｹ繝ｭ繝ｼ繝・, "Basic.Aids.繧ｹ繝ｭ繝ｼ繝・

    ' --- 繝ｪ繧ｹ繧ｯ・医メ繧ｧ繝・け・俄・ Basic.Risk.* 縺ｸ ---
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_霆｢蛟・, "Basic.Risk.霆｢蛟・
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_遯呈・", "Basic.Risk.遯呈・"
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_菴取・､・, "Basic.Risk.菴取・､・
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_縺帙ｓ螯・, "Basic.Risk.縺帙ｓ螯・
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_隱､蝴･", "Basic.Risk.隱､蝴･"
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_隍･逖｡", "Basic.Risk.隍･逖｡"
    AddAlias d, "BasicInfo_繝ｪ繧ｹ繧ｯ_ADL菴惹ｸ・, "Basic.Risk.ADL菴惹ｸ・
    AddAlias d, "Basic.Aids.Checks", "Basic.Aids.Checks"
    AddAlias d, "Basic.Risk.Checks", "Basic.Risk.Checks"
    AddAlias d, "BasicInfo_BI.HomeEnv.Note", "Basic.HomeEnv.Note"
    
    

    ' 1) 譌｢蟄倥・繝・ム繧偵・繝ｼ繧ｸ謾ｹ蜷・
    ApplyAliasesMerge_Basic ws, d

    ' 2) 譛菴朱剞蠢・ｦ√↑蛻励′縺ｪ縺代ｌ縺ｰ霑ｽ蜉・・ave/Load縺ｮ蟇ｾ雎｡繧呈ｼ上ｌ縺ｪ縺擾ｼ・
    Dim need As Variant, mustHave As Variant
    mustHave = Array( _
        "Basic.ID", "Basic.Name", "Basic.EvalDate", "Basic.Evaluator", _
        "BI.EvaluatorJob", _
        "Basic.Age", "Basic.Sex", "Basic.PrimaryDx", "Basic.OnsetDate", _
        "Basic.CareLevel", "Basic.DementiaADL", "BI.SocialParticipation", _
        "Basic.Needs.Patient", "Basic.Needs.Family", _
        "Basic.Medical.AdmitDate", "Basic.Medical.DischargeDate", _
        "Basic.Medical.CourseNote", "Basic.Medical.ComplicationNote", _
        "Basic.HomeEnv.Checks", "Basic.HomeEnv.Note", _
        "Basic.Aids.Checks", "Basic.Risk.Checks" _
    )
    For Each need In mustHave
        EnsureHeaderExists ws, CStr(need)
    Next need
End Sub

' === 繝倥Ν繝代・ ===
Private Sub AddAlias(ByVal d As Object, ByVal src As String, ByVal dst As String)
    d(src) = dst
End Sub

' 繧ｨ繧､繝ｪ繧｢繧ｹ謾ｹ蜷搾ｼ郁｡晉ｪ∵凾縺ｯ繝槭・繧ｸ縺励※譌ｧ蛻励ｒ蜑企勁・・
Private Sub ApplyAliasesMerge_Basic(ByVal ws As Worksheet, ByVal d As Object)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim j As Long
    For j = lastCol To 1 Step -1
        Dim h As String: h = Trim$(CStr(ws.Cells(1, j).value))
        If Len(h) = 0 Then GoTo NextJ
        If d.exists(h) Then
            Dim dst As String: dst = CStr(d(h))
            Dim dstCol As Long: dstCol = modSchema.FindColByHeaderExact(ws, dst)
            If dstCol > 0 And dstCol <> j Then
                ' 繝槭・繧ｸ・育ｩｺ谺・□縺大沂繧√ｋ・・
                Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, j).End(xlUp).row
                Dim r As Long
                For r = 2 To lastRow
                    If Len(ws.Cells(r, dstCol).value) = 0 And Len(ws.Cells(r, j).value) > 0 Then
                        ws.Cells(r, dstCol).value = ws.Cells(r, j).value
                    End If
                Next r
                ws.Columns(j).Delete
            Else
                ws.Cells(1, j).value = dst
            End If
        End If
NextJ:
    Next j
End Sub

Private Sub EnsureHeaderExists(ByVal ws As Worksheet, ByVal hdr As String)
    If modSchema.FindColByHeaderExact(ws, hdr) = 0 Then
        Dim lc As Long: lc = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        ws.Cells(1, lc + 1).value = hdr
    End If
End Sub










' EvalData縺ｮID陦後ｒ隕九▽縺代ｋ・育┌縺代ｌ縺ｰ菴懊ｋ・・
' 譌｢蟄倥せ繧ｭ繝ｼ繝槭・縺ｩ縺｡繧峨↓繧ょｯｾ蠢懶ｼ咤asic.ID / BasicInfo_ID
Public Function GetOrCreateRowByID_Basic(ByVal ws As Worksheet, ByVal idVal As String) As Long
    If Len(idVal) = 0 Then Err.Raise 5, , "ID縺檎ｩｺ縺ｧ縺吶・

    Dim idCol As Long
    idCol = FindColByHeaderExact(ws, "Basic.ID")
    If idCol = 0 Then idCol = FindColByHeaderExact(ws, "BasicInfo_ID")
    If idCol = 0 Then
        ' 辟｡縺代ｌ縺ｰ Basic.ID 繧剃ｽ懊ｋ・域里蟄倥↓蜷医ｏ縺帙※OK繝ｻ蠕後〒繧ｹ繧ｭ繝ｼ繝樒ｵｱ荳蜿ｯ・・
        idCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, idCol).value = "Basic.ID"
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, idCol).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, idCol).value) = idVal Then GetOrCreateRowByID_Basic = r: Exit Function
    Next r

    r = lastRow + 1
    ws.Cells(r, idCol).value = idVal
    GetOrCreateRowByID_Basic = r
End Function











'--- 繧ｳ繝ｳ繝懊・繝・け繧ｹ縺ｫ螳牙・縺ｫ蛟､繧貞渚譏・井ｸ隕ｧ縺ｫ辟｡縺・､縺ｪ繧画悴驕ｸ謚槭↓縺吶ｋ・・---
Private Sub SetComboSafely(owner As Object, ctlName As String, ByVal v As Variant)
    Dim cb As Object  ' MSForms.ComboBox 繧・late binding 縺ｧ謇ｱ縺・
    Dim i As Long, hit As Long
    Dim s As String

    On Error Resume Next
    Set cb = FindCtlDeep(owner, ctlName)
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub

    s = CStr(v)
    hit = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i)) = s Then
            hit = i
            Exit For
        End If
    Next

    If hit >= 0 Then
        cb.ListIndex = hit               ' 荳閾ｴ縺瑚ｦ九▽縺九▲縺溘ｉ驕ｸ謚・
    Else
        cb.ListIndex = -1                ' 隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ譛ｪ驕ｸ謚槭↓・・ropDownList縺ｧ繧ょｮ牙・・・
        ' 窶ｻDropDownList縺ｮ蝣ｴ蜷医…b.Text 縺ｫ縺ｯ蜈･繧後∪縺帙ｓ
    End If
End Sub











Private Function FindControlDeep(ByVal parent As Object, ByVal targetName As String) As Object
    Dim c As Object, hit As Object

    ' 1) 閾ｪ蛻・・霄ｫ縺御ｸ閾ｴ縺ｪ繧牙叉霑斐☆
    On Error Resume Next
    If Not parent Is Nothing Then
        If parent.name = targetName Then Set FindControlDeep = parent: Exit Function
    End If
    On Error GoTo 0

    ' 2) MultiPage 縺ｯ Pages 繧定ｵｰ譟ｻ
    If TypeName(parent) = "MultiPage" Then
        Dim pg As Object
        For Each pg In parent.Pages
            Set hit = FindControlDeep(pg, targetName)
            If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function
        Next pg
        Exit Function
    End If

    ' 3) 逶ｴ荳九↓蜷悟錐縺後≠繧後・蜿門ｾ暦ｼ亥ｭ伜惠縺励↑縺・梛縺ｧ繧ゆｾ句､悶↓縺励↑縺・ｼ・
    On Error Resume Next
    Set hit = parent.controls(targetName)
    On Error GoTo 0
    If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function

    ' 4) 蟄舌さ繝ｳ繝医Ο繝ｼ繝ｫ繧貞・蟶ｰ襍ｰ譟ｻ・・ontrols 繧呈戟縺溘↑縺・梛縺ｯ繧ｹ繧ｭ繝・・・・
    On Error Resume Next
    For Each c In parent.controls
        Err.Clear
        Set hit = FindControlDeep(c, targetName)
        If Not hit Is Nothing Then Set FindControlDeep = hit: Exit Function
    Next c
    On Error GoTo 0
End Function


' 莉｣陦ｨ繧ｭ繝｣繝励す繝ｧ繝ｳ縺九ｉ隕ｪ繝輔Ξ繝ｼ繝繧呈耳螳・
Private Function FindGroupByAnyCaption(frm As Object, captions As Variant) As Object
    Dim cont As Object, c As Object, cap As Variant
    For Each cont In frm.controls
        On Error Resume Next
        ' 繧ｳ繝ｳ繝・リ・・rame/Page縺ｪ縺ｩ・峨□縺題ｪｿ縺ｹ繧・
        If Not cont.controls Is Nothing Then
            For Each c In cont.controls
                If TypeName(c) = "CheckBox" Then
                    For Each cap In captions
                        If Trim$(c.caption) = CStr(cap) Then
                            Set FindGroupByAnyCaption = cont
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    Next
End Function

' 蜷榊燕竊堤┌縺代ｌ縺ｰ莉｣陦ｨ繧ｭ繝｣繝励す繝ｧ繝ｳ縺ｧ陬懷勧蜈ｷ/繝ｪ繧ｹ繧ｯ縺ｮ繝輔Ξ繝ｼ繝繧貞叙蠕・
Private Function ResolveGroup(frm As Object, targetName As String, isAids As Boolean) As Object
    ' 1) 蜷榊燕縺ｧ謗｢縺呻ｼ郁・蜑阪・FindControlDeep繧剃ｽｿ縺・ｼ・
    Set ResolveGroup = frm.controls(targetName)
    If Not ResolveGroup Is Nothing Then Exit Function

    ' 2) 繧ｭ繝｣繝励す繝ｧ繝ｳ縺九ｉ謗ｨ螳・
    Dim seeds As Variant
    If isAids Then
        seeds = Array("譚・, "豁ｩ陦悟勣", "繧ｷ繝ｫ繝舌・繧ｫ繝ｼ", "霆翫＞縺・, "莉句勧繝吶Ν繝・, "繧ｹ繝ｭ繝ｼ繝・, "邨御ｸ玖い陬・・", "謇九☆繧・)
    Else
        seeds = Array("霆｢蛟・, "隱､蝴･", "隍･逖｡", "螟ｱ遖・, "菴取・､・, "縺帙ｓ螯・, "蠕伜ｾ・, "ADL菴惹ｸ・)
    End If
    Set ResolveGroup = FindGroupByAnyCaption(frm, seeds)
End Function

' CSV蛹厄ｼ・aption繧偵く繝ｼ・会ｼ嗾argetName縺檎┌縺上※繧ゆｻ｣陦ｨ繧ｭ繝｣繝励す繝ｧ繝ｳ縺ｧ讀懷・
Public Function SerializeChecks(frm As Object, targetName As String, Optional isAids As Boolean = True) As String
    Dim grp As Object: Set grp = ResolveGroup(frm, targetName, isAids)
    If grp Is Nothing Then Exit Function

    Dim s As String, c As Object
    For Each c In grp.controls
        If TypeName(c) = "CheckBox" Then
            If c.value = True Then
                If LenB(s) > 0 Then s = s & ","
                s = s & Trim$(c.caption)
            End If
        End If
    Next
    SerializeChecks = s
End Function

' CSV 竊・繝√ぉ繝・け蠕ｩ蜈・
Public Sub DeserializeChecks(frm As Object, targetName As String, ByVal csv As String, Optional isAids As Boolean = True)
    Dim grp As Object: Set grp = ResolveGroup(frm, targetName, isAids)
    If grp Is Nothing Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim p As Variant
    If Len(Trim$(csv)) > 0 Then
        For Each p In Split(csv, ",")
            dict(Trim$(CStr(p))) = True
        Next
    End If

    Dim c As Object
    For Each c In grp.controls
        If TypeName(c) = "CheckBox" Then
            c.value = dict.exists(Trim$(c.caption))
        End If
    Next
End Sub

' ID縺ｮ譛螟ｧ蛟､+1
Public Function NextID(ws As Worksheet, ByVal cID As Long) As Long
    Dim last As Long: last = ws.Cells(ws.rows.count, cID).End(xlUp).row
    If last < 2 Then NextID = 1: Exit Function
    On Error Resume Next
    NextID = WorksheetFunction.Max(ws.Range(ws.Cells(2, cID), ws.Cells(last, cID))) + 1
    If Err.Number <> 0 Then NextID = 1: Err.Clear
    On Error GoTo 0
End Function


Private Function GetBool(owner As Object, ctlName As String, Optional defaultValue As Boolean = True) As Boolean
    On Error Resume Next
    GetBool = CBool(owner.controls(ctlName).value)
    If Err.Number <> 0 Then GetBool = defaultValue
    On Error GoTo 0
End Function




'=== Compat: SENSE_IO 繧・IO_Sensory 縺ｫ繝溘Λ繝ｼ・郁｡・r 縺ｮ縺ｿ・・===
Private Sub Mirror_SensoryIO(ws As Worksheet, ByVal r As Long)
    Dim cSrc As Variant, cDst As Long
    cSrc = Application.Match("SENSE_IO", ws.rows(1), 0)
    If IsError(cSrc) Then Exit Sub

    ' 螳帛・繝倥ャ繝 IO_Sensory 繧堤｢ｺ菫・
    Dim m As Variant, lastCol As Long
    m = Application.Match("IO_Sensory", ws.rows(1), 0)
    If IsError(m) Then
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = "IO_Sensory"
        cDst = lastCol + 1
    Else
        cDst = CLng(m)
    End If

    ws.Cells(r, cDst).Value2 = CStr(ws.Cells(r, CLng(cSrc)).value)
End Sub


'====================================================================
' Debug / Probe 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ・・valData 縺ｮ繧ｹ繝翫ャ繝励す繝ｧ繝・ヨ繝ｻROM繝倥ャ繝遲会ｼ・
'  - 譛ｬ逡ｪ蜃ｦ逅・ｼ井ｿ晏ｭ倥・隱ｭ霎ｼ・峨°繧峨・逶ｴ謗･蜻ｼ縺ｰ縺ｪ縺・
'  - 蠢・ｦ√↑縺ｨ縺阪□縺代！mmediate 繧・ｰら畑繝・せ繝医・繧ｯ繝ｭ縺九ｉ謇句虚縺ｧ蜻ｼ縺ｳ蜃ｺ縺・
'  - 蟆・擂逧・↓縺ｯ modEvalIODebug 縺ｪ縺ｩ蛻･繝｢繧ｸ繝･繝ｼ繝ｫ縺ｸ蛻・ｊ蜃ｺ縺吝呵｣・
'====================================================================



Public Sub Debug_IO_Sensory_ADL_Snapshot(ByVal ws As Worksheet, ByVal r As Long)
#If APP_DEBUG Then
    Dim s As String

    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[SENSE][IO]", _
                "row=" & r, _
                "| len=" & Len(s), _
                "| head=" & Left$(s, 200)

    s = ReadStr_Compat("IO_ADL", r, ws)
    Debug.Print "[ADL][IO]", _
                "row=" & r, _
                "| len=" & Len(s), _
                "| head=" & Left$(s, 200)
#End If
End Sub



Public Sub Debug_Sensory_ADL_Raw(ByVal ws As Worksheet, ByVal r As Long)
    Dim cSense As Variant, cADL As Variant, cIOSense As Variant
    Dim lastCol As Long

    cSense = Application.Match("SENSE_IO", ws.rows(1), 0)
    cADL = Application.Match("IO_ADL", ws.rows(1), 0)
    cIOSense = Application.Match("IO_Sensory", ws.rows(1), 0)

    Debug.Print "=== [RAW SENSE/ADL] row=" & r & " ==="

    If Not IsError(cSense) Then
        Debug.Print "SENSE_IO(col" & cSense & ") =", ws.Cells(r, cSense).text
    Else
        Debug.Print "SENSE_IO: <no header>"
    End If

    If Not IsError(cADL) Then
        Debug.Print "IO_ADL(col" & cADL & ") =", ws.Cells(r, cADL).text
    Else
        Debug.Print "IO_ADL: <no header>"
    End If

    If Not IsError(cIOSense) Then
        Debug.Print "IO_Sensory(col" & cIOSense & ") =", ws.Cells(r, cIOSense).text
    Else
        Debug.Print "IO_Sensory: <no header>"
    End If

    ' 霑大ｍ遒ｺ隱搾ｼ域ｧ矩隕九ｋ逕ｨ・・
    Debug.Print "SENSE霑大ｍ(146-155)=", Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 146), ws.Cells(r, 155)).value)), " | ")
    Debug.Print "ADL霑大ｍ  (156-165)=", Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 156), ws.Cells(r, 165)).value)), " | ")

    Debug.Print "=== [/RAW SENSE/ADL] ==="
End Sub







Public Sub Debug_Cols_146_165_Headers(ByVal ws As Worksheet)
    Dim c As Long
    Debug.Print "=== [HEADERS 146-165] ==="
    For c = 146 To 165
        Debug.Print c, ws.Cells(1, c).value
    Next c
    Debug.Print "=== [/HEADERS] ==="
End Sub



Public Sub Debug_Find_IO_Sense_ADL_Sample(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim sSense As String
    Dim sADL As String

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        sSense = Trim$(ws.Cells(r, 152).value) 'SENSE_IO
        sADL = Trim$(ws.Cells(r, 159).value)   'IO_ADL

        If Len(sSense) > 0 Or Len(sADL) > 0 Then
            Debug.Print "=== [Found IO Sample] row=" & r & " ==="
            Debug.Print "SENSE_IO:", Left$(sSense, 200)
            Debug.Print "IO_ADL:", Left$(sADL, 200)
            Exit For
        End If
    Next r

    If r > lastRow Then
        Debug.Print "=== [No SENSE_IO / IO_ADL found] ==="
    End If
End Sub



Public Sub Debug_ListROMHeaders()
    Dim ws As Worksheet
    Dim c As Long
    Dim firstCol As Long, lastCol As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")

    ' ROM邉ｻ縺御ｸｦ繧薙〒縺・ｋ諠ｳ螳壹Ξ繝ｳ繧ｸ縺縺代ｒ隕九ｋ・亥ｿ・ｦ√↑繧牙ｾ後〒蠕ｮ隱ｿ謨ｴ・・
    firstCol = 150
    lastCol = 260

    Debug.Print "=== [ROM_* HEADERS 150-260] ==="
    For c = firstCol To lastCol
        If LCase$(Left$(CStr(ws.Cells(1, c).value), 4)) = "rom_" Then
            Debug.Print c, ws.Cells(1, c).value
        End If
    Next c
    Debug.Print "=== [/ROM_* HEADERS] ==="
End Sub



Public Sub Debug_ROMRow_Values(ByVal r As Long)
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim h As String, v As String
    Dim hit As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Debug.Print "=== [ROM VALUES row=" & r & "] ==="
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If LCase$(Left$(h, 4)) = "rom_" Then
            v = CStr(ws.Cells(r, c).value)
            If Len(v) > 0 Then
                Debug.Print c, h, v
                hit = hit + 1
                If hit >= 40 Then Exit For   ' 繝ｭ繧ｰ證ｴ逋ｺ髦ｲ豁｢
            End If
        End If
    Next c

    If hit = 0 Then
        Debug.Print "(no ROM_* values found)"
    End If
    Debug.Print "=== [/ROM VALUES] ==="
End Sub



Public Sub Debug_Find_IO_ROM_Header()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim h As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Debug.Print "=== [FIND IO_ROM HEADER] ==="
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).value)
        If InStr(1, h, "IO_ROM", vbTextCompare) > 0 Then
            Debug.Print c, "[" & h & "]"
        End If
    Next c
    Debug.Print "=== [/FIND IO_ROM HEADER] ==="
End Sub



Public Sub Debug_Find_ROM_SampleRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim h As String, v As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    Debug.Print "=== [ROM SAMPLE ROW SEARCH] ==="
    For r = 2 To lastRow
        For c = 158 To 260
            h = CStr(ws.Cells(1, c).value)
            If LCase$(Left$(h, 4)) = "rom_" Then
                v = CStr(ws.Cells(r, c).value)
                If Len(v) > 0 Then
                    Debug.Print "row=" & r & ", col=" & c & ", header=" & h
                    Debug.Print "=== [/ROM SAMPLE ROW SEARCH] ==="
                    Exit Sub
                End If
            End If
        Next c
    Next r

    Debug.Print "(no ROM_* values found in 158-260)"
    Debug.Print "=== [/ROM SAMPLE ROW SEARCH] ==="
End Sub



Public Sub Cleanup_ExtraROMColumns()
    Dim ws As Worksheet
    Dim lastCol As Long, c As Long
    Dim h As String

    Set ws = ThisWorkbook.Worksheets("EvalData")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' 譛ｬ譚･菴ｿ縺・OM繝悶Ο繝・け繧医ｊ蜿ｳ蛛ｴ縺縺代ｒ繧ｴ繝溷呵｣懊→縺吶ｋ・医→繧翫≠縺医★300蛻嶺ｻ･髯搾ｼ・
    For c = lastCol To 261 Step -1
        h = CStr(ws.Cells(1, c).value)
        If LCase$(Left$(h, 4)) = "rom_" Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub




Public Function Build_TestEval_IO(owner As Object) As String
    Dim s As String
    Dim v10 As String
    Dim vTUG As String
    Dim v5x As String
    Dim vSemi As String
    Dim vGripR As String
    Dim vGripL As String

    With owner
        v10 = Trim$(.txtTenMWalk.value)
        vTUG = Trim$(.txtTUG.value)
        v5x = Trim$(.txtFiveSts.value)   ' 窶ｻ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ蜷阪′驕輔≧蝣ｴ蜷医・縺薙％縺縺題ｪｿ謨ｴ
        vSemi = Trim$(.txtSemi.value)
        vGripR = Trim$(.txtGripR.value)
        vGripL = Trim$(.txtGripL.value)
    End With

    s = "Test_10MWalk_sec=" & v10
    s = s & "|Test_TUG_sec=" & vTUG
    s = s & "|Test_Grip_R_kg=" & vGripR
    s = s & "|Test_Grip_L_kg=" & vGripL
    s = s & "|Test_5xSitStand_sec=" & v5x
    s = s & "|Test_SemiTandem_sec=" & vSemi

    Build_TestEval_IO = s
End Function





Public Sub Save_TestEvalToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_TestEval 逕ｨ縺ｮ蛻励ｒ遒ｺ菫・
    c = EnsureHeader(ws, "IO_TestEval")

    ' 繝輔か繝ｼ繝荳翫・蛟､縺九ｉ IO 譁・ｭ怜・繧堤函謌撰ｼ井ｻ翫・遨ｺ縺ｮ縺ｾ縺ｾ縺ｧ繧０K・・
    s = Build_TestEval_IO(owner)

        ' 謖・ｮ夊｡後↓荳頑嶌縺堺ｿ晏ｭ・
    ws.Cells(r, c).Value2 = CStr(s)
    SaveTestEvalMemoColumns ws, r, owner
    ws.Cells(r, 181).value = val(owner.txtTUG.value)


End Sub




Public Sub Load_TestEvalFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
      Dim s As String
    s = ReadStr_Compat("IO_TestEval", r, ws)

    If Len(s) > 0 Then s = Replace(s, "=", ": ")

    
    owner.txtTenMWalk.value = IO_GetVal(s, "Test_10MWalk_sec")
    owner.txtTUG.value = IO_GetVal(s, "Test_TUG_sec")
    owner.txtFiveSts.value = IO_GetVal(s, "Test_5xSitStand_sec")
    owner.txtGripR.value = IO_GetVal(s, "Test_Grip_R_kg")
    owner.txtGripL.value = IO_GetVal(s, "Test_Grip_L_kg")
    owner.txtSemi.value = IO_GetVal(s, "Test_SemiTandem_sec")
    LoadTestEvalMemoColumns ws, r, owner

    ' TODO: 縺薙％縺九ｉ荳九・蠕後〒螳溯｣・ｼ井ｻ翫・隗ｦ繧峨↑縺・ｼ・
    ' IO_TestEval 繧貞・隗｣縺励※
    ' owner・・rmEval・峨・ txtTenMWalk / txtTUG / txtFiveSts /
    ' txtGripR / txtGripL / txtSemi 縺ｫ豬√＠霎ｼ繧
    
    
    ws.Cells(r, 181).value = val(owner.txtTUG.value)
   
End Sub

Private Sub SaveTestEvalMemoColumns(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_10mWalk", "txtMemo_10mWalk"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_TUG", "txtMemo_TUG"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_STS5", "txtMemo_STS5"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_SemiTandem", "txtMemo_SemiTandem"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripR", "txtMemo_GripR"
    SaveTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripL", "txtMemo_GripL"
End Sub

Private Sub LoadTestEvalMemoColumns(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_10mWalk", "txtMemo_10mWalk"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_TUG", "txtMemo_TUG"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_STS5", "txtMemo_STS5"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_SemiTandem", "txtMemo_SemiTandem"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripR", "txtMemo_GripR"
    LoadTestEvalMemoColumn ws, r, owner, "TestEval_Memo_GripL", "txtMemo_GripL"
End Sub

Private Sub SaveTestEvalMemoColumn(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, _
                                   ByVal header As String, ByVal ctlName As String)
    Dim c As Long
    c = EnsureHeader(ws, header)
    ws.Cells(r, c).Value2 = GetCtlTextGeneric(owner, ctlName)
End Sub

Private Sub LoadTestEvalMemoColumn(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object, _
                                   ByVal header As String, ByVal ctlName As String)
    Dim c As Long
    c = FindColByHeaderExact(ws, header)
    If c = 0 Then Exit Sub

    SetCtlValueSafe owner, ctlName, CStr(ws.Cells(r, c).Value2)
End Sub

Public Sub Load_WalkIndepFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim v As String
    Dim cmb As Object
    Dim c As Object
    Dim parts() As String
    Dim i As Long
    Dim nm As String
    Dim vLevel As String       '笘・霑ｽ蜉・夊・遶句ｺｦ
    Dim cLvl As Object         '笘・霑ｽ蜉・夊・遶句ｺｦ繧ｳ繝ｳ繝懃畑


       ' IO_WalkIndep 縺ｮ譁・ｭ怜・繧貞叙蠕・
    s = ReadStr_Compat("IO_WalkIndep", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' TestEval 縺ｨ蜷後§繝代ち繝ｼ繝ｳ縺ｫ蜷医ｏ縺帙※ "Key=Val" 竊・"Key: Val" 縺ｫ螟牙ｽ｢
    s = Replace(s, "=", ": ")

    ' --- 閾ｪ遶句ｺｦ・・alk_IndepLevel・・---
    vLevel = IO_GetVal(s, "Walk_IndepLevel")
    If Len(vLevel) > 0 Then
        ' Tag="WalkIndepLevel" 縺ｮ繧ｳ繝ｳ繝懊ｒ謗｢縺励※蛟､繧呈綾縺・
        Set cLvl = Nothing
        For Each c In owner.controls
            If TypeName(c) = "ComboBox" Then
                If c.tag = "WalkIndepLevel" Then
                    Set cLvl = c
                    Exit For
                End If
            End If
        Next c
        If Not cLvl Is Nothing Then
            cLvl.value = vLevel
        End If
    End If

    ' --- 霍晞屬 ---
    v = IO_GetVal(s, "Walk_Distance")
    Set cmb = FindControlRecursive(owner, "cmbWalkDistance")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 螻句､・---
    v = IO_GetVal(s, "Walk_Outdoor")
    Set cmb = FindControlRecursive(owner, "cmbWalkOutdoor")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 騾溷ｺｦ ---
    v = IO_GetVal(s, "Walk_Speed")
    Set cmb = FindControlRecursive(owner, "cmbWalkSpeed")
    If cmb Is Nothing Then Set cmb = FindControlByTagRecursive(owner, "cmbGaitSpeedDetail")
    If Not cmb Is Nothing Then cmb.value = v

    ' --- 螳牙ｮ壽ｧ繝√ぉ繝・け・・hkWalkStab_*・峨ｒ荳蠎ｦ蜈ｨ驛ｨOFF ---
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If StrComp(Left$(nm, 12), "chkWalkStab_", vbTextCompare) = 0 Then
                c.value = False
            End If
        End If
    Next c

    ' --- 螳牙ｮ壽ｧ縺ｮ菫晏ｭ俶枚蟄怜・繧貞ｱ暮幕縺励※縲∬ｩｲ蠖薙メ繧ｧ繝・け繧丹N ---
    v = IO_GetVal(s, "Walk_Stab")   ' 萓具ｼ・"chkWalkStab_Furatsuki/chkWalkStab_FallRisk"
    If Len(v) > 0 Then
        parts = Split(v, "/")
        For i = LBound(parts) To UBound(parts)
            nm = Trim$(parts(i))
            If Len(nm) > 0 Then
                Set c = FindControlRecursive(owner, nm)
                If Not c Is Nothing Then
                    If TypeName(c) = "CheckBox" Then c.value = True
                End If
            End If
        Next i
    End If

    v = IO_GetVal(s, "Walk_Assistive")
    DeserializeCheckedCaptionsByTag owner, "AssistiveGroup", v
End Sub

Public Sub Load_WalkRLAFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim phases As Variant
    Dim phase As Variant
    Dim probs As String
    Dim level As String
    Dim parts() As String
    Dim i As Long
    Dim c As Object
    Dim nm As String

    ' IO_WalkRLA 縺ｮ譁・ｭ怜・繧貞叙蠕・
    s = ReadStr_Compat("IO_WalkRLA", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' 縺ｾ縺壹ヽLA 髢｢騾｣縺ｮ繝√ぉ繝・け繝ｻ繝ｬ繝吶Ν繧貞・驛ｨ繝ｪ繧ｻ繝・ヨ
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "RLA_", vbTextCompare) = 1 Then
                c.value = False
            End If
        ElseIf TypeName(c) = "OptionButton" Then
            If InStr(1, c.groupName, "IC", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "LR", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "MSt", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "TSt", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "PSw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "ISw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "MSw", vbTextCompare) = 1 _
               Or InStr(1, c.groupName, "TSw", vbTextCompare) = 1 Then
                c.value = False
            End If
        End If
    Next c

    ' TestEval 縺ｨ蜷後§繝代ち繝ｼ繝ｳ縺ｫ蜷医ｏ縺帙※ "Key=Val" 竊・"Key: Val"
    s = Replace(s, "=", ": ")

    ' 遶玖・譛滂ｼ矩♀閼壽悄縺ｮ繧ｭ繝ｼ
    phases = Array("IC", "LR", "MSt", "TSt", "PSw", "ISw", "MSw", "TSw")

    For Each phase In phases
        ' Problems 縺ｨ Level 繧貞叙繧雁・縺・
        probs = IO_GetVal(s, "RLA_" & CStr(phase) & "_Problems")
        level = IO_GetVal(s, "RLA_" & CStr(phase) & "_Level")

        ' --- 蝠城｡鯉ｼ・heckBox・咾aption荳閾ｴ縺ｧON・・---
        If Len(probs) > 0 Then
            parts = Split(probs, "/")
            For i = LBound(parts) To UBound(parts)
                Dim cap As String
                cap = Trim$(parts(i))
                If Len(cap) > 0 Then
                    For Each c In owner.controls
                        If TypeName(c) = "CheckBox" Then
                            nm = CStr(c.name)
                            If InStr(1, nm, "RLA_" & CStr(phase) & "_", vbTextCompare) = 1 Then
                                If CStr(c.caption) = cap Then
                                    c.value = True
                                End If
                            End If
                        End If
                    Next c
                End If
            Next i
        End If

        ' --- 繝ｬ繝吶Ν・・ptionButton・哦roupName=phase & Caption荳閾ｴ縺ｧON・・---
        If Len(level) > 0 Then
            For Each c In owner.controls
                If TypeName(c) = "OptionButton" Then
                    If StrComp(c.groupName, CStr(phase), vbTextCompare) = 0 Then
                        If CStr(c.caption) = level Then
                            c.value = True
                        End If
                    End If
                End If
            Next c
        End If
    Next phase
End Sub



Public Sub Load_WalkAbnFromSheet(ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim s As String
    Dim v As String
    Dim parts() As String
    Dim i As Long
    Dim nm As String
    Dim c As Object

    ' IO_WalkAbn 縺ｮ譁・ｭ怜・蜿門ｾ・
    s = ReadStr_Compat("IO_WalkAbn", r, ws)

    If Len(s) = 0 Then Exit Sub

    ' 荳譌ｦ縲’raWalkAbn_* 縺ｮ蜈ｨ繝√ぉ繝・け繧丹FF縺ｫ縺吶ｋ
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "fraWalkAbn_", vbTextCompare) = 1 Then
                c.value = False
            End If
        End If
    Next c

    ' s 縺ｮ荳ｭ霄ｫ・井ｾ具ｼ・"fraWalkAbn_A_chk0|fraWalkAbn_C_chk3"・峨ｒ螻暮幕
    parts = Split(s, "|")
    For i = LBound(parts) To UBound(parts)
        nm = Trim$(parts(i))
        If Len(nm) > 0 Then
            Set c = FindControlRecursive(owner, nm)
            If Not c Is Nothing Then
                If TypeName(c) = "CheckBox" Then
                    c.value = True
                End If
            End If
        End If
    Next i
End Sub





Public Sub Save_WalkIndepToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkIndep 逕ｨ縺ｮ蛻励ｒ遒ｺ菫・
    c = EnsureHeader(ws, "IO_WalkIndep")

    ' 繝輔か繝ｼ繝荳翫・蛟､縺九ｉ IO 譁・ｭ怜・繧堤函謌・
    s = Build_WalkIndep_IO(owner)

    ' 謖・ｮ夊｡後↓荳頑嶌縺堺ｿ晏ｭ・
    ws.Cells(r, c).Value2 = CStr(s)

End Sub



Private Function FindControlRecursive(parent As Object, name As String) As Object
    Dim ctl As Object
    For Each ctl In parent.controls
        If StrComp(ctl.name, name, vbTextCompare) = 0 Then
            Set FindControlRecursive = ctl
            Exit Function
        End If
        ' Frame 繧・MultiPage 縺ｮ蝣ｴ蜷医・蜀榊ｸｰ讀懃ｴ｢
        On Error Resume Next
        If ctl.controls.count > 0 Then
            Dim subCtl As Object
            Set subCtl = FindControlRecursive(ctl, name)
            If Not subCtl Is Nothing Then
                Set FindControlRecursive = subCtl
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next ctl
End Function

Private Function FindControlByTagRecursive(parent As Object, tagName As String) As Object
    Dim ctl As Object
    For Each ctl In parent.controls
        If StrComp(CStr(ctl.tag), tagName, vbTextCompare) = 0 Then
            Set FindControlByTagRecursive = ctl
            Exit Function
        End If
        On Error Resume Next
        If ctl.controls.count > 0 Then
            Dim subCtl As Object
            Set subCtl = FindControlByTagRecursive(ctl, tagName)
            If Not subCtl Is Nothing Then
                Set FindControlByTagRecursive = subCtl
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next ctl
End Function

Private Function SerializeCheckedCaptionsByTag(parent As Object, groupTag As String) As String
    Dim ctl As Object
    Dim s As String
    Dim childCsv As String

    For Each ctl In parent.controls
        If TypeName(ctl) = "CheckBox" Then
            If StrComp(CStr(ctl.tag), groupTag, vbTextCompare) = 0 Then
                If ctl.value = True Then
                    If LenB(s) > 0 Then s = s & ","
                    s = s & Trim$(ctl.caption)
                End If
            End If
        End If

        On Error Resume Next
        If ctl.controls.count > 0 Then
            childCsv = SerializeCheckedCaptionsByTag(ctl, groupTag)
            If LenB(childCsv) > 0 Then
                If LenB(s) > 0 Then s = s & ","
                s = s & childCsv
            End If
        End If
        On Error GoTo 0
    Next ctl

    SerializeCheckedCaptionsByTag = s
End Function

Private Sub DeserializeCheckedCaptionsByTag(parent As Object, groupTag As String, ByVal csv As String)
    Dim dict As Object
    Dim parts As Variant
    Dim p As Variant
    Dim ctl As Object

    Set dict = CreateObject("Scripting.Dictionary")
    If Len(Trim$(csv)) > 0 Then
        parts = Split(csv, ",")
        For Each p In parts
            dict(Trim$(CStr(p))) = True
        Next p
    End If

    For Each ctl In parent.controls
        If TypeName(ctl) = "CheckBox" Then
            If StrComp(CStr(ctl.tag), groupTag, vbTextCompare) = 0 Then
                ctl.value = dict.exists(Trim$(ctl.caption))
            End If
        End If

        On Error Resume Next
        If ctl.controls.count > 0 Then
            DeserializeCheckedCaptionsByTag ctl, groupTag, csv
        End If
        On Error GoTo 0
    Next ctl
End Sub


Public Function Build_WalkIndep_IO(owner As Object) As String
    Dim vDist As String
    Dim vOut As String
    Dim vSpeed As String
    Dim vAssistive As String
    Dim s As String
    Dim hits As Collection
    Dim c As Object
    Dim nm As String
    Dim stab As String
    Dim i As Long
    Dim vLevel As String   '笘・閾ｪ遶句ｺｦ



   Dim cLvl As Object
Set cLvl = FindControlRecursive(owner, "cmbWalkIndep")
If cLvl Is Nothing Then
    ' 繧ｿ繧ｰ縺ｧ讀懃ｴ｢縺吶ｋ・井ｻ雁屓縺ｮ豁｣蠑上Ν繝ｼ繝茨ｼ・
    For Each c In owner.controls
        If TypeName(c) = "ComboBox" Then
            If c.tag = "WalkIndepLevel" Then
                Set cLvl = c
                Exit For
            End If
        End If
    Next c
End If
If Not cLvl Is Nothing Then vLevel = Trim$(cLvl.value)



    ' 霍晞屬繝ｻ螻句､悶・騾溷ｺｦ
    Set c = FindControlRecursive(owner, "cmbWalkDistance")
    If Not c Is Nothing Then vDist = Trim$(c.value)

    Set c = FindControlRecursive(owner, "cmbWalkOutdoor")
    If Not c Is Nothing Then vOut = Trim$(c.value)

    Set c = FindControlRecursive(owner, "cmbWalkSpeed")
    If c Is Nothing Then Set c = FindControlByTagRecursive(owner, "cmbGaitSpeedDetail")
    If Not c Is Nothing Then vSpeed = Trim$(c.value)
    

    ' 螳牙ｮ壽ｧ繝√ぉ繝・け・・hkWalkStab_・・繧貞・驛ｨ諡ｾ縺・ｼ・
    Set hits = New Collection
    For Each c In owner.controls
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If StrComp(Left$(nm, 12), "chkWalkStab_", vbTextCompare) = 0 Then
                If c.value = True Then
                    ' 蜷榊燕縺昴・繧ゅ・縺九∵忰蟆ｾ縺縺代↓縺吶ｋ縺九・縺ゅ→縺ｧ隱ｿ謨ｴ蜿ｯ
                    hits.Add nm
                End If
            End If
        End If
    Next c

    ' 螳牙ｮ壽ｧ縺ｮ繝√ぉ繝・け蜷阪ｒ縲・縲榊玄蛻・ｊ縺ｧ1譛ｬ縺ｮ譁・ｭ怜・縺ｫ縺ｾ縺ｨ繧√ｋ
    For i = 1 To hits.count
        If i > 1 Then stab = stab & "/"
        stab = stab & hits(i)
    Next i

        ' IO 譁・ｭ怜・邨・∩遶九※
    s = "Walk_IndepLevel=" & vLevel
    s = s & "|Walk_Distance=" & vDist
    s = s & "|Walk_Outdoor=" & vOut
    s = s & "|Walk_Stab=" & stab
    s = s & "|Walk_Speed=" & vSpeed

    vAssistive = SerializeCheckedCaptionsByTag(owner, "AssistiveGroup")
    s = s & "|Walk_Assistive=" & vAssistive

    Build_WalkIndep_IO = s
End Function



Public Function Build_WalkAbn_IO(owner As Object) As String
    Dim c As Object
    Dim hits As Collection
    Dim s As String
    Dim nm As String
    
    Set hits = New Collection
    
    For Each c In owner.controls
        ' fraWalkAbn_?_chk? 縺ｨ縺・≧蜷榊燕縺ｮ CheckBox 縺縺第鏡縺・
        If TypeName(c) = "CheckBox" Then
            nm = CStr(c.name)
            If InStr(1, nm, "fraWalkAbn_", vbTextCompare) = 1 Then
                If c.value = True Then
                    hits.Add nm
                End If
            End If
        End If
    Next c
    
    ' 1縺､繧ゅメ繧ｧ繝・け縺檎┌縺代ｌ縺ｰ遨ｺ譁・ｭ励ｒ霑斐☆
    If hits.count = 0 Then
        Build_WalkAbn_IO = ""
        Exit Function
    End If
    
    ' fraWalkAbn_A_chk0|fraWalkAbn_A_chk3|窶ｦ 縺ｨ縺・≧蠖｢縺ｧ騾｣邨・
    Dim i As Long
    For i = 1 To hits.count
        If i > 1 Then s = s & "|"
        s = s & hits(i)
    Next i
    
    Build_WalkAbn_IO = s
End Function


Public Sub Save_WalkAbnToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkAbn 逕ｨ縺ｮ蛻励ｒ遒ｺ菫・
    c = EnsureHeader(ws, "IO_WalkAbn")

    ' 繝輔か繝ｼ繝縺ｮ繝√ぉ繝・け迥ｶ諷九°繧・IO 譁・ｭ怜・繧堤函謌・
    s = Build_WalkAbn_IO(owner)

    ' 謖・ｮ夊｡後↓荳頑嶌縺堺ｿ晏ｭ・
    ws.Cells(r, c).Value2 = CStr(s)

End Sub




Public Function Build_WalkRLA_IO(owner As Object) As String
    Dim phases As Variant
    Dim phase As Variant
    Dim c As Object
    Dim probs As Collection
    Dim probsStr As String
    Dim level As String
    Dim s As String
    Dim first As Boolean
    Dim i As Long
    Dim nm As String

    ' 遶玖・譛滂ｼ矩♀閼壽悄縺ｮ繧ｭ繝ｼ・・uild_RLA_ChecksPart 縺ｨ蜷後§・・
    phases = Array("IC", "LR", "MSt", "TSt", "PSw", "ISw", "MSw", "TSw")
    first = True

    For Each phase In phases
        Set probs = New Collection
        probsStr = ""
        level = ""

        ' --- 繝√ぉ繝・け・・LA_<phase>_・橸ｼ峨ｒ諡ｾ縺・---
        For Each c In owner.controls
            If TypeName(c) = "CheckBox" Then
                nm = CStr(c.name)
                If InStr(1, nm, "RLA_" & CStr(phase) & "_", vbTextCompare) = 1 Then
                    If c.value = True Then
                        probs.Add c.caption   ' 萓具ｼ牙庄蜍募沺荳崎ｶｳ / 遲句鴨菴惹ｸ・縺ｪ縺ｩ
                    End If
                End If
            End If
        Next c

        ' 蝠城｡後Μ繧ｹ繝医ｒ "/" 蛹ｺ蛻・ｊ縺ｧ 1 譛ｬ縺ｫ縺吶ｋ
        If probs.count > 0 Then
            For i = 1 To probs.count
                If i > 1 Then probsStr = probsStr & "/"
                probsStr = probsStr & probs(i)
            Next i
        End If

        ' --- 繝ｬ繝吶Ν・・ptionButton, GroupName=phase・峨ｒ諡ｾ縺・---
        For Each c In owner.controls
            If TypeName(c) = "OptionButton" Then
                If StrComp(c.groupName, CStr(phase), vbTextCompare) = 0 Then
                    If c.value = True Then
                        level = CStr(c.caption)   ' 霆ｽ蠎ｦ / 荳ｭ遲牙ｺｦ / 鬮伜ｺｦ
                        Exit For
                    End If
                End If
            End If
        Next c

        ' --- IO 繧ｻ繧ｰ繝｡繝ｳ繝育ｵ・∩遶九※ ---
        Dim seg As String
        seg = "RLA_" & CStr(phase) & "_Problems=" & probsStr & _
              "|RLA_" & CStr(phase) & "_Level=" & level

        If first Then
            s = seg
            first = False
        Else
            s = s & "|" & seg
        End If
    Next phase

    Build_WalkRLA_IO = s
End Function



Public Sub Save_WalkRLAToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)
    Dim c As Long
    Dim s As String

    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    ' IO_WalkRLA 逕ｨ縺ｮ蛻励ｒ遒ｺ菫晢ｼ亥・4縺ｫ繝倥ャ繝 IO_WalkRLA 縺後≠繧句燕謠撰ｼ・
    c = EnsureHeader(ws, "IO_WalkRLA")

    ' 繝輔か繝ｼ繝荳翫・RLA繝√ぉ繝・け繝ｻ繝ｬ繝吶Ν縺九ｉIO譁・ｭ怜・繧堤函謌・
    s = Build_WalkRLA_IO(owner)

    ' 謖・ｮ夊｡後↓荳頑嶌縺堺ｿ晏ｭ・
    ws.Cells(r, c).Value2 = CStr(s)

End Sub

Private Function GetCogTabsSafe(ByVal owner As Object) As Object
    Dim mp As Object

    On Error Resume Next
    Set mp = owner.GetCogTabs
    On Error GoTo 0
    
    Set GetCogTabsSafe = mp
    If Not mp Is Nothing Then Set GetCogTabsSafe = mp
End Function



Public Sub Save_CognitionMental_AtRow(ws As Worksheet, r As Long, owner As Object)
    Dim frm As Object
    Dim col As Long
    Dim v As Variant
    Dim f As MSForms.Frame
    Dim c As MSForms.Control
    Dim bpsd As String
    Dim mpCog As Object
    Dim pgCog As Object
    Dim pgMental As Object
    
    Set frm = owner   ' frmEval 繧貞女縺大叙繧区Φ螳・
    Set mpCog = GetCogTabsSafe(frm)
    If mpCog Is Nothing Then Exit Sub
    Set pgCog = mpCog.Pages("pgCognition")
    Set pgMental = mpCog.Pages("pgMental")
        
    
    '=== 隱咲衍・壻ｸｭ譬ｸ6鬆・岼 =====================================
    
    ' 險俶・
    col = HeaderCol_Compat("IO_Cog_Memory", ws)
    If col > 0 Then
        v = pgCog.controls("cmbCogMemory").value


        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 豕ｨ諢・
    col = HeaderCol_Compat("IO_Cog_Attention", ws)
    If col > 0 Then
        v = pgCog.controls("cmbCogAttention").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 隕句ｽ楢ｭ・
    col = HeaderCol_Compat("IO_Cog_Orientation", ws)
    If col > 0 Then
            v = pgCog.controls("cmbCogOrientation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 蛻､譁ｭ
    col = HeaderCol_Compat("IO_Cog_Judgment", ws)
    If col > 0 Then
            v = pgCog.controls("cmbCogJudgement").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 驕り｡梧ｩ溯・
    col = HeaderCol_Compat("IO_Cog_Executive", ws)
    If col > 0 Then
             v = pgCog.controls("cmbCogExecutive").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 險隱・
    col = HeaderCol_Compat("IO_Cog_Language", ws)
    If col > 0 Then
             v = pgCog.controls("cmbCogLanguage").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    '=== 隱咲衍・夊ｪ咲衍逞・・遞ｮ鬘橸ｼ句ｙ閠・==============================
    
    col = HeaderCol_Compat("IO_Cog_DementiaType", ws)
    If col > 0 Then
             v = pgCog.controls("cmbDementiaType").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    col = HeaderCol_Compat("IO_Cog_DementiaNote", ws)
    If col > 0 Then
            v = pgCog.controls("txtDementiaNote").text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
        '=== 隱咲衍・咤PSD・医メ繧ｧ繝・け縺悟・縺｣縺ｦ縺・ｋ鬆・岼繧・| 蛹ｺ蛻・ｊ縺ｧ菫晏ｭ假ｼ・===
    
    bpsd = ""
     With pgCog
        For Each c In .controls
            If TypeName(c) = "CheckBox" Then
                If c.value = True Then
                    If Len(bpsd) > 0 Then bpsd = bpsd & "|"
                    bpsd = bpsd & CStr(c.caption)
                End If
            End If
        Next c
    End With
    
    col = HeaderCol_Compat("IO_Cog_BPSD", ws)
    If col > 0 Then
        ws.Cells(r, col).value = bpsd
    End If

    
    '=== 邊ｾ逾樣擇繧ｿ繝・============================================
    
    ' 豌怜・
    col = HeaderCol_Compat("IO_Mental_Mood", ws)
    If col > 0 Then
             v = pgMental.controls("cmbMood").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 諢乗ｬｲ
    col = HeaderCol_Compat("IO_Mental_Motivation", ws)
    If col > 0 Then
            v = pgMental.controls("cmbMotivation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 荳榊ｮ・
    col = HeaderCol_Compat("IO_Mental_Anxiety", ws)
    If col > 0 Then
            v = pgMental.controls("cmbAnxiety").value
            
        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 蟇ｾ莠ｺ髢｢菫・
    col = HeaderCol_Compat("IO_Mental_Relation", ws)
    If col > 0 Then
            v = pgMental.controls("cmbRelation").value

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 逹｡逵
    col = HeaderCol_Compat("IO_Mental_Sleep", ws)
    If col > 0 Then
            v = pgMental.controls("cmbSleep").value
            
        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
    
    ' 邊ｾ逾樣擇繝ｻ蛯呵・
    col = HeaderCol_Compat("IO_Mental_Note", ws)
    If col > 0 Then
            v = pgMental.controls("txtMentalNote").text

        If IsNull(v) Then v = ""
        ws.Cells(r, col).value = v
    End If
End Sub

   



Public Sub Load_CognitionMental_FromRow(ws As Worksheet, ByVal r As Long, owner As Object)

    Dim f As MSForms.Frame
    Dim mp As Object
    Dim pgCog As Object
    Dim pgMental As Object

    Dim v As Variant
    Dim s As String
    Dim arr() As String
    Dim i As Long, j As Long
    Dim chk As MSForms.CheckBox

    '=== UI 繝ｫ繝ｼ繝亥叙蠕暦ｼ育ｵｶ蟇ｾ蜷阪ｒ蜑肴署・・==
    Set mp = GetCogTabsSafe(owner)
    If mp Is Nothing Then Exit Sub
    Set pgCog = mp.Pages("pgCognition")
    Set pgMental = mp.Pages("pgMental")

    '=== 隱咲衍蛛ｴ combobox 鄒､ ===
    LoadComboValueByHeader ws, r, "IO_Cog_Memory", pgCog, "cmbCogMemory"
    LoadComboValueByHeader ws, r, "IO_Cog_Attention", pgCog, "cmbCogAttention"
    LoadComboValueByHeader ws, r, "IO_Cog_Orientation", pgCog, "cmbCogOrientation"
    LoadComboValueByHeader ws, r, "IO_Cog_Judgment", pgCog, "cmbCogJudgement"
    LoadComboValueByHeader ws, r, "IO_Cog_Executive", pgCog, "cmbCogExecutive"
    LoadComboValueByHeader ws, r, "IO_Cog_Language", pgCog, "cmbCogLanguage"
    LoadComboValueByHeader ws, r, "IO_Cog_DementiaType", pgCog, "cmbDementiaType"

    v = ReadValueByCompatHeader(ws, r, "IO_Cog_DementiaNote")
    pgCog.controls("txtDementiaNote").text = CStr(v)

    '=== BPSD・・hkBPSD0?10・・==
    ' 1) 蜈ｨ驛ｨ荳蠎ｦ繧ｯ繝ｪ繧｢
    For i = 0 To 10
        Set chk = pgCog.controls("chkBPSD" & CStr(i))
        chk.value = False
    Next i

    ' 2) 繧ｻ繝ｫ譁・ｭ怜・繧・| 縺ｧ蛻・ｧ｣縺励，aption 縺ｨ荳閾ｴ縺吶ｋ繝√ぉ繝・け繝懊ャ繧ｯ繧ｹ繧丹N
    s = CStr(ReadValueByCompatHeader(ws, r, "IO_Cog_BPSD"))
    If Len(s) > 0 Then
        arr = Split(s, "|")
        For i = LBound(arr) To UBound(arr)
            For j = 0 To 10
                Set chk = pgCog.controls("chkBPSD" & CStr(j))
                If chk.caption = arr(i) Then
                    chk.value = True
                    Exit For
                End If
            Next j
        Next i
    End If

    '=== 邊ｾ逾樣擇 combobox / note ===
    LoadComboValueByHeader ws, r, "IO_Mental_Mood", pgMental, "cmbMood"
    LoadComboValueByHeader ws, r, "IO_Mental_Motivation", pgMental, "cmbMotivation"
    LoadComboValueByHeader ws, r, "IO_Mental_Anxiety", pgMental, "cmbAnxiety"
    LoadComboValueByHeader ws, r, "IO_Mental_Relation", pgMental, "cmbRelation"
    LoadComboValueByHeader ws, r, "IO_Mental_Sleep", pgMental, "cmbSleep"

    v = ReadValueByCompatHeader(ws, r, "IO_Mental_Note")
    pgMental.controls("txtMentalNote").text = CStr(v)
End Sub


Private Sub LoadComboValueByHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal headerName As String, _
                                   ByVal parent As Object, ByVal comboName As String)
    Dim cmb As MSForms.ComboBox
    Dim v As String

    Set cmb = parent.controls(comboName)
    v = CStr(ReadValueByCompatHeader(ws, r, headerName))

    If Len(v) = 0 Then
        cmb.ListIndex = -1
    ElseIf ComboBoxHasValue(cmb, v) Then
        cmb.value = v
    Else
        cmb.ListIndex = -1
        Debug.Print "[Load_CognitionMental] skip invalid combo value"; _
                    " header=" & headerName & " combo=" & comboName & " value=" & v
    End If
End Sub

Private Function ReadValueByCompatHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal headerName As String) As Variant
    Dim headers As Variant
    Dim i As Long
    Dim col As Long
    Dim v As Variant

    headers = CompatHeaderNames(headerName)
    For i = LBound(headers) To UBound(headers)
        col = HeaderCol(CStr(headers(i)), ws)
        If col > 0 Then
            v = ws.Cells(r, col).value
            If Not IsNull(v) Then
                If Len(CStr(v)) > 0 Then
                    ReadValueByCompatHeader = v
                    Exit Function
                End If
            End If
        End If
    Next i

    ReadValueByCompatHeader = vbNullString
End Function

Private Function ComboBoxHasValue(ByVal cmb As MSForms.ComboBox, ByVal target As String) As Boolean
    Dim i As Long

    For i = 0 To cmb.ListCount - 1
        If StrComp(CStr(cmb.List(i)), target, vbBinaryCompare) = 0 Then
            ComboBoxHasValue = True
            Exit Function
        End If
    Next i
End Function



Private Function ComposeDailyLogBody(ByVal training As String, ByVal reaction As String, ByVal abnormal As String, ByVal plan As String) As String
    ComposeDailyLogBody = "縲仙ｮ滓命蜀・ｮｹ縲・ & vbCrLf & training & vbCrLf & vbCrLf & _
                          "縲仙茜逕ｨ閠・・蜿榊ｿ懊・ & vbCrLf & reaction & vbCrLf & vbCrLf & _
                          "縲千焚蟶ｸ謇隕九・ & vbCrLf & abnormal & vbCrLf & vbCrLf & _
                          "縲蝉ｻ雁ｾ後・譁ｹ驥昴・ & vbCrLf & plan
End Function

Private Sub FillDailyLogFieldsFromBody(ByVal body As String, ByRef training As String, ByRef reaction As String, ByRef abnormal As String, ByRef plan As String)
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long

    training = ""
    reaction = ""
    abnormal = ""
    plan = ""

    p1 = InStr(body, "縲仙ｮ滓命蜀・ｮｹ縲・)
    p2 = InStr(body, "縲仙茜逕ｨ閠・・蜿榊ｿ懊・)
    p3 = InStr(body, "縲千焚蟶ｸ謇隕九・)
    p4 = InStr(body, "縲蝉ｻ雁ｾ後・譁ｹ驥昴・)

    If p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 Then
        training = Trim$(Mid$(body, p1 + Len("縲仙ｮ滓命蜀・ｮｹ縲・), p2 - (p1 + Len("縲仙ｮ滓命蜀・ｮｹ縲・))))
        reaction = Trim$(Mid$(body, p2 + Len("縲仙茜逕ｨ閠・・蜿榊ｿ懊・), p3 - (p2 + Len("縲仙茜逕ｨ閠・・蜿榊ｿ懊・))))
        abnormal = Trim$(Mid$(body, p3 + Len("縲千焚蟶ｸ謇隕九・), p4 - (p3 + Len("縲千焚蟶ｸ謇隕九・))))
        plan = Trim$(Mid$(body, p4 + Len("縲蝉ｻ雁ｾ後・譁ｹ驥昴・)))
    Else
        training = body
    End If
End Sub


Public Function EnsureDailyLogFolderPath() As String

    Dim dataFolder As String
    Dim logsFolder As String

    dataFolder = ThisWorkbook.path & "\data"
    If Dir(dataFolder, vbDirectory) = "" Then MkDir dataFolder

    logsFolder = dataFolder & "\logs"
    If Dir(logsFolder, vbDirectory) = "" Then MkDir logsFolder

    EnsureDailyLogFolderPath = logsFolder & "\"

End Function


Public Function GetDailyLogFilePath(ByVal d As Date) As String

    Dim basePath As String
    basePath = EnsureDailyLogFolderPath()

    GetDailyLogFilePath = basePath & "DailyLog_" & Format$(d, "yyyy") & ".xlsx"

End Function

Private Function EnsureDailyLogSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("DailyLog")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = "DailyLog"
    End If

      If Trim$(CStr(ws.Cells(1, 1).value)) = "" Then ws.Cells(1, 1).value = "LogID"
      If Trim$(CStr(ws.Cells(1, 2).value)) = "" Then ws.Cells(1, 2).value = "蛻ｩ逕ｨ閠・D"
      If Trim$(CStr(ws.Cells(1, 3).value)) = "" Then ws.Cells(1, 3).value = "蛻ｩ逕ｨ閠・錐"
      If Trim$(CStr(ws.Cells(1, 4).value)) = "" Then ws.Cells(1, 4).value = "蛻ｩ逕ｨ譌･"
      If Trim$(CStr(ws.Cells(1, 5).value)) = "" Then ws.Cells(1, 5).value = "險倬鹸譛ｬ譁・
      If Trim$(CStr(ws.Cells(1, 6).value)) = "" Then ws.Cells(1, 6).value = "險倬鹸閠・
      If Trim$(CStr(ws.Cells(1, 7).value)) = "" Then ws.Cells(1, 7).value = "譖ｴ譁ｰ譌･譎・

    Set EnsureDailyLogSheet = ws
End Function

Private Function OpenDailyLogWorkbook(ByVal d As Date, ByVal createIfMissing As Boolean, ByRef openedHere As Boolean) As Workbook
    Dim filePath As String
    Dim wb As Workbook

    Call EnsureDailyLogFolderPath
    filePath = GetDailyLogFilePath(d)
    
    

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenDailyLogWorkbook = wb
            openedHere = False
            Exit Function
        End If
    Next wb

    If Dir(filePath) <> "" Then
        Set OpenDailyLogWorkbook = Application.Workbooks.Open(filePath)
        openedHere = True
        Exit Function
    End If
     

    If Not createIfMissing Then Exit Function
    
    
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Call EnsureDailyLogSheet(wb)
    wb.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    Set OpenDailyLogWorkbook = wb
    openedHere = True
End Function


Public Function GetDailyLogSheetByDate(ByVal d As Date, ByVal createIfMissing As Boolean, ByRef wb As Workbook, ByRef openedHere As Boolean) As Worksheet
    Set wb = OpenDailyLogWorkbook(d, createIfMissing, openedHere)
    If wb Is Nothing Then Exit Function
    Set GetDailyLogSheetByDate = EnsureDailyLogSheet(wb)
End Function

Private Function GenerateDailyLogID(ByVal ws As Worksheet, ByVal logDate As Date) As String
    Dim y As String
    Dim lastRow As Long
    Dim r As Long
    Dim maxSeq As Long
    Dim token As String
    Dim seqPart As String

    y = Format$(logDate, "yyyy")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
    token = Trim$(CStr(ws.Cells(r, 1).value))
    If Len(token) >= 12 And Left$(token, 4) = y And Mid$(token, 5, 1) = "-" Then
        seqPart = Mid$(token, 6)
        If IsNumeric(seqPart) Then
            If CLng(seqPart) > maxSeq Then maxSeq = CLng(seqPart)
        End If
    End If
Next r

GenerateDailyLogID = y & "-" & Format$(maxSeq + 1, "000000")
End Function

Public Sub Save_DailyLog_FromForm(owner As Object)
    ' 謇句虚菫晏ｭ俶凾縺ｯ SaveDailyLog_Append 繧剃ｽｿ逕ｨ
    mDailyLogManual = True
    SaveDailyLog_Append owner
    mDailyLogManual = False
End Sub
    

Private Sub SortDailyLogFilesByYearDesc(ByRef years() As Long, ByRef files() As String, ByVal itemCount As Long)
    Dim i As Long
    Dim j As Long
    Dim tmpYear As Long
    Dim tmpFile As String

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If years(j) > years(i) Then
                tmpYear = years(i)
                years(i) = years(j)
                years(j) = tmpYear

                tmpFile = files(i)
                files(i) = files(j)
                files(j) = tmpFile
            End If
        Next j
    Next i
End Sub
    




Public Sub Load_DailyLog_Latest_FromForm(owner As Object)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wbOpenedHere As Boolean
    Dim candidateWb As Workbook
    Dim candidateOpenedHere As Boolean
    Dim f As Object
    Dim txtName As Object
    Dim txtDate As Object
    Dim txtStaff As Object
    Dim txtTraining As Object
    Dim txtReaction As Object
    Dim txtAbnormal As Object
    Dim txtPlan As Object
    Dim hdr As Object
    Dim txtHdrPID As Object
    Dim lastRow As Long
    Dim r As Long
    Dim targetName As String
    Dim targetPid As String
    Dim hit As Boolean
    Dim body As String
    Dim basePath As String
    Dim fileName As String
    Dim token As String
    Dim itemCount As Long
    Dim years() As Long
    Dim files() As String
    Dim idx As Long
    Dim y As Long
    Dim awb As Workbook
    Dim filePath As String


    '--- 繝輔か繝ｼ繝荳翫・繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ蜿門ｾ・---
    Set txtName = SafeGetControl(owner, "txtName")
    Set f = ResolveDailyLogRoot(owner)
    If txtName Is Nothing Or f Is Nothing Then Exit Sub

    Set txtDate = ResolveDailyLogControl(owner, "txtDailyDate")
    Set txtStaff = ResolveDailyLogControl(owner, "txtDailyStaff")
    Set txtTraining = ResolveDailyLogControl(owner, "txtDailyTraining")
    Set txtReaction = ResolveDailyLogControl(owner, "txtDailyReaction")
    Set txtAbnormal = ResolveDailyLogControl(owner, "txtDailyAbnormal")
    Set txtPlan = ResolveDailyLogControl(owner, "txtDailyPlan")
    Set hdr = SafeGetControl(owner, "frHeader")
    Set txtHdrPID = SafeGetControl(hdr, "txtHdrPID")
    If txtDate Is Nothing Or txtStaff Is Nothing Or txtTraining Is Nothing Or txtReaction Is Nothing Or txtAbnormal Is Nothing Or txtPlan Is Nothing Or txtHdrPID Is Nothing Then Exit Sub
    



    '--- 隧ｲ蠖灘茜逕ｨ閠・・縲梧怙譁ｰ・医＞縺｡縺ｰ繧謎ｸ具ｼ峨阪・陦後ｒ謗｢縺・---
    targetName = Trim$(CStr(txtName.value))
    targetPid = Trim$(CStr(txtHdrPID.value))
    If targetPid = "" And targetName = "" Then GoTo FinallyExit

    basePath = EnsureDailyLogFolderPath()
    fileName = Dir(basePath & "DailyLog_*.xlsx")

    Do While fileName <> ""
        token = Mid$(fileName, Len("DailyLog_") + 1, 4)
        If Len(token) = 4 And IsNumeric(token) Then
            y = CLng(token)
            itemCount = itemCount + 1
            ReDim Preserve years(1 To itemCount)
            ReDim Preserve files(1 To itemCount)
            years(itemCount) = y
            files(itemCount) = fileName
        End If
        fileName = Dir()
    Loop

    If itemCount = 0 Then GoTo FinallyExit

    SortDailyLogFilesByYearDesc years, files, itemCount


    hit = False
    For idx = 1 To itemCount
        filePath = basePath & files(idx)
        Set candidateWb = Nothing
        candidateOpenedHere = False

        For Each awb In Application.Workbooks
            If StrComp(awb.FullName, filePath, vbTextCompare) = 0 Then
                Set candidateWb = awb
                Exit For
            End If
        Next awb

        If candidateWb Is Nothing Then
            Set candidateWb = Application.Workbooks.Open(filePath)
            candidateOpenedHere = True
        End If
        

        Set ws = EnsureDailyLogSheet(candidateWb)
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.rows.count, 3).End(xlUp).row
            For r = lastRow To 2 Step -1
                 If targetPid <> "" Then
                    If Trim$(CStr(ws.Cells(r, 2).value)) = targetPid Then
                        hit = True
                        Set wb = candidateWb
                        wbOpenedHere = candidateOpenedHere
                        Exit For
                    End If
                ElseIf Trim$(CStr(ws.Cells(r, 3).value)) = targetName Then
                    hit = True
                    Set wb = candidateWb
                    wbOpenedHere = candidateOpenedHere
                    Exit For
                End If
            Next r
        End If

        If hit Then Exit For

        If candidateOpenedHere Then
            candidateWb.Close SaveChanges:=False
        End If
    Next idx


    If Not hit Then GoTo FinallyExit
    
    

    '--- 隕九▽縺九▲縺溯｡後ｒ繝輔か繝ｼ繝縺ｸ蜿肴丐 ---
    body = CStr(ws.Cells(r, 5).value)

    txtDate.value = ws.Cells(r, 4).value
    txtStaff.value = ws.Cells(r, 6).value
    FillDailyLogFieldsFromBody body, txtTraining.value, txtReaction.value, txtAbnormal.value, txtPlan.value

FinallyExit:
    If wbOpenedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub




Public Sub SaveDailyLog_Append(owner As Object)

    
    ' 蟆ら畑繝懊ち繝ｳ縺九ｉ縺ｮ蜻ｼ縺ｳ蜃ｺ縺嶺ｻ･螟悶〒縺ｯ菴輔ｂ縺励↑縺・
    If Not mDailyLogManual Then Exit Sub


    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wbOpenedHere As Boolean
    Dim r As Long
    Dim f As Object
    Dim dt As Variant
    Dim nm As String
    Dim pid As String
    Dim staff As String
    Dim note As String
    Dim training As String
    Dim reaction As String
    Dim abnormal As String
    Dim plan As String
    Dim logDate As Date
    Dim lastRow As Long
    Dim hitRow As Long


    Set f = ResolveDailyLogRoot(owner)
    If f Is Nothing Then Exit Sub

    Dim txtDailyDate As Object
    Dim txtDailyStaff As Object
    Dim txtDailyTraining As Object
    Dim txtDailyReaction As Object
    Dim txtDailyAbnormal As Object
    Dim txtDailyPlan As Object
    Dim hdr As Object
    Dim txtHdrName As Object
    Dim txtHdrPID As Object
    
    

    Set txtDailyDate = ResolveDailyLogControl(owner, "txtDailyDate")
    Set txtDailyStaff = ResolveDailyLogControl(owner, "txtDailyStaff")
    Set txtDailyTraining = ResolveDailyLogControl(owner, "txtDailyTraining")
    Set txtDailyReaction = ResolveDailyLogControl(owner, "txtDailyReaction")
    Set txtDailyAbnormal = ResolveDailyLogControl(owner, "txtDailyAbnormal")
    Set txtDailyPlan = ResolveDailyLogControl(owner, "txtDailyPlan")
    Set hdr = SafeGetControl(owner, "frHeader")
    Set txtHdrName = SafeGetControl(hdr, "txtHdrName")
    If txtDailyDate Is Nothing Or txtDailyStaff Is Nothing Or txtDailyTraining Is Nothing Or txtDailyReaction Is Nothing Or txtDailyAbnormal Is Nothing Or txtDailyPlan Is Nothing Or txtHdrName Is Nothing Then Exit Sub
    
    Set txtHdrPID = SafeGetControl(hdr, "txtHdrPID")
    If txtDailyDate Is Nothing Or txtDailyStaff Is Nothing Or txtDailyTraining Is Nothing Or txtDailyReaction Is Nothing Or txtDailyAbnormal Is Nothing Or txtDailyPlan Is Nothing Or txtHdrName Is Nothing Or txtHdrPID Is Nothing Then Exit Sub

    dt = txtDailyDate.value
    nm = Trim$(txtHdrName.value)
    pid = Trim$(txtHdrPID.value)
    staff = Trim$(txtDailyStaff.value)
    training = CStr(txtDailyTraining.value)
    reaction = CStr(txtDailyReaction.value)
    abnormal = CStr(txtDailyAbnormal.value)
    plan = CStr(txtDailyPlan.value)
    note = ComposeDailyLogBody(training, reaction, abnormal, plan)


    '--- 蜈･蜉帙メ繧ｧ繝・け ---
    If nm = "" Then
     MsgBox "蛻ｩ逕ｨ閠・錐繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・, vbExclamation
     Exit Sub
    End If

    If Not IsDate(dt) Then
        MsgBox "險倬鹸譌･縺ｮ谺・↓豁｣縺励＞譌･莉倥ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・, vbExclamation
        Exit Sub
    End If

    If Trim$(training & reaction & abnormal & plan) = "" Then
        If MsgBox("險倬鹸蜀・ｮｹ縺檎ｩｺ縺ｧ縺吶′菫晏ｭ倥＠縺ｾ縺吶°・・, vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
    
 End If

    logDate = CDate(dt)
    Set ws = GetDailyLogSheetByDate(logDate, True, wb, wbOpenedHere)
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    hitRow = 0

    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 4).value) Then
            If CLng(CDate(ws.Cells(r, 4).value)) = CLng(logDate) Then
                If pid <> "" Then
                    If Trim$(CStr(ws.Cells(r, 2).value)) = pid Then
                        hitRow = r
                        Exit For
                    End If
                ElseIf Trim$(CStr(ws.Cells(r, 3).value)) = nm Then
                    hitRow = r
                    Exit For
                End If
            End If
        End If
    Next r

    If hitRow = 0 Then
        hitRow = IIf(lastRow < 2, 2, lastRow + 1)
        ws.Cells(hitRow, 1).value = GenerateDailyLogID(ws, logDate)
    ElseIf Trim$(CStr(ws.Cells(hitRow, 1).value)) = "" Then
        ws.Cells(hitRow, 1).value = GenerateDailyLogID(ws, logDate)
    
    End If
    
    

    '--- 霑ｽ險倩｡後ｒ豎ｺ繧√ｋ・・陦檎岼縺ｫ隕句・縺励′縺ゅｋ蜑肴署・・--
    ws.Cells(hitRow, 2).value = pid
    ws.Cells(hitRow, 3).value = nm
    ws.Cells(hitRow, 4).value = logDate
    ws.Cells(hitRow, 4).NumberFormatLocal = "yyyy/mm/dd"
    ws.Cells(hitRow, 5).value = note
    ws.Cells(hitRow, 6).value = staff
    ws.Cells(hitRow, 7).value = Now
    ws.Cells(hitRow, 7).NumberFormatLocal = "yyyy/mm/dd hh:mm"
    

End Sub




'=== Basic.* 縺ｨ譌･譛ｬ隱槭・繝・ム蛻励ｒ繝溘Λ繝ｼ縺吶ｋ豎守畑繝倥Ν繝代・ =====================

Private Sub MirrorBasicPair( _
        ByVal ws As Worksheet, ByVal rowNum As Long, _
        ByVal colBasic As Long, ByVal colJp As Long)

    Dim vBasic As Variant, vJp As Variant

    If colBasic <= 0 Or colJp <= 0 Then Exit Sub

    vBasic = ws.Cells(rowNum, colBasic).value
    vJp = ws.Cells(rowNum, colJp).value

    ' 縺ｩ縺｡繧峨°迚・婿縺縺大・縺｣縺ｦ縺・ｋ蝣ｴ蜷医√ｂ縺・援譁ｹ縺ｸ繧ｳ繝斐・
    If Len(vBasic) = 0 And Len(vJp) > 0 Then
        ws.Cells(rowNum, colBasic).value = vJp
    ElseIf Len(vJp) = 0 And Len(vBasic) > 0 Then
        ws.Cells(rowNum, colJp).value = vBasic
    End If
End Sub

'=== 蝓ｺ譛ｬ諠・ｱ縺ｮ譁ｰ譌ｧ蛻励ｒ繝溘Λ繝ｼ縺吶ｋ =====================================
'  繝ｻBasic.* 縺ｨ 譌･譛ｬ隱槭・繝・ム 縺ｮ荳｡譁ｹ繧偵檎ｩｺ縺・※縺・ｋ譁ｹ縺ｸ縲阪さ繝斐・縺吶ｋ
'  繝ｻ縺ｩ縺｡繧峨°迚・婿縺ｫ縺励°蛟､縺後↑縺代ｌ縺ｰ縲√◎縺ｮ蛟､繧偵ｂ縺・援譁ｹ縺ｸ蜀吶☆縺縺・
'  繝ｻ荳｡譁ｹ縺ｫ蛟､縺後≠繧句ｴ蜷医・菴輔ｂ縺励↑縺・ｼ郁｡晉ｪ∝屓驕ｿ・・
Public Sub MirrorBasicRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    ' ID
    MirrorBasicPair ws, rowNum, "Basic.ID", "ID"
    ' 豌丞錐
    MirrorBasicPair ws, rowNum, "Basic.Name", "豌丞錐"
    ' 隧穂ｾ｡譌･
    MirrorBasicPair ws, rowNum, "Basic.EvalDate", "隧穂ｾ｡譌･"
    ' 蟷ｴ鮨｢
    MirrorBasicPair ws, rowNum, "Basic.Age", "蟷ｴ鮨｢"
    ' 諤ｧ蛻･
    MirrorBasicPair ws, rowNum, "Basic.Sex", "諤ｧ蛻･"
    ' 隧穂ｾ｡閠・
    MirrorBasicPair ws, rowNum, "Basic.Evaluator", "隧穂ｾ｡閠・
    ' 隧穂ｾ｡閠・・遞ｮ
    MirrorBasicPair ws, rowNum, "Basic.EvaluatorJob", "隧穂ｾ｡閠・・遞ｮ"
    ' 逋ｺ逞・律
    MirrorBasicPair ws, rowNum, "Basic.OnsetDate", "逋ｺ逞・律"
    ' 謔｣閠・eeds
    MirrorBasicPair ws, rowNum, "Basic.Needs.Patient", "謔｣閠・eeds"
    ' 螳ｶ譌蒐eeds
    MirrorBasicPair ws, rowNum, "Basic.Needs.Family", "螳ｶ譌蒐eeds"
    ' 逕滓ｴｻ迥ｶ豕・
    MirrorBasicPair ws, rowNum, "BI.SocialParticipation", "逕滓ｴｻ迥ｶ豕・
    ' 荳ｻ險ｺ譁ｭ
    MirrorBasicPair ws, rowNum, "Basic.PrimaryDx", "荳ｻ險ｺ譁ｭ"
    ' 隕∽ｻ玖ｭｷ蠎ｦ
    MirrorBasicPair ws, rowNum, "Basic.CareLevel", "隕∽ｻ玖ｭｷ蠎ｦ"

    Exit Sub

ErrHandler:
    
End Sub
'=== 蝓ｺ譛ｬ諠・ｱ縺ｮ譁ｰ譌ｧ蛻励ｒ繝溘Λ繝ｼ縺吶ｋ =====================================
'  繝ｻBasic.* 縺ｨ 譌･譛ｬ隱槭・繝・ム 縺ｮ荳｡譁ｹ繧偵檎ｩｺ縺・※縺・ｋ譁ｹ縺ｸ縲阪さ繝斐・縺吶ｋ
'  繝ｻ縺ｩ縺｡繧峨°迚・婿縺ｫ縺励°蛟､縺後↑縺代ｌ縺ｰ縲√◎縺ｮ蛟､繧偵ｂ縺・援譁ｹ縺ｸ蜀吶☆縺縺・
'  繝ｻ荳｡譁ｹ縺ｫ蛟､縺後≠繧句ｴ蜷医・菴輔ｂ縺励↑縺・ｼ郁｡晉ｪ∝屓驕ｿ・・
Public Sub MirrorBasicRow_Eval(ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    Dim pairs As Variant
    Dim i As Long
    Dim headerNew As String, headerOld As String
    Dim cNew As Long, cOld As Long
    Dim vNew As Variant, vOld As Variant
    Dim sNew As String, sOld As String

    ' 蟇ｾ雎｡繝壹い荳隕ｧ・亥ｷｦ縺・Basic.*縲∝承縺梧律譛ｬ隱槭・繝・ム・・
    pairs = Array( _
        Array("Basic.ID", "ID"), _
        Array("Basic.Name", "豌丞錐"), _
        Array("Basic.EvalDate", "隧穂ｾ｡譌･"), _
        Array("Basic.Age", "蟷ｴ鮨｢"), _
        Array("Basic.Sex", "諤ｧ蛻･"), _
        Array("Basic.Evaluator", "隧穂ｾ｡閠・), _
        Array("隧穂ｾ｡閠・・遞ｮ", "txtEvaluatorJob"), _
        Array("Basic.OnsetDate", "逋ｺ逞・律"), _
        Array("Basic.Needs.Patient", "謔｣閠・eeds"), _
        Array("Basic.Needs.Family", "螳ｶ譌蒐eeds"), _
        Array("BI.SocialParticipation", "逕滓ｴｻ迥ｶ豕・), _
        Array("Basic.PrimaryDx", "荳ｻ險ｺ譁ｭ"), _
        Array("Basic.CareLevel", "隕∽ｻ玖ｭｷ蠎ｦ") _
    )

    For i = LBound(pairs) To UBound(pairs)
        headerNew = pairs(i)(0)
        headerOld = pairs(i)(1)

        ' 隕句・縺怜・繧貞叙蠕暦ｼ医←縺｡繧峨°辟｡縺代ｌ縺ｰ繧ｹ繧ｭ繝・・・・
        cNew = FindColByHeaderExact(ws, headerNew)
        cOld = FindColByHeaderExact(ws, headerOld)
        If cNew = 0 Or cOld = 0 Then GoTo NextPair

        vNew = ws.Cells(rowNum, cNew).value
        vOld = ws.Cells(rowNum, cOld).value

        sNew = Trim$(CStr(vNew))
        sOld = Trim$(CStr(vOld))

        ' 縺ｩ縺｡繧峨°縺縺大沂縺ｾ縺｣縺ｦ縺・ｋ蝣ｴ蜷医∫ｩｺ縺・※縺・ｋ譁ｹ縺ｸ繧ｳ繝斐・
        If sNew = "" And sOld <> "" Then
            ws.Cells(rowNum, cNew).value = vOld
        ElseIf sOld = "" And sNew <> "" Then
            ws.Cells(rowNum, cOld).value = vNew
        End If

NextPair:
    Next i

    Exit Sub

ErrHandler:
    Debug.Print "[WARN] MirrorBasicRow_Eval error at row=" & rowNum & _
                " : " & Err.Number & " " & Err.Description
End Sub






Sub Probe_ID_Candidates_Deep()
    Walk frmEval
End Sub

Private Sub Walk(ByVal parent As Object)
    Dim c As Object
    For Each c In parent.controls
        DumpIfIDLike c
        If HasControls(c) Then Walk c
    Next
End Sub

Private Function HasControls(ByVal o As Object) As Boolean
    On Error Resume Next
    Dim n As Long
    n = o.controls.count
    HasControls = (Err.Number = 0 And n > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub DumpIfIDLike(ByVal c As Object)
    Dim nm As String, tg As String, cap As String, val As String
    nm = LCase$(c.name)
    On Error Resume Next
    tg = LCase$(Trim$(c.tag))
    cap = LCase$(CStr(c.caption))
    val = CStr(c.value)
    On Error GoTo 0
    
    If InStr(nm, "id") > 0 Or InStr(tg, "id") > 0 Or InStr(cap, "id") > 0 Then
        Debug.Print TypeName(c) & "  name=" & c.name & "  tag=" & tg & "  caption=" & cap & "  value=" & val
    End If
End Sub





Private Function EvalIndexHeaders() As Variant
    EvalIndexHeaders = Array(HDR_USER_ID, HDR_NAME, HDR_KANA, HDR_SHEET, HDR_FIRST_EVAL, HDR_LATEST_EVAL, HDR_RECORD_COUNT)
End Function

Private Function EnsureEvalIndexSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EVAL_INDEX_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        ws.name = EVAL_INDEX_SHEET_NAME
    End If

    Dim headers As Variant: headers = EvalIndexHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = CStr(headers(i))
    Next i

    Set EnsureEvalIndexSheet = ws
End Function
Private Function BasicInfoLegacyHeaders() As Variant
    BasicInfoLegacyHeaders = Array( _
        "豌丞錐", "繝輔Μ繧ｬ繝・, "諤ｧ蛻･", "逕溷ｹｴ譛域律", "蟷ｴ鮨｢", "菴乗園", _
        "髮ｻ隧ｱ逡ｪ蜿ｷ", "譛ｬ莠ｺNeeds", "螳ｶ譌蒐eeds", "荳ｻ逞・錐", "隕∽ｻ玖ｭｷ蠎ｦ", _
        "逋ｺ逞・律", "譌｢蠕豁ｴ", "鬮倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", "隱咲衍逞・ｫ倬ｽ｢閠・・譌･蟶ｸ逕滓ｴｻ閾ｪ遶句ｺｦ", _
        "隧穂ｾ｡譌･", "蛻晏屓隧穂ｾ｡譌･", "邨碁℃", "蛯呵・, "隕∵髪謠ｴ", "陬懷勧蜈ｷ", "逕滓ｴｻ迥ｶ豕・)
End Function

Private Function CommonHistoryHeaders() As Variant

    Dim headers As Collection: Set headers = New Collection
    Dim v As Variant

    For Each v In Array( _
        HDR_ROWNO, "Basic.ID", "Basic.Name", "Basic.NameKana", _
        "Basic.EvalDate", "Basic.Evaluator", "Basic.EvaluatorJob", "Basic.Age", _
        "Basic.BirthDate", "Basic.Sex", "Basic.PrimaryDx", "Basic.OnsetDate", _
        "Basic.CareLevel", "Basic.DementiaADL", "BI.SocialParticipation", _
        "Basic.Needs.Patient", "Basic.Needs.Family", _
        "Basic.Medical.AdmitDate", "Basic.Medical.DischargeDate", _
        "Basic.Medical.CourseNote", "Basic.Medical.ComplicationNote", _
        HDR_HOMEENV_CHECKS, HDR_HOMEENV_NOTE, HDR_AIDS_CHECKS, HDR_RISK_CHECKS, _
         "IO_Cog_Memory", "IO_Cog_Attention", "IO_Cog_Orientation", _
        "IO_Cog_Judgment", "IO_Cog_Executive", "IO_Cog_Language", _
        "IO_Cog_DementiaType", "IO_Cog_DementiaNote", "IO_Cog_BPSD", _
        "IO_Mental_Mood", "IO_Mental_Motivation", "IO_Mental_Anxiety", _
        "IO_Mental_Relation", "IO_Mental_Sleep", "IO_Mental_Note")
        headers.Add CStr(v)
    Next v

    For Each v In BasicInfoLegacyHeaders()
        headers.Add CStr(v)
    Next v

    Dim arr() As String
    ReDim arr(0 To headers.count - 1) As String
    Dim i As Long
    For i = 1 To headers.count
        arr(i - 1) = CStr(headers(i))
    Next i
    CommonHistoryHeaders = arr
    
    End Function

Private Sub EnsureHistorySheetInitialized(ByVal ws As Worksheet)
    Dim headers As Variant: headers = CommonHistoryHeaders()
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If FindColByHeaderExact(ws, CStr(headers(i))) = 0 Then
            ws.Cells(1, ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1).value = CStr(headers(i))
        End If
    Next i
End Sub

Private Function NextHistorySheetName(ByVal indexWs As Worksheet) As String
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 4).End(xlUp).row
    Dim maxNo As Long, r As Long, nm As String, n As Long
    For r = 2 To lastRow
        nm = CStr(indexWs.Cells(r, 4).value)
        If Left$(nm, Len(EVAL_HISTORY_SHEET_PREFIX)) = EVAL_HISTORY_SHEET_PREFIX Then
            On Error Resume Next
            n = CLng(Mid$(nm, Len(EVAL_HISTORY_SHEET_PREFIX) + 1))
            On Error GoTo 0
            If n > maxNo Then maxNo = n
        End If
    Next r
    NextHistorySheetName = EVAL_HISTORY_SHEET_PREFIX & Format$(maxNo + 1, "0000")
End Function

Private Function FindEvalIndexRowsByName(ByVal indexWs As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 2).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(NormalizeName(CStr(indexWs.Cells(r, 2).value)), NormalizeName(nameText), vbTextCompare) = 0 Then c.Add r
    Next r
    Set FindEvalIndexRowsByName = c
End Function



Private Function FindEvalIndexRowByUserID(ByVal indexWs As Worksheet, ByVal userID As String) As Long
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(indexWs.Cells(r, 1).value)), Trim$(userID), vbTextCompare) = 0 Then
            FindEvalIndexRowByUserID = r
            Exit Function
        End If
    Next r
End Function


Private Function FindEvalIndexRowBySheetName(ByVal indexWs As Worksheet, ByVal sheetName As String) As Long
    Dim lastRow As Long: lastRow = indexWs.Cells(indexWs.rows.count, 4).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(indexWs.Cells(r, 4).value), sheetName, vbTextCompare) = 0 Then FindEvalIndexRowBySheetName = r: Exit Function
    Next r
End Function

Private Function FindEvalIndexRowsByNameWithoutUserID(ByVal indexWs As Worksheet, ByVal nameText As String) As Collection
    Dim c As New Collection
    Dim rowsByName As Collection
    Dim i As Long
    Dim rowNo As Long

    Set rowsByName = FindEvalIndexRowsByName(indexWs, nameText)
    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        If Len(Trim$(CStr(indexWs.Cells(rowNo, 1).value))) = 0 Then c.Add rowNo
    Next i

    Set FindEvalIndexRowsByNameWithoutUserID = c
End Function

Private Function TryGetWorksheetByName(ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    If Len(Trim$(sheetName)) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    TryGetWorksheetByName = Not ws Is Nothing
End Function

Private Function TryResolveExistingHistorySheetByName(ByVal storedSheetName As String, _
                                                      ByVal targetName As String, _
                                                      ByRef wsResolved As Worksheet) As Boolean
    Dim wsCandidate As Worksheet
    Dim ws As Worksheet
    Dim matchedCount As Long

    If TryGetWorksheetByName(storedSheetName, wsCandidate) Then
        If FindLatestRowByName(wsCandidate, targetName) > 0 Then
            Set wsResolved = wsCandidate
            TryResolveExistingHistorySheetByName = True
            Exit Function
        End If
    End If

    For Each ws In ThisWorkbook.Worksheets
        If (Left$(ws.name, Len(EVAL_HISTORY_SHEET_PREFIX)) = EVAL_HISTORY_SHEET_PREFIX) _
           Or StrComp(ws.name, EVAL_SHEET_NAME, vbTextCompare) = 0 Then
            If FindLatestRowByName(ws, targetName) > 0 Then
                matchedCount = matchedCount + 1
                Set wsResolved = ws
                If matchedCount > 1 Then
                    Set wsResolved = Nothing
                    Exit Function
                End If
            End If
        End If
    Next ws

    If matchedCount = 1 Then
        TryResolveExistingHistorySheetByName = True
    End If
End Function



Private Function BuildLegacyTransferCandidatesMessage(ByVal indexWs As Worksheet, ByVal rowsByName As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long
    Dim kanaVal As String
    Dim latestVal As String
    Dim recCount As String
    Dim sheetName As String

    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        kanaVal = Trim$(CStr(indexWs.Cells(rowNo, 3).value))
        sheetName = Trim$(CStr(indexWs.Cells(rowNo, 4).value))
        latestVal = Trim$(CStr(indexWs.Cells(rowNo, 6).value))
        recCount = Trim$(CStr(indexWs.Cells(rowNo, 7).value))

        lines = lines & CStr(i) & ") Sheet:" & sheetName
        If Len(kanaVal) > 0 Then lines = lines & " / Kana:" & kanaVal
        If Len(latestVal) > 0 Then lines = lines & " / Latest:" & latestVal
        If Len(recCount) > 0 Then lines = lines & " / Count:" & recCount
        If i < rowsByName.count Then lines = lines & vbCrLf
    Next i

    If Len(lines) = 0 Then Exit Function

    BuildLegacyTransferCandidatesMessage = _
        "蜷悟ｧ灘酔蜷阪・蛻ｩ逕ｨ閠・′隍・焚蟄伜惠縺励∪縺吶・ & vbCrLf & _
        "蟇ｾ雎｡閠・ｒ迚ｹ螳壹☆繧九◆繧√！D繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・ & vbCrLf & _
        "・医∪縺溘・蛟呵｣懊°繧蛾∈謚槭＠縺ｦ縺上□縺輔＞・・ & vbCrLf & vbCrLf & _
        lines
End Function

Private Function PickLegacyTransferIndexRow(ByVal indexWs As Worksheet, _
                                            ByVal rowsByName As Collection, _
                                            ByVal userID As String, _
                                            ByVal personName As String, _
                                            ByVal forSave As Boolean) As Long
    Dim prompt As String
    Dim picked As Variant
    Dim i As Long
    Dim n As Long
    Dim hasUnassigned As Boolean
    
    If rowsByName Is Nothing Then Exit Function
    If rowsByName.count = 0 Then Exit Function

    For i = 1 To rowsByName.count
        If Len(Trim$(CStr(indexWs.Cells(CLng(rowsByName(i)), 1).value))) = 0 Then
            hasUnassigned = True
            Exit For
        End If
    Next i

    prompt = BuildLegacyTransferCandidatesMessage(indexWs, rowsByName) & vbCrLf & vbCrLf
    If hasUnassigned Then
        prompt = prompt & "ID譛ｪ險ｭ螳壹・譌ｧ險倬鹸縺瑚ｦ九▽縺九ｊ縺ｾ縺励◆縲・ & vbCrLf
    Else
        prompt = prompt & "蜷悟ｧ灘酔蜷阪・蛻ｩ逕ｨ閠・′隍・焚蟄伜惠縺励∪縺吶・ & vbCrLf
    End If

    prompt = prompt & "蟇ｾ雎｡閠・ " & personName & " / ID: " & userID & vbCrLf & _
         "蠑輔″邯吶＄險倬鹸縺ｮ逡ｪ蜿ｷ繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・

    picked = Application.InputBox(prompt, "譌ｧ險倬鹸縺ｮ蠑輔″邯吶℃", Type:=1)
    If VarType(picked) = vbBoolean Then Exit Function
    If IsError(picked) Then Exit Function
    If Not IsNumeric(picked) Then Exit Function
    If Len(CStr(picked)) = 0 Then Exit Function

    n = CLng(picked)
    If n < 1 Or n > rowsByName.count Then Exit Function

    PickLegacyTransferIndexRow = CLng(rowsByName(n))
End Function

Private Function WriteUserIDToLegacyHistory(ByVal wsTarget As Worksheet, _
                                            ByVal userID As String, _
                                            ByVal personName As String) As Long
    Dim cID As Long
    Dim cName As Long
    Dim lastRow As Long
    Dim r As Long

    cID = FindColByHeaderExact(wsTarget, "Basic.ID")
    If cID = 0 Then cID = EnsureHeader(wsTarget, "Basic.ID")

    cName = FindColByHeaderExact(wsTarget, "Basic.Name")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "豌丞錐")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "蛻ｩ逕ｨ閠・錐")
       If cName = 0 Then cName = FindHeaderCol(wsTarget, "Name")

    lastRow = LastDataRow(wsTarget)
    For r = 2 To lastRow
        If Len(Trim$(CStr(wsTarget.Cells(r, cID).value))) > 0 Then GoTo NextRow
        If cName > 0 Then
            If NormalizeName(CStr(wsTarget.Cells(r, cName).value)) <> NormalizeName(personName) Then GoTo NextRow
        End If
        wsTarget.Cells(r, cID).value = userID
        WriteUserIDToLegacyHistory = WriteUserIDToLegacyHistory + 1
NextRow:
    Next r
End Function

Private Function AssignUserIDToHistoryEntry(ByVal indexWs As Worksheet, _
                                            ByVal indexRow As Long, _
                                            ByVal userID As String, _
                                            ByVal personName As String, _
                                            ByVal kanaVal As String, _
                                            ByRef wsTarget As Worksheet) As Boolean
    Dim sheetName As String

    If indexRow <= 0 Then Exit Function

    If Len(CStr(indexWs.Cells(indexRow, 4).value)) = 0 Then indexWs.Cells(indexRow, 4).value = NextHistorySheetName(indexWs)
    sheetName = CStr(indexWs.Cells(indexRow, 4).value)

    indexWs.Cells(indexRow, 1).value = userID
    indexWs.Cells(indexRow, 2).value = personName
    If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal

    Set wsTarget = EnsureEvalSheet(sheetName)
    EnsureHistorySheetInitialized wsTarget
    Call WriteUserIDToLegacyHistory(wsTarget, userID, personName)

    AssignUserIDToHistoryEntry = True
End Function



Private Function ResolveUserHistorySheet(owner As Object, ByVal forSave As Boolean, ByRef wsTarget As Worksheet, ByRef message As String) As Boolean
    Dim nm As String: nm = Trim$(owner.txtName.text)
    If Len(nm) = 0 Then message = "豌丞錐縺梧悴蜈･蜉帙〒縺・: Exit Function

    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    Dim rowsByName As Collection
    Dim rowsByID As Collection
    Dim rowsByNameWithoutID As Collection
    Dim indexRow As Long
    Dim newRow As Long
    Dim pickedRow As Long

    Set rowsByName = FindEvalIndexRowsByName(indexWs, nm)
    HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                           "forSave=" & CStr(forSave), _
                           "name=" & HistoryLoadDebug_Quote(nm), _
                           "id=" & HistoryLoadDebug_Quote(idVal), _
                           "kana=" & HistoryLoadDebug_Quote(kanaVal), _
                           "rowsByName=" & CStr(rowsByName.count)
    
    
    If Len(idVal) = 0 And rowsByName.count = 1 Then
        Dim storedSheetName As String

        indexRow = CLng(rowsByName(1))
        storedSheetName = Trim$(CStr(indexWs.Cells(indexRow, 4).value))
        HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                               "branch=noID_uniqueName", _
                               "indexRow=" & CStr(indexRow), _
                               "indexSheetCellBefore=" & HistoryLoadDebug_Quote(storedSheetName)

        If forSave Then
            If Len(storedSheetName) = 0 Then
                storedSheetName = NextHistorySheetName(indexWs)
                indexWs.Cells(indexRow, 4).value = storedSheetName
            End If
            Set wsTarget = EnsureEvalSheet(storedSheetName)
            EnsureHistorySheetInitialized wsTarget
        Else
            If TryResolveExistingHistorySheetByName(storedSheetName, nm, wsTarget) Then
                If StrComp(storedSheetName, wsTarget.name, vbTextCompare) <> 0 Then
                    indexWs.Cells(indexRow, 4).value = wsTarget.name
                End If
            ElseIf TryGetWorksheetByName(storedSheetName, wsTarget) Then
                EnsureHistorySheetInitialized wsTarget
            Else
                message = "蟇ｾ雎｡縺ｮ隧穂ｾ｡螻･豁ｴ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・
                Exit Function
            End If
        End If

 
       
        HistoryLoadDebug_Print "[ResolveUserHistorySheet]", _
                               "branch=noID_uniqueName", _
                               "resolvedSheet=" & HistoryLoadDebug_SheetName(wsTarget), _
                               "sheetLastDataRow=" & CStr(LastDataRow(wsTarget))


        If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
        ResolveUserHistorySheet = True
        Exit Function
    End If

    
    If Len(idVal) > 0 Then
        Set rowsByID = FindEvalIndexRowsByUserID(indexWs, idVal)
        If rowsByID.count > 1 Then
            message = BuildDuplicateUserIDMessage(indexWs, idVal, rowsByID)
            Exit Function
        End If

        If rowsByID.count = 1 Then
            indexRow = CLng(rowsByID(1))

            Dim indexName As String: indexName = Trim$(CStr(indexWs.Cells(indexRow, 2).value))
            Dim indexKana As String: indexKana = Trim$(CStr(indexWs.Cells(indexRow, 3).value))
            If NormalizeName(indexName) <> NormalizeName(nm) _
               Or Not IsSameKanaIfAvailable(kanaVal, indexKana) Then
                message = BuildUserIdentityMismatchMessage(idVal, nm, indexName, kanaVal, indexKana)
                Exit Function
            End If

            If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
            If Len(CStr(indexWs.Cells(indexRow, 4).value)) = 0 Then indexWs.Cells(indexRow, 4).value = NextHistorySheetName(indexWs)

            Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(indexRow, 4).value))

            EnsureHistorySheetInitialized wsTarget
            ResolveUserHistorySheet = True
            Exit Function
        End If
        
    End If

    
    If rowsByName.count = 0 Then
        
        If Not forSave Then
           message = "蟇ｾ雎｡縺ｮ隧穂ｾ｡螻･豁ｴ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・
            Exit Function
        End If
        
        newRow = NextAppendRow(indexWs)
        indexWs.Cells(newRow, 1).value = idVal
        indexWs.Cells(newRow, 2).value = nm
        indexWs.Cells(newRow, 3).value = kanaVal
        indexWs.Cells(newRow, 4).value = NextHistorySheetName(indexWs)
        Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(newRow, 4).value))
        
        ResolveUserHistorySheet = True
        Exit Function
    
    End If

    If Len(idVal) > 0 Then
        Set rowsByNameWithoutID = FindEvalIndexRowsByNameWithoutUserID(indexWs, nm)
        pickedRow = PickLegacyTransferIndexRow(indexWs, rowsByNameWithoutID, idVal, nm, forSave)
        If pickedRow > 0 Then
            If AssignUserIDToHistoryEntry(indexWs, pickedRow, idVal, nm, kanaVal, wsTarget) Then
                ResolveUserHistorySheet = True
                Exit Function
            End If
        End If

        If forSave Then
            newRow = NextAppendRow(indexWs)
            indexWs.Cells(newRow, 1).value = idVal
            indexWs.Cells(newRow, 2).value = nm
            indexWs.Cells(newRow, 3).value = kanaVal
            indexWs.Cells(newRow, 4).value = NextHistorySheetName(indexWs)
            Set wsTarget = EnsureEvalSheet(CStr(indexWs.Cells(newRow, 4).value))
            EnsureHistorySheetInitialized wsTarget
            ResolveUserHistorySheet = True
            Exit Function
        End If

        message = "蜷悟ｧ灘酔蜷阪・蛻ｩ逕ｨ閠・′隍・焚蟄伜惠縺励∪縺吶・ & vbCrLf & _
          "隧ｲ蠖薙☆繧句ｱ･豁ｴ繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・
        If Not rowsByNameWithoutID Is Nothing Then
            If rowsByNameWithoutID.count > 0 Then
                message = message & vbCrLf & vbCrLf & BuildLegacyTransferCandidatesMessage(indexWs, rowsByNameWithoutID)
            End If
        End If
        Exit Function
    End If
    
    If Not forSave Then
        pickedRow = PickDuplicateNameIndexRow(indexWs, rowsByName, nm)
        If pickedRow > 0 Then
            If TryResolveHistorySheetFromIndexRow(indexWs, pickedRow, nm, wsTarget) Then
                ResolveUserHistorySheet = True
                Exit Function
            End If
            message = "蛟呵｣懊′縺ゅｊ縺ｾ縺帙ｓ縲・
            Exit Function
        End If
        Exit Function
    End If

    message = "蜷悟ｧ灘酔蜷阪・蛻ｩ逕ｨ閠・′隍・焚縺・ｋ縺溘ａ縲！D縺ｾ縺溘・螻･豁ｴ繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・ & _
          BuildDuplicateNameCandidatesMessage(indexWs, rowsByName)
End Function



Public Function TryGetUserHistorySheet(ByVal owner As Object, ByRef wsTarget As Worksheet) As Boolean
    Dim message As String
    TryGetUserHistorySheet = ResolveUserHistorySheet(owner, False, wsTarget, message)
End Function


Private Function TryResolveHistorySheetFromIndexRow(ByVal indexWs As Worksheet, _
                                                    ByVal indexRow As Long, _
                                                    ByVal targetName As String, _
                                                    ByRef wsTarget As Worksheet) As Boolean
    Dim storedSheetName As String

    If indexRow <= 0 Then Exit Function

    storedSheetName = Trim$(CStr(indexWs.Cells(indexRow, 4).value))
    If Len(storedSheetName) = 0 Then Exit Function

    If TryGetWorksheetByName(storedSheetName, wsTarget) Then
        EnsureHistorySheetInitialized wsTarget
        TryResolveHistorySheetFromIndexRow = True
        Exit Function
    End If

    If TryResolveExistingHistorySheetByName(storedSheetName, targetName, wsTarget) Then
        If StrComp(storedSheetName, wsTarget.name, vbTextCompare) <> 0 Then
            indexWs.Cells(indexRow, 4).value = wsTarget.name
        End If
        EnsureHistorySheetInitialized wsTarget
        TryResolveHistorySheetFromIndexRow = True
    End If
End Function

Private Function BuildDuplicateNameCandidateLine(ByVal indexWs As Worksheet, _
                                                 ByVal rowNo As Long, _
                                                 ByVal itemNo As Long) As String
    Dim idVal As String
    Dim kanaVal As String
    Dim latestVal As String
    Dim recCount As String
    Dim sheetName As String

    idVal = Trim$(CStr(indexWs.Cells(rowNo, 1).value))
    kanaVal = Trim$(CStr(indexWs.Cells(rowNo, 3).value))
    sheetName = Trim$(CStr(indexWs.Cells(rowNo, 4).value))
    latestVal = Trim$(CStr(indexWs.Cells(rowNo, 6).value))
    recCount = Trim$(CStr(indexWs.Cells(rowNo, 7).value))

    BuildDuplicateNameCandidateLine = CStr(itemNo) & ") Sheet:" & sheetName
    If Len(kanaVal) > 0 Then BuildDuplicateNameCandidateLine = BuildDuplicateNameCandidateLine & " / Kana:" & kanaVal
    If Len(latestVal) > 0 Then BuildDuplicateNameCandidateLine = BuildDuplicateNameCandidateLine & " / Latest:" & latestVal
    If Len(recCount) > 0 Then BuildDuplicateNameCandidateLine = BuildDuplicateNameCandidateLine & " / Count:" & recCount
    If Len(idVal) > 0 Then
        BuildDuplicateNameCandidateLine = BuildDuplicateNameCandidateLine & " / ID:" & idVal
    Else
        BuildDuplicateNameCandidateLine = BuildDuplicateNameCandidateLine & " / ID:(none)"
    End If
End Function

Private Function BuildDuplicateNameSelectionMessage(ByVal indexWs As Worksheet, _
                                                    ByVal rowsByName As Collection, _
                                                    ByVal personName As String) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long

    If rowsByName Is Nothing Then Exit Function
    If rowsByName.count = 0 Then Exit Function

    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        lines = lines & BuildDuplicateNameCandidateLine(indexWs, rowNo, i)
        If i < rowsByName.count Then lines = lines & vbCrLf
    Next i

    
    BuildDuplicateNameSelectionMessage = _
        "蜷悟ｧ灘酔蜷阪・蛻ｩ逕ｨ閠・′隍・焚蟄伜惠縺励∪縺吶・ & vbCrLf & _
        "蟇ｾ雎｡縺ｮ螻･豁ｴ繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・ & vbCrLf & _
        "蛻ｩ逕ｨ閠・錐: " & personName & vbCrLf & vbCrLf & _
        lines & vbCrLf & vbCrLf & _
        "逡ｪ蜿ｷ繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・
End Function

Private Function PickDuplicateNameIndexRow(ByVal indexWs As Worksheet, _
                                           ByVal rowsByName As Collection, _
                                           ByVal personName As String) As Long
    Dim prompt As String
    Dim picked As Variant
    Dim n As Long

    If rowsByName Is Nothing Then Exit Function
    If rowsByName.count = 0 Then Exit Function

    prompt = BuildDuplicateNameSelectionMessage(indexWs, rowsByName, personName)
    picked = Application.InputBox(prompt, "蜷悟ｧ灘酔蜷阪・蛟呵｣憺∈謚・, Type:=1)
    If VarType(picked) = vbBoolean Then Exit Function
    If IsError(picked) Then Exit Function
    If Not IsNumeric(picked) Then Exit Function
    If Len(CStr(picked)) = 0 Then Exit Function

    n = CLng(picked)
    If n < 1 Or n > rowsByName.count Then Exit Function

    PickDuplicateNameIndexRow = CLng(rowsByName(n))
End Function


Private Function BuildDuplicateNameCandidatesMessage(ByVal indexWs As Worksheet, ByVal rowsByName As Collection) As String
    Dim lines As String
    Dim i As Long
    Dim rowNo As Long
    Dim idVal As String
    Dim kanaVal As String
    Dim latestVal As String

    For i = 1 To rowsByName.count
        rowNo = CLng(rowsByName(i))
        idVal = Trim$(CStr(indexWs.Cells(rowNo, 1).value))
        kanaVal = Trim$(CStr(indexWs.Cells(rowNo, 3).value))
        latestVal = Trim$(CStr(indexWs.Cells(rowNo, 6).value))
        lines = lines & "- ID: " & idVal & " / 縺九↑: " & kanaVal & " / 譛譁ｰ: " & latestVal
        If i < rowsByName.count Then lines = lines & vbCrLf
    Next i

    If Len(lines) > 0 Then
                BuildDuplicateNameCandidatesMessage = vbCrLf & vbCrLf & ":" & vbCrLf & lines & _
            vbCrLf & vbCrLf & "譁ｰ隕上・蝣ｴ蜷医・谺｡縺ｮID繧剃ｽｿ逕ｨ縺ｧ縺阪∪縺・" & vbCrLf & _
            BuildNextAvailableUserIDCandidate(indexWs)
    End If
End Function

Private Function BuildNextAvailableUserIDCandidate(ByVal indexWs As Worksheet) As String
    Dim lastRow As Long
    Dim r As Long
    Dim rawID As String
    Dim idNum As Long
    Dim maxID As Long
    Dim maxDigits As Long
    Dim hasNumericID As Boolean

    lastRow = indexWs.Cells(indexWs.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        rawID = Trim$(CStr(indexWs.Cells(r, 1).value))
        If TryParseNumericUserID(rawID, idNum) Then
            hasNumericID = True
            If idNum > maxID Then maxID = idNum
            If Len(rawID) > maxDigits Then maxDigits = Len(rawID)
        End If
    Next r

    If Not hasNumericID Then
        BuildNextAvailableUserIDCandidate = "001"
        Exit Function
    End If

    If maxDigits < 3 Then maxDigits = 3
    BuildNextAvailableUserIDCandidate = Format$(maxID + 1, String$(maxDigits, "0"))
End Function

Private Function TryParseNumericUserID(ByVal rawID As String, ByRef parsedID As Long) As Boolean
    Dim numericValue As Double

    If Len(rawID) = 0 Then Exit Function
    If Not IsNumeric(rawID) Then Exit Function

    On Error GoTo EH
    numericValue = CDbl(rawID)
    If numericValue < 0 Then Exit Function
    If numericValue <> Fix(numericValue) Then Exit Function
    If numericValue > CLng(&H7FFFFFFF) Then Exit Function

    parsedID = CLng(numericValue)
    TryParseNumericUserID = True
    Exit Function
EH:
    TryParseNumericUserID = False
End Function

Private Sub HistoryLoadDebug_Print(ParamArray args())
    If Not HISTORY_LOAD_DEBUG Then Exit Sub

    Dim i As Long
    Dim msg As String
    For i = LBound(args) To UBound(args)
        msg = msg & IIf(i > 0, " | ", "") & CStr(args(i))
    Next i
    Debug.Print Format$(Now, "hh:nn:ss"), msg
End Sub

Private Function HistoryLoadDebug_Quote(ByVal valueText As String) As String
    HistoryLoadDebug_Quote = Chr$(34) & Replace$(CStr(valueText), Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function HistoryLoadDebug_SheetName(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        HistoryLoadDebug_SheetName = "(Nothing)"
    Else
        HistoryLoadDebug_SheetName = ws.name
    End If
End Function

Private Sub HistoryLoadDebug_ScanWorkbookForName(ByVal targetName As String, ByVal resolvedSheet As Worksheet)
    Dim ws As Worksheet
    Dim c As Long
    Dim lastRow As Long
    Dim r As Long
    Dim rowName As String
    Dim normalizedTarget As String
    Dim matched As Boolean

    normalizedTarget = NormalizeName(targetName)
    HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                           "targetName=" & HistoryLoadDebug_Quote(targetName), _
                           "resolvedSheet=" & HistoryLoadDebug_SheetName(resolvedSheet)

    For Each ws In ThisWorkbook.Worksheets
        If (Left$(ws.name, Len(EVAL_HISTORY_SHEET_PREFIX)) = EVAL_HISTORY_SHEET_PREFIX) _
           Or StrComp(ws.name, EVAL_SHEET_NAME, vbTextCompare) = 0 Then
            c = FindColByHeaderExact(ws, "Basic.Name")
                 If c = 0 Then c = FindHeaderCol(ws, "豌丞錐")
                 If c = 0 Then c = FindHeaderCol(ws, "蛻ｩ逕ｨ閠・錐")
                 If c = 0 Then c = FindHeaderCol(ws, "Name")

            If c = 0 Then
                HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                       "sheet=" & ws.name, _
                                       "nameHeaderMissing=True"
            Else
                lastRow = LastDataRow(ws)
                matched = False
                For r = 2 To lastRow
                    rowName = CStr(ws.Cells(r, c).value)
                    If NormalizeName(rowName) = normalizedTarget Then
                        HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                               "sheet=" & ws.name, _
                                               "row=" & CStr(r), _
                                               "raw=" & HistoryLoadDebug_Quote(rowName), _
                                               "isResolvedSheet=" & CStr(StrComp(ws.name, HistoryLoadDebug_SheetName(resolvedSheet), vbTextCompare) = 0)
                        matched = True
                    End If
                Next r
                If Not matched Then
                    HistoryLoadDebug_Print "[HistoryScanWorkbook]", _
                                           "sheet=" & ws.name, _
                                           "matchedRow=0", _
                                           "nameCol=" & CStr(c), _
                                           "lastRow=" & CStr(lastRow)
                End If
            End If
        End If
    Next ws
End Sub


Private Sub UpdateEvalIndexMetadata(ByVal owner As Object, ByVal indexRow As Long, ByVal sheetName As String)
    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim kanaVal As String: kanaVal = Trim$(GetHdrKanaText(owner))
    If Len(Trim$(owner.txtName.text)) > 0 Then indexWs.Cells(indexRow, 2).value = Trim$(owner.txtName.text)
    If Len(idVal) > 0 And Len(CStr(indexWs.Cells(indexRow, 1).value)) = 0 Then indexWs.Cells(indexRow, 1).value = idVal
    If Len(kanaVal) > 0 Then indexWs.Cells(indexRow, 3).value = kanaVal
    indexWs.Cells(indexRow, 4).value = sheetName
End Sub

Private Function TryParseEvalDate(ByVal v As Variant, ByRef normalizedDate As Date) As Boolean
    On Error GoTo EH
    If IsDate(v) Then normalizedDate = DateValue(CDate(v)): TryParseEvalDate = True
    Exit Function
EH:
    TryParseEvalDate = False
End Function

Private Function GetLatestValidEvalRow(ByVal ws As Worksheet) As Long
    Dim cEval As Long: cEval = FindColByHeaderExact(ws, "Basic.EvalDate")
    If cEval = 0 Then Exit Function
    Dim lastRow As Long: lastRow = LastDataRow(ws)
    Dim r As Long, d As Date
    For r = lastRow To 2 Step -1
        If TryParseEvalDate(ws.Cells(r, cEval).value, d) Then GetLatestValidEvalRow = r: Exit Function
    Next r
End Function

Public Sub GetUserEvalDateStats(ByVal wsTarget As Worksheet, _
                                ByRef firstEvalDate As String, _
                                ByRef latestEvalDate As String, _
                                ByRef previousEvalDate As String, _
                                ByRef recordCount As Long)
    Dim cEval As Long: cEval = FindColByHeaderExact(wsTarget, "Basic.EvalDate")
    If cEval = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = LastDataRow(wsTarget)
    Dim r As Long, d As Date
    Dim firstD As Date, latestD As Date, prevD As Date
    Dim hasFirst As Boolean, hasLatest As Boolean, hasPrev As Boolean

    For r = 2 To lastRow
        If TryParseEvalDate(wsTarget.Cells(r, cEval).value, d) Then
            recordCount = recordCount + 1

            If (Not hasFirst) Or d < firstD Then
                firstD = d
                hasFirst = True
            End If

            If (Not hasLatest) Or d > latestD Then
                If hasLatest And d <> latestD Then
                    If (Not hasPrev) Or latestD > prevD Then
                        prevD = latestD
                        hasPrev = True
                    End If
                End If
                latestD = d
                hasLatest = True
            ElseIf d < latestD Then
                If (Not hasPrev) Or d > prevD Then
                    prevD = d
                    hasPrev = True
                End If
            End If
        End If
    Next r

    If hasFirst Then firstEvalDate = Format$(firstD, "yyyy/mm/dd")
    If hasLatest Then latestEvalDate = Format$(latestD, "yyyy/mm/dd")
    If hasPrev Then previousEvalDate = Format$(prevD, "yyyy/mm/dd")
End Sub


Public Function GetPreviousEvalDateText(ByVal wsTarget As Worksheet) As String
    Dim firstEvalDate As String, latestEvalDate As String, previousEvalDate As String
    Dim recordCount As Long
    GetUserEvalDateStats wsTarget, firstEvalDate, latestEvalDate, previousEvalDate, recordCount
    GetPreviousEvalDateText = previousEvalDate
End Function


Public Function GetClientMasterCreatedDateText(ByVal owner As Object) As String
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = EnsureClientMasterSheet()
    Dim idVal As String: idVal = Trim$(GetID_FromBasicInfo(owner))
    Dim nameVal As String: nameVal = Trim$(GetCtlTextGeneric(owner, "txtName"))
    Dim shouldSkip As Boolean
    Dim rowNo As Long

    rowNo = FindClientMasterRow(ws, idVal, nameVal, shouldSkip)
    If rowNo <= 0 Then Exit Function

    GetClientMasterCreatedDateText = Trim$(CStr(ws.Cells(rowNo, 7).value))
    Exit Function
EH:
    Err.Clear
End Function


Private Sub UpdateEvalIndexStats(ByVal indexRow As Long, ByVal wsTarget As Worksheet)
    Dim firstEvalDate As String, latestEvalDate As String, previousEvalDate As String
    Dim recordCount As Long

    GetUserEvalDateStats wsTarget, firstEvalDate, latestEvalDate, previousEvalDate, recordCount

    Dim indexWs As Worksheet: Set indexWs = EnsureEvalIndexSheet()
    
    indexWs.Cells(indexRow, 5).value = firstEvalDate
    indexWs.Cells(indexRow, 6).value = latestEvalDate
    indexWs.Cells(indexRow, 7).value = recordCount
End Sub
