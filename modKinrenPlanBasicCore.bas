Attribute VB_Name = "modKinrenPlanBasicCore"
Public Function BuildBasicPlanStructure(ByVal mainCause As String, _
                                        ByVal needSelf As String, _
                                        ByVal needFamily As String, _
                                        ByVal needByDifficulty As String, _
                                        ByVal mmtMap As Object) As Object

                                        
    Dim result As Object
    Dim reason As String
    Dim shortCore As String
    Dim mmtTargetMuscle As String
    Dim fxCore As String

    Set result = CreateObject("Scripting.Dictionary")
    result("Activity_Long") = PickActivityLong(needSelf, needFamily, needByDifficulty)
    
    Set mmtMap = FilterMMTMap(mmtMap, result("Activity_Long"))
    
    Select Case result("Activity_Long")
          Case "螻句・豁ｩ陦・
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "閧｡螟冶ｻ｢,閭悟ｱ・閹昜ｼｸ螻・)
          Case "繝医う繝ｬ蜍穂ｽ・
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "閧｡螟冶ｻ｢,閹昜ｼｸ螻・)
          Case "螻句､匁ｭｩ陦・
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "閭悟ｱ・閧｡螟冶ｻ｢,閹昜ｼｸ螻・)
          Case "遶九■荳翫′繧・
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "閹昜ｼｸ螻・閧｡螟冶ｻ｢,閧｡莨ｸ螻・)
          Case Else
              mmtTargetMuscle = PickMMTTarget(mmtMap)
    End Select

    result("MMT_TargetMuscle") = mmtTargetMuscle
    result("MMT_MinScore") = mmtMap(mmtTargetMuscle)
    result("MainCause") = mainCause

    
        If Len(Trim$(needSelf)) > 0 Then
          reason = "譛ｬ莠ｺ蟶梧悍"
        ElseIf Len(Trim$(needFamily)) > 0 Then
          reason = "螳ｶ譌丞ｸ梧悍"
        Else
          reason = "蝗ｰ髮｣蠎ｦ荳贋ｽ・
        End If

    result("Activity_Reason") = reason

    Select Case mainCause
        Case "鮗ｻ逞ｺ"
        
        If result("MMT_MinScore") <= 2 Then
            fxCore = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励↓繧医ｊ"
        Else
            fxCore = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・↓繧医ｊ"
        End If

    Select Case result("Activity_Long")
    
        Case "螻句・豁ｩ陦・
            result("Function_Long") = fxCore & "遶玖・譛溷ｮ牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
        
        Case "遶九■荳翫′繧・
            result("Function_Long") = fxCore & "遶九■荳翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
        
        Case "繝医う繝ｬ", "繝医う繝ｬ蜍穂ｽ・
            result("Function_Long") = fxCore & "萓ｿ蠎ｧ遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
        
        Case "螻句､匁ｭｩ陦・
            result("Function_Long") = fxCore & "谿ｵ蟾ｮ譏・剄譎ゅ・螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
        
        Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
            result("Function_Long") = fxCore & "豬ｴ讒ｽ縺ｾ縺溘℃蜍穂ｽ懈凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "遘ｻ荵・
            result("Function_Long") = fxCore & "遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "襍ｷ螻・ｸ騾｣蜍穂ｽ・
            result("Function_Long") = fxCore & "襍ｷ縺堺ｸ翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
           
        Case Else
            result("Function_Long") = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ蝗ｳ繧九・
            
    End Select
   
    Case "逍ｼ逞・

    Select Case result("Activity_Long")

        Case "螻句・豁ｩ陦・
            result("Function_Long") = "豁ｩ陦梧凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case "螻句､匁ｭｩ陦・
            result("Function_Long") = "螻句､匁ｭｩ陦梧凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case "繝医う繝ｬ蜍穂ｽ・
            result("Function_Long") = "遶九■荳翫′繧頑凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
            result("Function_Long") = "蜈･豬ｴ蜍穂ｽ懈凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case "遘ｻ荵・
            result("Function_Long") = "遘ｻ荵怜虚菴懈凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case "襍ｷ螻・ｸ騾｣蜍穂ｽ・
            result("Function_Long") = "襍ｷ螻・虚菴懈凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・

        Case Else
            result("Function_Long") = "逍ｼ逞帙・霆ｽ貂帙ｒ蝗ｳ繧九・

    End Select
       
    Case "蝗ｰ髮｣蠎ｦ"

    Select Case result("Activity_Long")

        Case "螻句・豁ｩ陦・
            result("Function_Long") = "譁ｹ蜷題ｻ｢謠帶凾縺ｮ螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "螻句､匁ｭｩ陦・
            result("Function_Long") = "谿ｵ蟾ｮ譏・剄譎ゅ・螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "繝医う繝ｬ蜍穂ｽ・
            result("Function_Long") = "譁ｹ蜷題ｻ｢謠帛虚菴懊・螳牙ｮ壼喧繧貞峙繧九・

        Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
            result("Function_Long") = "豬ｴ螳､蜀・婿蜷題ｻ｢謠帙・螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "遘ｻ荵・
            result("Function_Long") = "蛛ｴ譁ｹ遘ｻ蜍墓凾縺ｮ螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・

        Case "襍ｷ螻・ｸ騾｣蜍穂ｽ・
            result("Function_Long") = "襍ｷ縺堺ｸ翫′繧雁虚菴懊・螳牙ｮ壼喧繧貞峙繧九・

        Case Else
            result("Function_Long") = "荳玖い讖溯・縺ｮ蜈ｨ菴鍋噪蜷台ｸ翫ｒ蝗ｳ繧九・

    End Select
    
    Case Else
        result("Function_Long") = ""
    End Select

    Select Case mainCause
      Case "鮗ｻ逞ｺ"
       If result("MMT_MinScore") <= 2 Then
    result("Function_Short") = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ蝗ｳ繧九・
Else
    result("Function_Short") = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ蝗ｳ繧九・
End If

      Case "逍ｼ逞・
        result("Function_Short") = "逍ｼ逞幄ｪ倡匱蜍穂ｽ懊・霆ｽ貂帙♀繧医・雋闕ｷ隱ｿ謨ｴ繧貞峙繧九・
      Case "蝗ｰ髮｣蠎ｦ"
        result("Function_Short") = "荳ｻ隕√・繝医Ν繝阪ャ繧ｯ遲九・讖溯・謾ｹ蝟・ｒ蝗ｳ繧九・
      Case Else
        result("Function_Short") = ""
    End Select

    result("Activity_Short") = BuildActivityShort_ByActivity(mainCause, result("Activity_Long"), mmtTargetMuscle, result("MMT_MinScore"))
    result("Participation_Long") = "遘ｻ蜍戊・蜉帙・蜷台ｸ翫↓繧医ｊ" & result("Activity_Long") & "縺ｮ讖滉ｼ壹ｒ謖√※繧狗憾諷九ｒ逶ｮ謖・☆縲・
      
    shortCore = Replace(result("Activity_Short"), "繧貞峙繧九・, "")


Select Case result("Activity_Long")

    Case "螻句､匁ｭｩ陦・
        result("Participation_Short") = shortCore & "繧貞峙繧翫∝､門・讖滉ｼ壹・諡｡螟ｧ縺ｫ蜷代￠縺滓ｺ門ｙ繧定｡後≧縲・

    Case "繝医う繝ｬ蜍穂ｽ・
        result("Participation_Short") = shortCore & "繧貞峙繧翫∬・遶区賜豕・ｩ滉ｼ壹・諡｡螟ｧ縺ｫ蜷代￠縺滓ｺ門ｙ繧定｡後≧縲・

    Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
        result("Participation_Short") = shortCore & "繧貞峙繧翫∝・豬ｴ閾ｪ遶区ｩ滉ｼ壹・諡｡螟ｧ縺ｫ蜷代￠縺滓ｺ門ｙ繧定｡後≧縲・

    Case "遘ｻ荵・
        result("Participation_Short") = shortCore & "繧貞峙繧翫∵律蟶ｸ逕滓ｴｻ蜀・ｧｻ蜍墓ｩ滉ｼ壹・諡｡螟ｧ縺ｫ蜷代￠縺滓ｺ門ｙ繧定｡後≧縲・

    Case Else
        result("Participation_Short") = shortCore & "繧貞峙繧翫・ & result("Activity_Long") & "縺ｮ讖滉ｼ壽僑螟ｧ縺ｫ蜷代￠縺滓ｺ門ｙ繧定｡後≧縲・

End Select
    


    
    Set BuildBasicPlanStructure = result
    
End Function

Public Function FilterMMTMap(ByVal mmtMap As Object, ByVal activityLong As String) As Object
    Dim candidateCsv As String
    Dim muscles() As String
    Dim filtered As Object
    Dim i As Long
    Dim keyName As String

    candidateCsv = GetCandidateMuscles(activityLong)

    If Len(Trim$(candidateCsv)) = 0 Then
        Set FilterMMTMap = mmtMap
        Exit Function
    End If

    Set filtered = CreateObject("Scripting.Dictionary")
    muscles = Split(candidateCsv, ",")

    For i = LBound(muscles) To UBound(muscles)
        keyName = Trim$(muscles(i))
        If Len(keyName) > 0 Then
            If mmtMap.exists(keyName) Then
                filtered(keyName) = mmtMap(keyName)
            End If
        End If
    Next i

    If filtered.count = 0 Then
        Set FilterMMTMap = mmtMap
    Else
        Set FilterMMTMap = filtered
    End If
End Function

Public Function GetCandidateMuscles(ByVal activityLong As String) As String
    Select Case activityLong
        Case "螻句・豁ｩ陦・
            GetCandidateMuscles = "閧｡螟冶ｻ｢,閭悟ｱ・閹昜ｼｸ螻・
        Case "螻句､匁ｭｩ陦・
            GetCandidateMuscles = "閭悟ｱ・閧｡螟冶ｻ｢,閹昜ｼｸ螻・
        Case "繝医う繝ｬ蜍穂ｽ・
            GetCandidateMuscles = "閧｡螟冶ｻ｢,閹昜ｼｸ螻・閭悟ｱ・
        Case "遶九■荳翫′繧・
            GetCandidateMuscles = "閹昜ｼｸ螻・閧｡莨ｸ螻・閧｡螟冶ｻ｢"
        Case "遘ｻ荵・
            GetCandidateMuscles = "閧｡螟冶ｻ｢,閹昜ｼｸ螻・
        Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
            GetCandidateMuscles = "閧｡螟冶ｻ｢,閹昜ｼｸ螻・閭悟ｱ・
        Case "襍ｷ螻・ｸ騾｣蜍穂ｽ・
            GetCandidateMuscles = "閧｡螟冶ｻ｢,閹昜ｼｸ螻・
        Case Else
            GetCandidateMuscles = ""
    End Select
End Function

Public Function PickActivityLong(ByVal needSelf As String, _
                                 ByVal needFamily As String, _
                                 ByVal needByDifficulty As String) As String
    Dim rawValue As String
    
    If Len(Trim$(needSelf)) > 0 Then
        rawValue = Trim$(needSelf)
    ElseIf Len(Trim$(needFamily)) > 0 Then
        rawValue = Trim$(needFamily)
    Else
        rawValue = Trim$(needByDifficulty)
    End If
    
    ' ---- 豁｣隕丞喧蜃ｦ逅・----
    Select Case rawValue
        Case "繝医う繝ｬ"
            PickActivityLong = "繝医う繝ｬ蜍穂ｽ・
        Case Else
            PickActivityLong = rawValue
    End Select
End Function


Public Function BuildActivityShort(ByVal mainCause As String, ByVal activityLong As String) As String
    
    Select Case mainCause
        Case "鮗ｻ逞ｺ"
            BuildActivityShort = activityLong & "譎ゅ・鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
            
        Case "逍ｼ逞・
            BuildActivityShort = activityLong & "譎ゅ・逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
            
        Case Else
            BuildActivityShort = activityLong & "蜍穂ｽ懊・螳牙ｮ壼喧繧貞峙繧九・
    End Select
    
    
    
    
End Function


Public Function BuildActivityShort_ByActivity(ByVal mainCause As String, _
                                              ByVal activityLong As String, _
                                              ByVal mmtTargetMuscle As String, _
                                              ByVal mmtMinScore As Double) As String
                                              
                                              
    Select Case activityLong
    
        Case "繝医う繝ｬ", "繝医う繝ｬ蜍穂ｽ・
            Select Case mainCause
                 Case "鮗ｻ逞ｺ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∽ｾｿ蠎ｧ遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∽ｾｿ蠎ｧ遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    End If
                    
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "遶九■荳翫′繧頑凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "譁ｹ蜷題ｻ｢謠帛虚菴懊・螳牙ｮ壼喧繧貞峙繧九・
                End Select
            
        Case "螻句・豁ｩ陦・
            Select Case mainCause
  Case "鮗ｻ逞ｺ"
    If mmtMinScore <= 2 Then
        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∝ｷｦ蜿ｳ闕ｷ驥榊ｷｮ縺ｮ霆ｽ貂帙ｒ蝗ｳ繧九・
    Else
        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∝ｷｦ蜿ｳ闕ｷ驥榊ｷｮ縺ｮ霆ｽ貂帙ｒ蝗ｳ繧九・
    End If
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "豁ｩ陦梧凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case "蝗ｰ髮｣蠎ｦ": BuildActivityShort_ByActivity = "譁ｹ蜷題ｻ｢謠帶凾縺ｮ螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                Case Else: BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
 
            End Select
            
        Case "螻句､匁ｭｩ陦・
            Select Case mainCause
                Case "鮗ｻ逞ｺ"
                   If mmtMinScore <= 2 Then
                       BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∵ｮｵ蟾ｮ譏・剄譎ゅ・鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                Else
                       BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∵ｮｵ蟾ｮ譏・剄譎ゅ・鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                End If
                
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "螻句､匁ｭｩ陦梧凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "谿ｵ蟾ｮ譏・剄譎ゅ・螳牙ｮ壽ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
            End Select
            
    
        Case "遘ｻ荵・
            Select Case mainCause
                Case "鮗ｻ逞ｺ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲√・繝・ラ繝ｻ讀・ｭ宣俣遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲√・繝・ラ繝ｻ讀・ｭ宣俣遘ｻ荵玲凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    End If
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "遘ｻ荵玲凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "遘ｻ荵怜虚菴懊・螳牙ｮ壼喧繧貞峙繧九・
            End Select
            
        Case "蜈･豬ｴ荳騾｣蜍穂ｽ・
            Select Case mainCause
                Case "鮗ｻ逞ｺ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∵ｵｴ螳､蜀・ｧｻ蜍輔・遶九■蠎ｧ繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∵ｵｴ螳､蜀・ｧｻ蜍輔・遶九■蠎ｧ繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    End If
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "蜈･豬ｴ蜍穂ｽ懈凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "蜈･豬ｴ荳騾｣蜍穂ｽ懊・螳牙ｮ壼喧繧貞峙繧九・
            End Select
            
        Case "襍ｷ螻・ｸ騾｣蜍穂ｽ・
            Select Case mainCause
                Case "鮗ｻ逞ｺ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∬ｵｷ縺堺ｸ翫′繧翫・遶九■荳翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∬ｵｷ縺堺ｸ翫′繧翫・遶九■荳翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    End If
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "襍ｷ螻・虚菴懈凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "襍ｷ螻・ｸ騾｣蜍穂ｽ懊・螳牙ｮ壼喧繧貞峙繧九・
            End Select
            
        Case "遶九■荳翫′繧・
            Select Case mainCause
                Case "鮗ｻ逞ｺ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ髫乗э蜿守ｸｮ迯ｲ蠕励ｒ騾壹§縺ｦ縲∫ｫ九■荳翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "縺ｮ遲句鴨謾ｹ蝟・ｒ騾壹§縺ｦ縲∫ｫ九■荳翫′繧頑凾縺ｮ鮗ｻ逞ｺ蛛ｴ謾ｯ謖∵ｧ蜷台ｸ翫ｒ蝗ｳ繧九・
                    End If
                Case "逍ｼ逞・: BuildActivityShort_ByActivity = "遶九■荳翫′繧頑凾縺ｮ逍ｼ逞幄ｻｽ貂帙ｒ蝗ｳ繧九・
                Case Else:  BuildActivityShort_ByActivity = "遶九■荳翫′繧雁虚菴懊・螳牙ｮ壼喧繧貞峙繧九・
            End Select
            
        Case Else
            BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
            
    End Select
    
End Function



Public Function DumpBasicPlan(ByVal plan As Object) As String
    Dim keys As Variant, i As Long, s As String
    
    keys = Array( _
        "MainCause", _
        "Activity_Long", _
        "Activity_Reason", _
        "Function_Long", _
        "Function_Short", _
        "Activity_Short", _
        "Participation_Long", _
        "Participation_Short" _
    )
    
    For i = LBound(keys) To UBound(keys)
        If plan.exists(keys(i)) Then
            s = s & plan(keys(i)) & vbCrLf
        Else
            s = s & "" & vbCrLf
        End If
    Next i
    
    DumpBasicPlan = s
End Function


Public Function DumpBasicGoalsOnly(ByVal plan As Object) As String
    Dim keys As Variant, i As Long, s As String
    
    keys = Array( _
    "Function_Short", _
    "Function_Long", _
    "Activity_Short", _
    "Activity_Long", _
    "Participation_Short", _
    "Participation_Long" _
)
    
    For i = LBound(keys) To UBound(keys)
        If plan.exists(keys(i)) Then
            s = s & plan(keys(i)) & vbCrLf
        Else
            s = s & "" & vbCrLf
        End If
    Next i
    
    DumpBasicGoalsOnly = s
End Function


Public Function PickMMTTarget(ByVal mmtMap As Object) As String

    Dim k As Variant
    Dim bestMuscle As String
    Dim bestScore As Double
    
    bestMuscle = ""
    bestScore = 9999
    
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) < bestScore Then
                bestScore = CDbl(mmtMap(k))
                bestMuscle = CStr(k)
            End If
        End If
    Next k
    
    PickMMTTarget = bestMuscle
End Function




Public Function PickMMTTarget_FromPairs(ParamArray pairs() As Variant) As String
    Dim d As Object
    Dim i As Long
    
    Set d = CreateObject("Scripting.Dictionary")
    
    i = LBound(pairs)
    Do While i <= UBound(pairs) - 1
        d(CStr(pairs(i))) = CDbl(pairs(i + 1))
        i = i + 2
    Loop
    
    PickMMTTarget_FromPairs = PickMMTTarget(d)
End Function




Public Function BuildBasicPlan_FromPairs( _
    ByVal mainCause As String, _
    ByVal needSelf As String, _
    ByVal needFamily As String, _
    ByVal needByDifficulty As String, _
    ParamArray mmtPairs() As Variant) As Object
    
    Dim d As Object
    Dim i As Long
    
    Set d = CreateObject("Scripting.Dictionary")
    
    i = LBound(mmtPairs)
    Do While i <= UBound(mmtPairs) - 1
        d(CStr(mmtPairs(i))) = CDbl(mmtPairs(i + 1))
        i = i + 2
    Loop
    
    Set BuildBasicPlan_FromPairs = _
        BuildBasicPlanStructure(mainCause, needSelf, needFamily, needByDifficulty, d)
End Function



Public Function PickMMTMinScore(ByVal mmtMap As Object) As Double
    Dim k As Variant
    Dim bestScore As Double
    
    bestScore = 9999
    
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) < bestScore Then
                bestScore = CDbl(mmtMap(k))
            End If
        End If
    Next k
    
    If bestScore = 9999 Then bestScore = 0
    PickMMTMinScore = bestScore
End Function




Public Function PickMMTTarget_WithPriority(ByVal mmtMap As Object, ByVal priorityCsv As String) As String
    Dim pri() As String, i As Long
    Dim best As String, bestScore As Double
    Dim k As Variant, sc As Double
    
    best = ""
    bestScore = 9999
    
    ' 譛蟆上せ繧ｳ繧｢繧貞叙繧・
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            sc = CDbl(mmtMap(k))
            If sc < bestScore Then bestScore = sc
        End If
    Next k
    
    If bestScore = 9999 Then
        PickMMTTarget_WithPriority = ""
        Exit Function
    End If
    
    ' 蜷檎紫縺ｮ荳ｭ縺ｧ蜆ｪ蜈磯・↓驕ｸ縺ｶ
    pri = Split(priorityCsv, ",")
    For i = LBound(pri) To UBound(pri)
        If mmtMap.exists(Trim$(pri(i))) Then
            If IsNumeric(mmtMap(Trim$(pri(i)))) Then
                If CDbl(mmtMap(Trim$(pri(i)))) = bestScore Then
                    PickMMTTarget_WithPriority = Trim$(pri(i))
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' 蜆ｪ蜈医Μ繧ｹ繝医↓辟｡縺代ｌ縺ｰ譛蛻昴↓隕九▽縺九▲縺滓怙蟆上ｒ霑斐☆
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) = bestScore Then
                PickMMTTarget_WithPriority = CStr(k)
                Exit Function
            End If
        End If
    Next k
End Function




Public Function GetLowerMMTMap_FromFrmEval() As Object
    Dim mp As Object, p As Object
    Dim c As Object
    Dim dict As Object
    Dim nm As String
    Dim vR As Double, vL As Double, vMin As Double
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim pgMMT As Object, host As Object
    Set pgMMT = GetMMTPage(frmEval)
    If pgMMT Is Nothing Then
        Set GetLowerMMTMap_FromFrmEval = dict
        Exit Function
    End If

    Set host = GetMMTHost(pgMMT)
    Set mp = GetMMTChildTabs(pgMMT, host)
    If mp Is Nothing Then
        Set GetLowerMMTMap_FromFrmEval = dict
        Exit Function
    End If
    
    Set p = mp.Pages(1) ' 荳玖い
    
    For Each c In p.controls
        If TypeName(c) = "Label" Then
            If Left$(c.name, 4) = "lbl_" Then
                nm = CStr(c.caption)
                
                vR = GetMMTValueSafe(p, "cboR_" & nm)
                vL = GetMMTValueSafe(p, "cboL_" & nm)
                
                vMin = vR
                If vL < vMin Then vMin = vL
                
                ' 譛ｪ蜈･蜉・99)縺ｯ謐ｨ縺ｦ繧・
                If vMin < 99 Then
                    dict(nm) = vMin
                End If
            End If
        End If
    Next c
    
    Set GetLowerMMTMap_FromFrmEval = dict
End Function

Private Function GetMMTValueSafe(ByVal container As Object, ByVal cboName As String) As Double
    On Error GoTo EH
    Dim v As String
    v = Trim$(container.controls(cboName).value & "")
    If Len(v) = 0 Then
        GetMMTValueSafe = 99
        Exit Function
    End If
    If IsNumeric(v) Then
        GetMMTValueSafe = CDbl(v)
    Else
        GetMMTValueSafe = 99
    End If
    Exit Function
EH:
    GetMMTValueSafe = 99
End Function


Public Function BuildBasicPlanStructureFromJudge(ByVal judged As Object) As Object
    Dim mainCause As String
    Dim needSelf As String
    Dim needFamily As String
    Dim needByDifficulty As String
    Dim mmtMap As Object

    mainCause = CStr(judged("MainCause"))
    needSelf = CStr(judged("NeedPatient"))
    needFamily = CStr(judged("NeedFamily"))
    needByDifficulty = CStr(judged("ActivityCandidate"))
    Set mmtMap = BuildMMTMapFromIO(CStr(judged("MMT_IO")))

    Set BuildBasicPlanStructureFromJudge = BuildBasicPlanStructure(mainCause, needSelf, needFamily, needByDifficulty, mmtMap)
End Function

Private Function BuildMMTMapFromIO(ByVal mmtIO As String) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    ' TODO: 譌｢蟄弄MT_IO繝輔か繝ｼ繝槭ャ繝医・豁｣蠑上ヱ繝ｼ繧ｵ繝ｼ縺ｫ鄂ｮ謠帙☆繧九・
    ' 譛菴朱剞縺ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ蛟､繧偵そ繝・ヨ縲・
    m("螟ｧ閻ｿ蝗幃ｭ遲・) = 3
    m("荳ｭ谿ｿ遲・) = 3
    m("閻ｸ閻ｰ遲・) = 3

    Set BuildMMTMapFromIO = m
End Function
