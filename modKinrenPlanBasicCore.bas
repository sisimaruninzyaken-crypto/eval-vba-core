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
          Case "‰®“à•às"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "ŒÒŠO“],”w‹ü,•GL“W")
          Case "ƒgƒCƒŒ“®ì"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "ŒÒŠO“],•GL“W")
          Case "‰®ŠO•às"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "”w‹ü,ŒÒŠO“],•GL“W")
          Case "—§‚¿ã‚ª‚è"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "•GL“W,ŒÒŠO“],ŒÒL“W")
          Case Else
              mmtTargetMuscle = PickMMTTarget(mmtMap)
    End Select

    result("MMT_TargetMuscle") = mmtTargetMuscle
    result("MMT_MinScore") = mmtMap(mmtTargetMuscle)
    result("MainCause") = mainCause

    
        If Len(Trim$(needSelf)) > 0 Then
          reason = "–{lŠó–]"
        ElseIf Len(Trim$(needFamily)) > 0 Then
          reason = "‰Æ‘°Šó–]"
        Else
          reason = "¢“ï“xãˆÊ"
        End If

    result("Activity_Reason") = reason

    Select Case mainCause
        Case "–ƒáƒ"
        
        If result("MMT_MinScore") <= 2 Then
            fxCore = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚É‚æ‚è"
        Else
            fxCore = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚É‚æ‚è"
        End If

    Select Case result("Activity_Long")
    
        Case "‰®“à•às"
            result("Function_Long") = fxCore & "—§‹rŠúˆÀ’è«Œüã‚ğ}‚éB"
        
        Case "—§‚¿ã‚ª‚è"
            result("Function_Long") = fxCore & "—§‚¿ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
        
        Case "ƒgƒCƒŒ", "ƒgƒCƒŒ“®ì"
            result("Function_Long") = fxCore & "•ÖÀˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
        
        Case "‰®ŠO•às"
            result("Function_Long") = fxCore & "’i·¸~‚ÌˆÀ’è«Œüã‚ğ}‚éB"
        
        Case "“ü—ˆê˜A“®ì"
            result("Function_Long") = fxCore & "—‘…‚Ü‚½‚¬“®ì‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"

        Case "ˆÚæ"
            result("Function_Long") = fxCore & "ˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"

        Case "‹N‹ˆê˜A“®ì"
            result("Function_Long") = fxCore & "‹N‚«ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
           
        Case Else
            result("Function_Long") = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ}‚éB"
            
    End Select
   
    Case "áu’É"

    Select Case result("Activity_Long")

        Case "‰®“à•às"
            result("Function_Long") = "•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case "‰®ŠO•às"
            result("Function_Long") = "‰®ŠO•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case "ƒgƒCƒŒ“®ì"
            result("Function_Long") = "—§‚¿ã‚ª‚è‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case "“ü—ˆê˜A“®ì"
            result("Function_Long") = "“ü—“®ì‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case "ˆÚæ"
            result("Function_Long") = "ˆÚæ“®ì‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case "‹N‹ˆê˜A“®ì"
            result("Function_Long") = "‹N‹“®ì‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"

        Case Else
            result("Function_Long") = "áu’É‚ÌŒyŒ¸‚ğ}‚éB"

    End Select
       
    Case "¢“ï“x"

    Select Case result("Activity_Long")

        Case "‰®“à•às"
            result("Function_Long") = "•ûŒü“]Š·‚ÌˆÀ’è«Œüã‚ğ}‚éB"

        Case "‰®ŠO•às"
            result("Function_Long") = "’i·¸~‚ÌˆÀ’è«Œüã‚ğ}‚éB"

        Case "ƒgƒCƒŒ“®ì"
            result("Function_Long") = "•ûŒü“]Š·“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"

        Case "“ü—ˆê˜A“®ì"
            result("Function_Long") = "—º“à•ûŒü“]Š·‚ÌˆÀ’è«Œüã‚ğ}‚éB"

        Case "ˆÚæ"
            result("Function_Long") = "‘¤•ûˆÚ“®‚ÌˆÀ’è«Œüã‚ğ}‚éB"

        Case "‹N‹ˆê˜A“®ì"
            result("Function_Long") = "‹N‚«ã‚ª‚è“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"

        Case Else
            result("Function_Long") = "‰ºˆ‹@”\‚Ì‘S‘Ì“IŒüã‚ğ}‚éB"

    End Select
    
    Case Else
        result("Function_Long") = ""
    End Select

    Select Case mainCause
      Case "–ƒáƒ"
       If result("MMT_MinScore") <= 2 Then
    result("Function_Short") = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ}‚éB"
Else
    result("Function_Short") = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ}‚éB"
End If

      Case "áu’É"
        result("Function_Short") = "áu’É—U”­“®ì‚ÌŒyŒ¸‚¨‚æ‚Ñ•‰‰×’²®‚ğ}‚éB"
      Case "¢“ï“x"
        result("Function_Short") = "å—vƒ{ƒgƒ‹ƒlƒbƒN‹Ø‚Ì‹@”\‰ü‘P‚ğ}‚éB"
      Case Else
        result("Function_Short") = ""
    End Select

    result("Activity_Short") = BuildActivityShort_ByActivity(mainCause, result("Activity_Long"), mmtTargetMuscle, result("MMT_MinScore"))
    result("Participation_Long") = "ˆÚ“®”\—Í‚ÌŒüã‚É‚æ‚è" & result("Activity_Long") & "‚Ì‹@‰ï‚ğ‚Ä‚éó‘Ô‚ğ–Úw‚·B"
      
    shortCore = Replace(result("Activity_Short"), "‚ğ}‚éB", "")


Select Case result("Activity_Long")

    Case "‰®ŠO•às"
        result("Participation_Short") = shortCore & "‚ğ}‚èAŠOo‹@‰ï‚ÌŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"

    Case "ƒgƒCƒŒ“®ì"
        result("Participation_Short") = shortCore & "‚ğ}‚èA©—§”rŸ•‹@‰ï‚ÌŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"

    Case "“ü—ˆê˜A“®ì"
        result("Participation_Short") = shortCore & "‚ğ}‚èA“ü—©—§‹@‰ï‚ÌŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"

    Case "ˆÚæ"
        result("Participation_Short") = shortCore & "‚ğ}‚èA“úí¶Šˆ“àˆÚ“®‹@‰ï‚ÌŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"

    Case Else
        result("Participation_Short") = shortCore & "‚ğ}‚èA" & result("Activity_Long") & "‚Ì‹@‰ïŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"

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
        Case "‰®“à•às"
            GetCandidateMuscles = "ŒÒŠO“],”w‹ü,•GL“W"
        Case "‰®ŠO•às"
            GetCandidateMuscles = "”w‹ü,ŒÒŠO“],•GL“W"
        Case "ƒgƒCƒŒ“®ì"
            GetCandidateMuscles = "ŒÒŠO“],•GL“W,”w‹ü"
        Case "—§‚¿ã‚ª‚è"
            GetCandidateMuscles = "•GL“W,ŒÒL“W,ŒÒŠO“]"
        Case "ˆÚæ"
            GetCandidateMuscles = "ŒÒŠO“],•GL“W"
        Case "“ü—ˆê˜A“®ì"
            GetCandidateMuscles = "ŒÒŠO“],•GL“W,”w‹ü"
        Case "‹N‹ˆê˜A“®ì"
            GetCandidateMuscles = "ŒÒŠO“],•GL“W"
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
    
    ' ---- ³‹K‰»ˆ— ----
    Select Case rawValue
        Case "ƒgƒCƒŒ"
            PickActivityLong = "ƒgƒCƒŒ“®ì"
        Case Else
            PickActivityLong = rawValue
    End Select
End Function


Public Function BuildActivityShort(ByVal mainCause As String, ByVal activityLong As String) As String
    
    Select Case mainCause
        Case "–ƒáƒ"
            BuildActivityShort = activityLong & "‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
            
        Case "áu’É"
            BuildActivityShort = activityLong & "‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
            
        Case Else
            BuildActivityShort = activityLong & "“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
    End Select
    
    
    
    
End Function


Public Function BuildActivityShort_ByActivity(ByVal mainCause As String, _
                                              ByVal activityLong As String, _
                                              ByVal mmtTargetMuscle As String, _
                                              ByVal mmtMinScore As Double) As String
                                              
                                              
    Select Case activityLong
    
        Case "ƒgƒCƒŒ", "ƒgƒCƒŒ“®ì"
            Select Case mainCause
                 Case "–ƒáƒ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA•ÖÀˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA•ÖÀˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    End If
                    
                Case "áu’É": BuildActivityShort_ByActivity = "—§‚¿ã‚ª‚è‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "•ûŒü“]Š·“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
                End Select
            
        Case "‰®“à•às"
            Select Case mainCause
  Case "–ƒáƒ"
    If mmtMinScore <= 2 Then
        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA¶‰E‰×d·‚ÌŒyŒ¸‚ğ}‚éB"
    Else
        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA¶‰E‰×d·‚ÌŒyŒ¸‚ğ}‚éB"
    End If
                Case "áu’É": BuildActivityShort_ByActivity = "•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case "¢“ï“x": BuildActivityShort_ByActivity = "•ûŒü“]Š·‚ÌˆÀ’è«Œüã‚ğ}‚éB"
                Case Else: BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
 
            End Select
            
        Case "‰®ŠO•às"
            Select Case mainCause
                Case "–ƒáƒ"
                   If mmtMinScore <= 2 Then
                       BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA’i·¸~‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                Else
                       BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA’i·¸~‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                End If
                
                Case "áu’É": BuildActivityShort_ByActivity = "‰®ŠO•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "’i·¸~‚ÌˆÀ’è«Œüã‚ğ}‚éB"
            End Select
            
    
        Case "ˆÚæ"
            Select Case mainCause
                Case "–ƒáƒ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄAƒxƒbƒhEˆÖqŠÔˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄAƒxƒbƒhEˆÖqŠÔˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    End If
                Case "áu’É": BuildActivityShort_ByActivity = "ˆÚæ‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "ˆÚæ“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
            End Select
            
        Case "“ü—ˆê˜A“®ì"
            Select Case mainCause
                Case "–ƒáƒ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA—º“àˆÚ“®E—§‚¿À‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA—º“àˆÚ“®E—§‚¿À‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    End If
                Case "áu’É": BuildActivityShort_ByActivity = "“ü—“®ì‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "“ü—ˆê˜A“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
            End Select
            
        Case "‹N‹ˆê˜A“®ì"
            Select Case mainCause
                Case "–ƒáƒ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA‹N‚«ã‚ª‚èE—§‚¿ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA‹N‚«ã‚ª‚èE—§‚¿ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    End If
                Case "áu’É": BuildActivityShort_ByActivity = "‹N‹“®ì‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "‹N‹ˆê˜A“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
            End Select
            
        Case "—§‚¿ã‚ª‚è"
            Select Case mainCause
                Case "–ƒáƒ"
                    If mmtMinScore <= 2 Then
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA—§‚¿ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    Else
                        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA—§‚¿ã‚ª‚è‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                    End If
                Case "áu’É": BuildActivityShort_ByActivity = "—§‚¿ã‚ª‚è‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "—§‚¿ã‚ª‚è“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
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
    
    ' Å¬ƒXƒRƒA‚ğæ‚é
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
    
    ' “¯—¦‚Ì’†‚Å—Dæ‡‚É‘I‚Ô
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
    
    ' —DæƒŠƒXƒg‚É–³‚¯‚ê‚ÎÅ‰‚ÉŒ©‚Â‚©‚Á‚½Å¬‚ğ•Ô‚·
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
    
    Set p = mp.Pages(1) ' ‰ºˆ
    
    For Each c In p.controls
        If typeName(c) = "Label" Then
            If Left$(c.name, 4) = "lbl_" Then
                nm = CStr(c.caption)
                
                vR = GetMMTValueSafe(p, "cboR_" & nm)
                vL = GetMMTValueSafe(p, "cboL_" & nm)
                
                vMin = vR
                If vL < vMin Then vMin = vL
                
                ' –¢“ü—Í(99)‚ÍÌ‚Ä‚é
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
    Dim result As Object

    mainCause = CStr(judged("MainCause"))
    needSelf = CStr(judged("NeedPatient"))
    needFamily = CStr(judged("NeedFamily"))
    needByDifficulty = CStr(judged("ActivityCandidate"))
    Set mmtMap = BuildMMTMapFromIO(CStr(judged("MMT_IO")))

    Set result = BuildBasicPlanStructure(mainCause, needSelf, needFamily, needByDifficulty, mmtMap)
    result("FunctionCandidate") = CStr(judged("FunctionCandidate"))
    result("TrunkROMLimitTags") = CStr(judged("TrunkROMLimitTags"))
    result("EvalTestNote") = CStr(judged("EvalTestNote"))

    Set BuildBasicPlanStructureFromJudge = result
End Function

Private Function BuildMMTMapFromIO(ByVal mmtIO As String) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    ' TODO: Šù‘¶MMT_IOƒtƒH[ƒ}ƒbƒg‚Ì³®ƒp[ƒT[‚É’uŠ·‚·‚éB
    ' Å’áŒÀ‚ÌƒtƒH[ƒ‹ƒoƒbƒN’l‚ğƒZƒbƒgB
    m("‘å‘Úl“ª‹Ø") = 3
    m("’†“a‹Ø") = 3
    m("’°˜‹Ø") = 3

    Set BuildMMTMapFromIO = m
End Function
