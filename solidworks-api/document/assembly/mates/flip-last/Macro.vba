Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swAssy As SldWorks.AssemblyDoc
    
    Set swAssy = swApp.ActiveDoc
    
    Dim swLastMate As SldWorks.Mate2
    
    Set swLastMate = GetLastMate(swAssy)
    
    Dim curAlignment As swMateAlign_e
    curAlignment = swLastMate.Alignment
    
    Dim destAlignment As swMateAlign_e
    
    If curAlignment = swMateAlignALIGNED Then
        destAlignment = swMateAlignANTI_ALIGNED
    ElseIf curAlignment = swMateAlignANTI_ALIGNED Then
        destAlignment = swMateAlignALIGNED
    Else
        Exit Sub
    End If
    
    Dim swMateFeat As SldWorks.Feature
    Set swMateFeat = swLastMate
    
    Dim swMateFeatData As SldWorks.MateFeatureData
    
    Set swMateFeatData = swMateFeat.GetDefinition
    
    Select Case swMateFeatData.TypeName
        Case swMateType_e.swMateANGLE
            Dim swAngleMateFeatData As SldWorks.AngleMateFeatureData
            Set swAngleMateFeatData = swMateFeatData
            swAngleMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateCAMFOLLOWER
            Dim swCamFollowerMateFeatData As SldWorks.CamFollowerMateFeatureData
            Set swCamFollowerMateFeatData = swMateFeatData
            swCamFollowerMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateCOINCIDENT
            Dim swCoincidentMateFeatData As SldWorks.CoincidentMateFeatureData
            Set swCoincidentMateFeatData = swMateFeatData
            swCoincidentMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateCONCENTRIC
            Dim swConcentricMateFeatData As SldWorks.ConcentricMateFeatureData
            Set swConcentricMateFeatData = swMateFeatData
            swConcentricMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateDISTANCE
            Dim swDistanceMateFeatData As SldWorks.DistanceMateFeatureData
            Set swDistanceMateFeatData = swMateFeatData
            swDistanceMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateHINGE
            Dim swHingeMateFeatData As SldWorks.HingeMateFeatureData
            Set swHingeMateFeatData = swMateFeatData
            swHingeMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMatePARALLEL
            Dim swParallelMateFeatData As SldWorks.ParallelMateFeatureData
            Set swParallelMateFeatData = swMateFeatData
            swParallelMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMatePROFILECENTER
            Dim swProfileCenterMateFeatData As SldWorks.ProfileCenterMateFeatureData
            Set swProfileCenterMateFeatData = swMateFeatData
            swProfileCenterMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateSCREW
            Dim swScrewMateFeatData As SldWorks.ScrewMateFeatureData
            Set swScrewMateFeatData = swMateFeatData
            swScrewMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateSLOT
            Dim swSlotMateFeatData As SldWorks.SlotMateFeatureData
            Set swSlotMateFeatData = swMateFeatData
            swSlotMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateSYMMETRIC
            Dim swSymmetricMateFeatData As SldWorks.SymmetricMateFeatureData
            Set swSymmetricMateFeatData = swMateFeatData
            swSymmetricMateFeatData.MateAlignment = destAlignment
        Case swMateType_e.swMateTANGENT
            Dim swTangentMateFeatData As SldWorks.TangentMateFeatureData
            Set swTangentMateFeatData = swMateFeatData
            swTangentMateFeatData.MateAlignment = destAlignment
        Case Else
            Err.Raise vbError, "", "Not supported mate type"
    End Select
        
    swMateFeat.ModifyDefinition swMateFeatData, swAssy, Nothing
    
End Sub

Function GetLastMate(assm As SldWorks.AssemblyDoc) As SldWorks.Mate2
    
    Dim swMates() As SldWorks.Feature
    Dim isInit As Boolean
    isInit = False
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = assm
    
    Dim swMateGroupFeat As SldWorks.Feature
    
    Dim featIndex As Integer
    featIndex = 0
        
    Do
        Set swMateGroupFeat = swModel.FeatureByPositionReverse(featIndex)
        
        featIndex = featIndex + 1
    Loop While swMateGroupFeat.GetTypeName2() <> "MateGroup"
    
    Dim swLastMateFeat As SldWorks.Feature
    
    Dim swMateFeat As SldWorks.Feature
    
    Set swMateFeat = swMateGroupFeat.GetFirstSubFeature
    
    While Not swMateFeat Is Nothing
        
        If TypeOf swMateFeat.GetSpecificFeature2() Is SldWorks.Mate2 Then
            Set swLastMateFeat = swMateFeat
        End If
        
        Set swMateFeat = swMateFeat.GetNextSubFeature
    Wend
    
    Debug.Print swLastMateFeat.Name
    
    Set GetLastMate = swLastMateFeat.GetSpecificFeature2
    
End Function