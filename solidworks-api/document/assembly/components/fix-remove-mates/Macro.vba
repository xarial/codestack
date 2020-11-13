Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() <> swDocumentTypes_e.swDocASSEMBLY Then
            Err.Raise vbError, "Only assembly document is supported"
        End If
        
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = swModel
        
        Dim vMates As Variant
        vMates = GetAllMates(swAssy)
        
        If Not IsEmpty(vMates) Then
            
            If swModel.Extension.MultiSelect2(vMates, False, Nothing) = UBound(vMates) + 1 Then
                
                If False <> swModel.Extension.DeleteSelection2(swDeleteSelectionOptions_e.swDelete_Absorbed) Then
                    
                    Dim vComps As Variant
                    vComps = swAssy.GetComponents(True)
                    
                    If swModel.Extension.MultiSelect2(vComps, False, Nothing) = UBound(vComps) + 1 Then
                        swAssy.FixComponent
                    Else
                        Err.Raise vbError, "", "Faield to select components"
                    End If
                    
                Else
                    Err.Raise vbError, "", "Failed to delete mates"
                End If
            Else
                Err.Raise vbError, "", "Failed to select mates for deletion"
            End If
        Else
            Err.Raise vbError, "", "No mates in the assembly"
        End If
    Else
        Err.Raise vbError, "", "Please open assemby document"
    End If
    
End Sub

Function GetAllMates(assm As SldWorks.AssemblyDoc) As Variant
    
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
    
    Dim swMateFeat As SldWorks.Feature
    
    Set swMateFeat = swMateGroupFeat.GetFirstSubFeature
    
    While Not swMateFeat Is Nothing
        
        If TypeOf swMateFeat.GetSpecificFeature2() Is SldWorks.Mate2 Then
            If isInit Then
                ReDim Preserve swMates(UBound(swMates) + 1)
            Else
                ReDim swMates(0)
                isInit = True
            End If
            Set swMates(UBound(swMates)) = swMateFeat
        End If
        
        Set swMateFeat = swMateFeat.GetNextSubFeature
    Wend
    
    GetAllMates = swMates
    
End Function