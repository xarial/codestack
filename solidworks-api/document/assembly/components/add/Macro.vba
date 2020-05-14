Dim swApp As SldWorks.SldWorks
Dim swMathUtils As SldWorks.MathUtility
Dim swAssy As SldWorks.AssemblyDoc

Sub main()

    Set swApp = Application.SldWorks
    
    Set swMathUtils = swApp.GetMathUtility
    
    Set swAssy = swApp.ActiveDoc
    
    If Not swAssy Is Nothing Then
        
        Dim comp1Path As String
        Dim comp2Path As String
        
        comp1Path = swApp.GetCurrentMacroPathFolder() & "\Part1.sldprt"
        comp2Path = swApp.GetCurrentMacroPathFolder() & "\Part2.sldprt"
        
        Dim swComp As SldWorks.Component2
        
        'Following API call will fail as it is required to have the model loaded into the memory
        Set swComp = swAssy.AddComponent4(comp1Path, "", 0, 0, 0)
        
        Debug.Assert Not swComp Is Nothing
                
        'Loading model invisibly
        swApp.DocumentVisible False, swDocumentTypes_e.swDocPART
        swApp.OpenDoc6 comp1Path, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0
        swApp.DocumentVisible True, swDocumentTypes_e.swDocPART
        
        'Now this API call succeeded
        Set swComp = swAssy.AddComponent4(comp1Path, "", 0, 0, 0)
        
        Debug.Assert Not swComp Is Nothing
        
        Dim strCompNames(0) As String
        Dim vTransformData As Variant
        Dim vComps As Variant
        strCompNames(0) = comp2Path
        
        vTransformData = swMathUtils.CreateTransform(Empty).ArrayData
        
        'It is not required to load document into memory if this method is used
        vComps = swAssy.AddComponents(strCompNames, vTransformData)
    
        Debug.Assert UBound(vComps) <> 1
        
    Else
        
        MsgBox "Please open or create assembly"
        
    End If

End Sub

