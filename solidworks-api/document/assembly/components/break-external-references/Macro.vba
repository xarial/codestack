Dim swApp As SldWorks.SldWorks
Dim swAssy As SldWorks.AssemblyDoc

Sub main()

    Set swApp = Application.SldWorks
    
    Set swAssy = swApp.ActiveDoc
    
    If Not swAssy Is Nothing Then
    
        swAssy.ResolveAllLightWeightComponents True
        
        Dim vComps As Variant
        Dim swComp As SldWorks.Component2
        Dim swCompModel As SldWorks.ModelDoc2
        
        vComps = swAssy.GetComponents(False)
        
        Dim i As Integer
        
        For i = 0 To UBound(vComps)
            Set swComp = vComps(i)
            Set swCompModel = swComp.GetModelDoc2
            If Not swCompModel Is Nothing Then
                swCompModel.Extension.BreakAllExternalFileReferences2 False
            End If
        Next
    Else
        MsgBox "Please open assembly"
    End If
     
End Sub
