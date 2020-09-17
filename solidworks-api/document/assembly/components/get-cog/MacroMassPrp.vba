Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    
    Dim swComp As SldWorks.Component2
    
    Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, -1)
    
    Dim swMassPrps As SldWorks.MassProperty
    Set swMassPrps = swModel.Extension.CreateMassProperty()
    
    Dim vCompBodies As Variant
    vCompBodies = swComp.GetBodies3(swBodyType_e.swSolidBody, Empty)
    
    If False <> swMassPrps.AddBodies(vCompBodies) Then
    
        Dim vCog As Variant
        vCog = swMassPrps.CenterOfMass
        
        Debug.Print "COG: " & vCog(0) & "; " & vCog(1) & "; " & vCog(2)
    
    Else
        Err.Raise vbError, "", "Failed to add bodies for calculation"
    End If
    
End Sub