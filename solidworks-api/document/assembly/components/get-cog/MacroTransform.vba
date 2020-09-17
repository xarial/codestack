Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    
    Dim swComp As SldWorks.Component2
    
    Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, -1)
    
    Dim swCompModel As SldWorks.ModelDoc2
    Set swCompModel = swComp.GetModelDoc2
    
    Const ACCURACY_DEFAULT As Integer = 1
    Dim status As swMassPropertiesStatus_e
    
    Dim vMassPrps As Variant
    vMassPrps = swCompModel.Extension.GetMassProperties2(ACCURACY_DEFAULT, status, False)
    
    Dim dCog(2) As Double
    
    dCog(0) = vMassPrps(0): dCog(1) = vMassPrps(1): dCog(2) = vMassPrps(2)
    
    Dim swMathUtils As SldWorks.MathUtility
    
    Set swMathUtils = swApp.GetMathUtility
    
    Dim swMathPt As SldWorks.MathPoint
    Set swMathPt = swMathUtils.CreatePoint(dCog)
    
    Set swMathPt = swMathPt.MultiplyTransform(swComp.Transform2)
    
    Dim vCog As Variant
    vCog = swMathPt.ArrayData
    
    Debug.Print "COG: " & vCog(0) & "; " & vCog(1) & "; " & vCog(2)
    
End Sub