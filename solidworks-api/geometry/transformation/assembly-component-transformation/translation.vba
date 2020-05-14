Dim swApp As SldWorks.SldWorks
Dim swMathUtils As SldWorks.MathUtility
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swComp As SldWorks.Component2

Sub main()

    Set swApp = Application.SldWorks
    
    Set swMathUtils = swApp.GetMathUtility
    
    Set swModel = swApp.ActiveDoc
    
    Set swSelMgr = swModel.SelectionManager
    
    Set swComp = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim swTransform As SldWorks.MathTransform
    Set swTransform = swComp.Transform2
    
    Dim dOrigPt(2) As Double
    dOrigPt(0) = 0: dOrigPt(1) = 0: dOrigPt(2) = 0
    
    Dim swMathPt As SldWorks.MathPoint
    
    Set swMathPt = swMathUtils.CreatePoint(dOrigPt)
    
    Set swMathPt = swMathPt.MultiplyTransform(swTransform)
    
    Dim vCompOriginPt As Variant

    vCompOriginPt = swMathPt.ArrayData
    
    Debug.Print "Along X: " & vCompOriginPt(0) * 1000 & "mm; " & "Along Y: " & vCompOriginPt(1) * 1000 & "mm; " & "Along Z: " & vCompOriginPt(2) * 1000 & "mm"
    
End Sub