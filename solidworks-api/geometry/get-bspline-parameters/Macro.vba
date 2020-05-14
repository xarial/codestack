Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swEdge As SldWorks.Edge
    
    Set swEdge = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim swCurve As SldWorks.Curve
    
    Set swCurve = swEdge.GetCurve
    
    Dim swSplineData As SldWorks.SplineParamData
    Set swSplineData = swCurve.GetBCurveParams5(False, False, False, False)
    
    Dim i As Integer
    
    Debug.Print "Props:"
    Debug.Print swSplineData.Dimension
    Debug.Print swSplineData.Order
    Debug.Print swSplineData.ControlPointsCount
    Debug.Print swSplineData.Periodic
    
    Debug.Print "Knots:"
    Dim vKnotPts As Variant
    swSplineData.GetKnotPoints vKnotPts
    
    For i = 0 To UBound(vKnotPts)
        Debug.Print vKnotPts(i)
    Next
    
    Debug.Print "Control Points:"
    Dim vCtrlPts As Variant
    swSplineData.GetControlPoints vCtrlPts
    For i = 0 To UBound(vCtrlPts)
        Debug.Print vCtrlPts(i)
    Next
    
End Sub