Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim swView As SldWorks.view
        Set swView = swDraw.SelectionManager.GetSelectedObject6(1, -1)
        
        If Not swView Is Nothing Then
            DrawPoint swDraw, swView
        Else
            MsgBox "Please select drawing view"
        End If
    Else
        MsgBox "Please open the drawing document"
    End If
    
End Sub

Sub DrawPoint(draw As SldWorks.DrawingDoc, view As SldWorks.view)
    
    Dim vBoundings As Variant
    vBoundings = view.GetOutline()
    
    Dim dCenterPt(2) As Double
    dCenterPt(0) = (vBoundings(0) + vBoundings(2)) / 2
    dCenterPt(1) = (vBoundings(1) + vBoundings(3)) / 2
    dCenterPt(2) = 0
    
    Dim swViewSketch As SldWorks.Sketch
    Set swViewSketch = view.GetSketch
    
    Dim swViewSketchXForm As SldWorks.MathTransform
    Set swViewSketchXForm = swViewSketch.ModelToSketchTransform
    
    Dim swMathUtils As SldWorks.MathUtility
    Set swMathUtils = swApp.GetMathUtility
    
    Dim swMathPt As SldWorks.MathPoint
    Set swMathPt = swMathUtils.CreatePoint(dCenterPt)
    
    Set swMathPt = swMathPt.MultiplyTransform(swViewSketchXForm)
    
    draw.ActivateView view.Name
    
    Dim vPt As Variant
    vPt = swMathPt.ArrayData
    
    draw.SketchManager.CreatePoint vPt(0), vPt(1), vPt(2)
    
End Sub