Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr

Sub main()

    On Error Resume Next

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim isSketchActive As Boolean
    
    isSketchActive = Not swModel.SketchManager.ActiveSketch Is Nothing
    
    If isSketchActive Then
        If Not swModel.SketchManager.ActiveSketch.Is3D Then
            MsgBox "Points can only be inserted into 3D sketch"
            End
        End If
    End If
    
    Dim swEdge As SldWorks.Edge
    
    Set swEdge = swSelMgr.GetSelectedObject6(1, -1)
    
    If Not swEdge Is Nothing Then
        
        Dim swCurve As SldWorks.Curve
        
        Set swCurve = swEdge.GetCurve
        
        Dim vPts As Variant
        
        Dim pointsCount As Integer
        pointsCount = CInt(InputBox("Specify the number of points"))
        
        If pointsCount <= 0 Then
            MsgBox "Please specify the valid integer number more than 1"
            End
        End If
        
        vPts = SplitCurveByPoints(swCurve, pointsCount)
    
        swModel.ClearSelection2 True
    
        If Not isSketchActive Then 'open new 3D sketch
            swModel.SketchManager.Insert3DSketch True
        End If
        
        Dim i As Integer
        
        For i = 0 To (UBound(vPts) + 1) / 3 - 1
        
            swModel.SketchManager.CreatePoint vPts(i * 3), vPts(i * 3 + 1), vPts(i * 3 + 2)
            
        Next
    
    If Not isSketchActive Then 'only close sketch if it wasn't opened at the beginning
        swModel.SketchManager.Insert3DSketch True
    End If
        
    Else
        MsgBox "Please select edge"
    End If
            
End Sub

Function SplitCurveByPoints(swCurve As SldWorks.Curve, pointsNumber As Integer) As Variant
    
    Dim nStartParam As Double
    Dim nEndParam As Double
    Dim bIsClosed As Boolean
    Dim bIsPeriodic As Boolean
    
    Dim incr As Double
    Dim i As Integer
    Dim vParam As Variant
    
    Dim retVal() As Double
    
    ReDim retVal(pointsNumber * 3 - 1)
    
    swCurve.GetEndParams nStartParam, nEndParam, bIsClosed, bIsPeriodic
    
    incr = (nEndParam - nStartParam) / (pointsNumber - 1)
    
    For i = 0 To pointsNumber - 1
    
        vParam = swCurve.Evaluate(nStartParam + i * incr)
        
        retVal(i * 3) = vParam(0)
        retVal(i * 3 + 1) = vParam(1)
        retVal(i * 3 + 2) = vParam(2)
        
    Next
    
    SplitCurveByPoints = retVal
    
End Function
