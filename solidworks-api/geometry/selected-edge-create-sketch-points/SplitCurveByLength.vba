Function SplitCurveByLength(swCurve As SldWorks.Curve, chordLength As Double) As Variant
    
    Dim nStartParam As Double
    Dim nEndParam As Double
    Dim bIsClosed As Boolean
    Dim bIsPeriodic As Boolean
        
    swCurve.GetEndParams nStartParam, nEndParam, bIsClosed, bIsPeriodic
    
    SplitCurveByLength = swCurve.GetTessPts(0.01, chordLength, swCurve.Evaluate2(nStartParam, 1), swCurve.Evaluate2(nEndParam, 1))
    
End Function