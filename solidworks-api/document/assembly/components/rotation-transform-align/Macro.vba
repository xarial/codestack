Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swFace As SldWorks.Face2
    Dim swEdge As SldWorks.Edge
    
    Set swFace = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim swComp As SldWorks.Component2
    Set swComp = swFace.GetComponent()
    Dim swCompTransform As SldWorks.MathTransform
    Set swCompTransform = swComp.Transform2
    
    Set swEdge = swSelMgr.GetSelectedObject6(2, -1)
    
    Dim swMathUtils As SldWorks.MathUtility
    Set swMathUtils = swApp.GetMathUtility
    
    Dim swNormalDir As SldWorks.MathVector
    Set swNormalDir = swMathUtils.CreateVector(swFace.Normal)
    Set swNormalDir = swNormalDir.MultiplyTransform(swCompTransform)
    
    Dim swAlignDir As SldWorks.MathVector
    Dim vLineParams As Variant
    vLineParams = swEdge.GetCurve().lineParams
    Dim dVec(2) As Double
    dVec(0) = vLineParams(3): dVec(1) = vLineParams(4): dVec(2) = vLineParams(5)
    Set swAlignDir = swMathUtils.CreateVector(dVec)
    Set swAlignDir = swAlignDir.MultiplyTransform(swEdge.GetComponent().Transform2)
    
    Dim swOrigin As SldWorks.MathPoint
    Dim dOrigin(2) As Double
    dOrigin(0) = 0: dOrigin(1) = 0: dOrigin(2) = 0
    Set swOrigin = swMathUtils.CreatePoint(dOrigin)
    
    Set swOrigin = swOrigin.MultiplyTransform(swCompTransform)
    
    Dim swRotVect As SldWorks.MathVector
    Set swRotVect = swNormalDir.Cross(swAlignDir)
        
    Dim angle As Double
    angle = GetAngle(swNormalDir, swAlignDir)
    
    Dim swTransform As SldWorks.MathTransform
    Set swTransform = swMathUtils.CreateTransformRotateAxis(swOrigin, swRotVect, angle)
    
    Set swTransform = swCompTransform.Multiply(swTransform)
    
    swComp.Transform2 = swTransform
    
    swModel.GraphicsRedraw2
    
End Sub

Function GetAngle(vec1 As MathVector, vec2 As MathVector) As Double
    
    'cos a= a*b/(|a|*|b|)
    GetAngle = ACos(vec1.Dot(vec2) / (vec1.GetLength() * vec2.GetLength()))
    
End Function

Function ACos(val As Double) As Double
    
    If val = 1 Then
        ACos = 0
    ElseIf val = -1 Then
        ACos = 4 * Atn(1)
    Else
        ACos = Atn(-val / Sqr(-val * val + 1)) + 2 * Atn(1)
    End If
    
End Function