Const PI As Double = 3.14159265359

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swSelMgr As SldWorks.SelectionMgr
        
        Set swSelMgr = swModel.SelectionManager
        
        Dim swBody As SldWorks.Body2
        Set swBody = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swBody Is Nothing Then
            RunRotationAnimation swModel, swBody
        Else
            MsgBox "Please select body"
        End If
        
    Else
        MsgBox "Please open part document"
    End If
    
End Sub

Sub RunRotationAnimation(part As SldWorks.PartDoc, body As SldWorks.Body2, Optional speed As Double = 1)
    
    body.HideBody True
    
    Dim rotStep As Double
    rotStep = PI * 2 / 360 * speed
    
    Dim curAng As Double
    
    Dim swModelView As SldWorks.ModelView
    Set swModelView = part.ActiveView
    
    Dim swTempBody As SldWorks.Body2
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = part.SelectionManager
    
    While swSelMgr.GetSelectedObjectCount2(-1) <> 0
        For curAng = 0 To PI * 2 Step rotStep
            Dim animStep As MathTransform
            Set animStep = GetTransform(curAng)
            Set swTempBody = Nothing
            Set swTempBody = body.Copy()
            swTempBody.ApplyTransform animStep
            swTempBody.Display3 part, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
            swModelView.GraphicsRedraw Nothing
            DoEvents
        Next
    Wend
    
    Set swTempBody = Nothing
    body.HideBody False
    
End Sub

Function GetTransform(angle As Double) As MathTransform
    
    Dim swMathUtils As SldWorks.MathUtility
    
    Set swMathUtils = swApp.GetMathUtility
    
    Dim swOrigPt As SldWorks.MathPoint
    Dim dPt(2) As Double
    dPt(0) = 0: dPt(1) = 0: dPt(2) = 0
    
    Set swOrigPt = swMathUtils.CreatePoint(dPt)
    
    Dim swAxisVec As SldWorks.MathVector
    Dim dVec(2) As Double
    dVec(0) = 0: dVec(1) = 1: dVec(2) = 0
    
    Set swAxisVec = swMathUtils.CreateVector(dVec)
    
    Set GetTransform = swMathUtils.CreateTransformRotateAxis(swOrigPt, swAxisVec, angle)
    
End Function