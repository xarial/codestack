Const TRANSFORM As String = "1 0 0 0 1 0 0 0 1 0 0 0 1 0 0 0"
Const FIX_POSITION As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim vComps As Variant
    
        vComps = GetSelectedComponents(swModel)
        
        If Not IsEmpty(vComps) Then
            
            Dim swTransform As SldWorks.MathTransform
            
            Dim transformData As String
        
            If Not TryGetTransformFromArguments(transformData) Then
                transformData = TRANSFORM
            End If
            
            Set swTransform = ParseMathTransform(transformData)
            
            Dim swAssy As SldWorks.AssemblyDoc
            
            Set swAssy = swModel
            
            Dim i As Integer
            
            For i = 0 To UBound(vComps)
                Dim swComp As SldWorks.Component2
                Set swComp = vComps(i)
                swComp.Transform2 = swTransform
                
                If FIX_POSITION Then
                    If False <> swComp.Select4(False, Nothing, False) Then
                        swAssy.FixComponent
                    Else
                        Err.Raise vbError, "", "Failed to select component"
                    End If
                End If
                
            Next
            
            swModel.GraphicsRedraw2
            
        Else
            Err.Raise vbError, "", "Select components"
        End If
        
    Else
        Err.Raise vbError, "", "Open assembly"
    End If
    
End Sub

Function ParseMathTransform(matrix As String) As SldWorks.MathTransform
    
    Dim vMatrix As Variant
    vMatrix = Split(matrix, " ")
    
    If UBound(vMatrix) + 1 = 16 Then
    
        Dim dData(15) As Double
        Dim i As Integer
        
        For i = 0 To UBound(vMatrix)
            dData(i) = CDbl(Trim(vMatrix(i)))
        Next
        
        Dim swMathUtils As SldWorks.MathUtility
        Set swMathUtils = swApp.GetMathUtility
        
        Dim swTransform As SldWorks.MathTransform
        Set swTransform = swMathUtils.CreateTransform(dData)
        
        Set ParseMathTransform = swTransform
        
    Else
        Err.Raise vbError, "", "Transform must contain 16 elements separated by space"
    End If
    
End Function

Function GetSelectedComponents(model As SldWorks.ModelDoc2) As Variant
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    Dim swComps() As SldWorks.Component2

    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
                
        Dim swComp As SldWorks.Component2
    
        Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
        
        If Not swComp Is Nothing Then
            
            If (Not swComps) = -1 Then
                ReDim swComps(0)
                Set swComps(0) = swComp
            Else
                If Not Contains(swComps, swComp) Then
                    ReDim Preserve swComps(UBound(swComps) + 1)
                    Set swComps(UBound(swComps)) = swComp
                End If
            End If
                        
        End If
    
    Next

    If (Not swComps) = -1 Then
        GetSelectedComponents = Empty
    Else
        GetSelectedComponents = swComps
    End If

End Function

Function Contains(vArr As Variant, item As Object) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(vArr)
        If vArr(i) Is item Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function

Function TryGetTransformFromArguments(ByRef TRANSFORM As String) As Boolean

try_:

    On Error GoTo catch_

    Dim macroOprMgr As Object
    Set macroOprMgr = CreateObject("CadPlus.MacroOperationManager")
        
    Set macroOper = macroOprMgr.PopOperation(swApp)
    
    Dim vArgs As Variant
    vArgs = macroOper.Arguments
   
    Dim macroArg As Object
    Set macroArg = vArgs(0)
    
    TRANSFORM = CStr(macroArg.GetValue())
    TryGetTransformFromArguments = True
    GoTo finally_
    
catch_:
    TryGetTransformFromArguments = False
finally_:

End Function