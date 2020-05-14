Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swEqMgr As SldWorks.EquationMgr

Const EQUATION = "sin(0.5) * 2 + (10 - 5)"

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Set swSelMgr = swModel.SelectionManager
        
        Dim swDispDim As SldWorks.DisplayDimension
        
        Set swDispDim = swSelMgr.GetSelectedObject6(1, -1)
                
        If Not swDispDim Is Nothing Then
                
            Set swEqMgr = swModel.GetEquationMgr
            
            Dim formula As String
            
            formula = """" & swDispDim.GetNameForSelection & """ = " & EQUATION
            
            swEqMgr.Add2 -1, formula, True
        
        Else
            MsgBox "Please select dimension"
        End If
    
    Else
        MsgBox "Please open model"
    End If
    
End Sub

