Dim swApp As SldWorks.SldWorks
Dim swAssy As SldWorks.AssemblyDoc

Sub main()

    Set swApp = Application.SldWorks
    
    Set swAssy = swApp.ActiveDoc
    
    swAssy.ResolveAllLightWeightComponents True
    
    Dim swComp As SldWorks.Component2
    Set swComp = swAssy.SelectionManager.GetSelectedObject6(1, -1)
        
    Dim swRefModel As SldWorks.ModelDoc2
    Set swRefModel = swComp.GetModelDoc2
        
    If Not swRefModel Is Nothing Then 'Check if the pointer is not nothing
        MsgBox "Material of " & swComp.Name2 & ": " & swRefModel.MaterialIdName
    Else
        MsgBox "Component's model is not loaded into the memory" 'display the error
    End If
    
End Sub
