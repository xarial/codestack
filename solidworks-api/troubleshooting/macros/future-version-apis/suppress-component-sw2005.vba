Dim swApp As SldWorks.SldWorks
Dim swAssy As SldWorks.AssemblyDoc

Sub main()

    Set swApp = Application.SldWorks
    
    Set swAssy = swApp.ActiveDoc
    
    swAssy.CompConfigProperties4 swComponentSuppressionState_e.swComponentSuppressed, _
            swComponentSolvingOption_e.swComponentRigidSolving, _
            True, False, "", False
    
End Sub