Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
    
    ArgumentHelper.SetArgument "Argument from VBA macro"
    
    Dim err As Long
    
    If False = swApp.RunMacro2("D:\Macros\GetArgumentMacro.swp", _
        "Macro1", "main", swRunMacroOption_e.swRunMacroUnloadAfterRun, err) Then
        
        swApp.SendMsgToUser "Failed to run macro. Error code: " & err
        
    End If
    
End Sub