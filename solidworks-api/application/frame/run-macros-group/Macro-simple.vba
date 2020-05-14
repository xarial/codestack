Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    RunMacro "C:\Macros\Macro1.swp", "Macro11", "main"
    RunMacro "C:\Macros\Macro2.swp", "Macro21", "main"
    RunMacro "C:\Macros\Macro3.swp", "Macro31", "main"
    
End Sub

Sub RunMacro(path As String, moduleName As String, procName As String)
    swApp.RunMacro2 path, moduleName, procName, swRunMacroOption_e.swRunMacroUnloadAfterRun, 0
End Sub