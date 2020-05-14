#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
 
Dim swApp As SldWorks.SldWorks
 
Sub main()
 
     Set swApp = Application.SldWorks
     
     Dim swAssy As SldWorks.AssemblyDoc
     Set swAssy = swApp.ActiveDoc
     
     If Not swAssy Is Nothing Then
     
        Dim swComp As SldWorks.Component2
        Set swComp = swAssy.SelectionManager.GetSelectedObjectsComponent3(1, -1)
        
        If swComp Is Nothing Then
            Set swComp = swAssy.ConfigurationManager.ActiveConfiguration.GetRootComponent3(False)
        End If
        
        ShowWithDependents swComp
        
     Else
        MsgBox "Please open assembly"
     End If
     
 End Sub
 
Sub ShowWithDependents(comp As SldWorks.Component2)
    
    comp.Select4 False, Nothing, False
    
    Const WM_COMMAND As Long = &H111
    Const SHOW_WITH_DEPENDENTS_CMD As Long = 33227
    
    Dim swFrame As SldWorks.Frame
    
    Set swFrame = swApp.Frame
    
    SendMessage swFrame.GetHWnd(), WM_COMMAND, SHOW_WITH_DEPENDENTS_CMD, 0
     
End Sub