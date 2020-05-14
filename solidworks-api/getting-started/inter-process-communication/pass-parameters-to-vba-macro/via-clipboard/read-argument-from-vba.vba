Dim swApp As SldWorks.SldWorks

Sub main()
        
    Set swApp = Application.SldWorks
        
     swApp.SendMsgToUser "Specified argument: " & ArgumentHelper.GetArgument()
    
End Sub