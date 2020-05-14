Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swComp As SldWorks.Component2

Sub main()

    Set swApp = Application.SldWorks

    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Set swSelMgr = swModel.SelectionManager
        
        Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, -1)
        
        If Not swComp Is Nothing Then
        
            Dim instId As Integer
            Dim compName As String
            compName = swComp.Name2
            instId = CInt(Right(compName, Len(compName) - InStrRev(compName, "-")))
            
            MsgBox "Selected component's instance id is " & instId
                
        Else
            
            MsgBox "Please select component"
            
        End If
        
    Else
        
        MsgBox "Please open assembly"
        
    End If
    
End Sub

