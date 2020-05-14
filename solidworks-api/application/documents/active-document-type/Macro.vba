Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Select Case swModel.GetType
            
            Case swDocPART:
                MsgBox "Active document is Part"
            
            Case swDocASSEMBLY:
                MsgBox "Active document is Assembly"
                
            Case swDocDRAWING:
                MsgBox "Active document is Drawing"
        End Select
        
    Else
        
        MsgBox "No document opened"
        
    End If
    
End Sub