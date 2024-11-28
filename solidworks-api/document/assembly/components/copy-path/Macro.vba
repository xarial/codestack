Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

try_:
    On Error GoTo catch_

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    Dim path As String
    
    If Not swModel Is Nothing Then
        
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        Dim i As Integer
        
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
            Dim swComp As SldWorks.Component2
            Set swComp = Nothing
            
            If TypeOf swModel Is SldWorks.AssemblyDoc Then
                
                Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
                
            ElseIf TypeOf swModel Is SldWorks.DrawingDoc Then
                
                Dim swDrawComp As SldWorks.DrawingComponent
                Set swDrawComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
                
                If Not swDrawComp Is Nothing Then
                    Set swComp = swDrawComp.Component
                End If
                
            Else
                Err.Raise vbError, "", "Only parts and drawings are supported"
            End If
            
            If Not swComp Is Nothing Then
                If path <> "" Then
                    path = path & vbLf
                End If
                path = path & swComp.GetPathName
            End If
            
        Next
        
        If path <> "" Then
            Debug.Print path
            SetTextToClipboard path
        Else
            Err.Raise vbError, "", "Please select components"
        End If
        
    Else
        Err.Raise vbError, "", "Please open document"
    End If
    
    GoTo finally_
    
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Sub SetTextToClipboard(text As String)
    
    Dim dataObject As Object
    Set dataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObject.SetText text
    dataObject.PutInClipboard
    Set dataObject = Nothing
    
End Sub