Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swSelMgr As SldWorks.SelectionMgr
        
        Set swSelMgr = swModel.SelectionManager
        
        Dim swNote As SldWorks.Note
        
        Set swNote = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swNote Is Nothing Then
            Dim prpLinkedText As String
            prpLinkedText = swNote.PropertyLinkedText
            SetClipboard prpLinkedText
            Debug.Print prpLinkedText
        Else
            Err.Raise vbError, "", "Select note"
        End If
        
    Else
        Err.Raise vbError, "", "Open the model"
    End If
    
End Sub

Sub SetClipboard(text As String)
    
    Dim vText As Variant
    vText = text
    
    Dim htmlFile As Object
    Set htmlFile = CreateObject("htmlfile")
    
    htmlFile.parentWindow.clipboardData.SetData "text", vText

End Sub