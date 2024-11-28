Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swSelMgr As SldWorks.SelectionMgr
        
        Set swSelMgr = swModel.SelectionManager
        
        Dim swBody As SldWorks.Body2
        
        Set swBody = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swBody Is Nothing Then
                        
            swBody.Select2 False, Nothing
            
            Dim templatePath As String
            templatePath = swApp.GetDocumentTemplate(swDocumentTypes_e.swDocDRAWING, "", swDwgPaperSizes_e.swDwgPaperA4size, 0, 0)
            
            Dim swDraw As SldWorks.DrawingDoc
            Set swDraw = swApp.NewDocument(templatePath, swDwgPaperSizes_e.swDwgPaperA4size, 0, 0)
            
            Dim swView As SldWorks.View
            
            Set swView = swDraw.CreateFlatPatternViewFromModelView3(swModel.GetPathName(), "", 0, 0, 0, False, False)
            
        Else
            Err.Raise vbError, "", "Body is not selected"
        End If
    
    Else
        Err.Raise vbError, "", "Open part document"
    End If
    
End Sub
