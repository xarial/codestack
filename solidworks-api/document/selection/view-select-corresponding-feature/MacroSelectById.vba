Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swRefModel As SldWorks.ModelDoc2
    
    Set swRefModel = swApp.ActiveDoc
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = swRefModel.SelectionManager.GetSelectedObject6(1, -1)
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swRefModel.SelectionManager
    
    Dim selName As String
    Dim selType As String
    selName = swFeat.GetNameForSelection(selType)
    
    Stop
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swApp.ActiveDoc
    
    Dim swView As SldWorks.View
    Set swView = swDraw.SelectionManager.GetSelectedObject6(1, -1)
    
    Dim drwSelPrefix As String
    drwSelPrefix = swFeat.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
    
    selName = Right(selName, Len(selName) - InStr(selName, "@"))
    
    If False = swDraw.Extension.SelectByID2(drwSelPrefix & "/" & selName, selType, 0, 0, 0, False, 0, Nothing, 0) Then
        Err.Raise vbError, "", "Failed to select corresponding feature in the drawing view"
    End If

End Sub