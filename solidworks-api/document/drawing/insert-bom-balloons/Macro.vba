Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    Dim swSheet As SldWorks.Sheet
    
    Dim swView As SldWorks.View
    
    Set swSheet = swDraw.GetCurrentSheet()
    
    Set swView = swSheet.GetViews()(0)
    
    Dim vComps As Variant
    vComps = swView.GetVisibleComponents
    
    Dim i As Integer
    
    Dim swBomBalloonParams As SldWorks.BalloonOptions
    Set swBomBalloonParams = swDraw.Extension.CreateBalloonOptions()
    swBomBalloonParams.UpperTextContent = swBalloonTextItemNumber
    
    For i = 0 To UBound(vComps)
        
        Dim swDrawComp As SldWorks.Component2
        Set swDrawComp = vComps(i)
        Dim vVisEnts As Variant
        
        vVisEnts = swView.GetVisibleEntities2(swDrawComp, swSelectType_e.swSelEDGES)
        
        Dim swEnt As SldWorks.Entity
        Set swEnt = vVisEnts(0)
        
        If False <> swEnt.Select4(False, Nothing) Then
            Dim swNote As SldWorks.Note
            Set swNote = swDraw.Extension.InsertBOMBalloon2(swBomBalloonParams)
        Else
            Err.Raise vbError, "", "Failed to select entity for baloon"
        End If
                
    Next
    
End Sub