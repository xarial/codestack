Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
        
    Set swDraw = swApp.ActiveDoc
    
    Dim swView As SldWorks.view
    Set swView = swDraw.FeatureByName("Drawing View1").GetSpecificFeature()
    
    Dim swEdge As SldWorks.edge
    Set swEdge = FindEdge(swDraw, swView, "Part1-1", "MyEdge")
    
    Debug.Print swView.SelectEntity(swEdge, False)
    
End Sub

Function FindEdge(draw As SldWorks.DrawingDoc, view As SldWorks.view, compName As String, edgeName As String) As SldWorks.edge
    
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = view.ReferencedDocument
    
    Dim swComp As SldWorks.Component2
    Set swComp = swAssy.GetComponentByName(compName)
    
    Dim swRefPart As SldWorks.PartDoc
    Set swRefPart = swComp.GetModelDoc2
    
    Dim swEdge As SldWorks.edge
    Set swEdge = swRefPart.GetEntityByName(edgeName, swSelectType_e.swSelEDGES)
    
    Set swEdge = swComp.GetCorresponding(swEdge)
    
    Set FindEdge = swEdge
    
End Function