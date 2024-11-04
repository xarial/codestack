Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swBody1 As SldWorks.Body2
    Dim swBody2 As SldWorks.Body2
    
    Set swBody1 = swSelMgr.GetSelectedObject6(1, -1)
    Set swBody2 = swSelMgr.GetSelectedObject6(2, -1)
    
    If Not swBody1 Is Nothing And Not swBody2 Is Nothing Then
    
        Set swBody1 = swBody1.Copy2(False)
        Set swBody2 = swBody2.Copy2(False)
            
        Dim swModeler As SldWorks.Modeler
        
        Set swModeler = swApp.GetModeler
        
        Dim vBody1Faces As Variant
        Dim vBody2Faces As Variant
        
        Dim vIntersectBodies As Variant
        
        If False <> swModeler.CheckInterferenceBetweenTwoBodies(swBody1, swBody2, True, vBody1Faces, vBody2Faces, vIntersectBodies) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vIntersectBodies)
                Dim swIntersectBody As SldWorks.Body2
                Set swIntersectBody = vIntersectBodies(i)
                swIntersectBody.Display3 swModel, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
            Next
            
            Stop
            
            For i = 0 To UBound(vIntersectBodies)
                Set vIntersectBodies(i) = Nothing
            Next
        Else
            Debug.Print "No Interferences"
        End If
        
    Else
        Err.Raise vbError, "", "Select 2 bodies"
    End If
    
End Sub