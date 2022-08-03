Const DIMS_TRUE As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim vSheets As Variant
        vSheets = swDraw.GetViews
        
        If Not IsEmpty(vSheets) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vSheets)
            
                Dim vViews As Variant
                vViews = vSheets(i)
                
                Dim j As Integer
                
                For j = 1 To UBound(vViews)
                    Dim swView As SldWorks.View
                    Set swView = vViews(j)
                    
                    swView.ProjectedDimensions = Not DIMS_TRUE
                Next
            
            Next
            
        End If
        
    Else
        Err.Raise vbError, "", "Open drawing"
    End If
    
End Sub