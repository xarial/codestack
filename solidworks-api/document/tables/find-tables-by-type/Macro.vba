Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim vTables As Variant
        vTables = FindTables(swDraw, Array(swTableAnnotationType_e.swTableAnnotation_BillOfMaterials, swTableAnnotationType_e.swTableAnnotation_RevisionBlock))
        
        If Not IsEmpty(vTables) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vTables)
                
                Dim swTable As SldWorks.TableAnnotation
                Set swTable = vTables(i)
                
                Debug.Print swTable.Title
                
            Next
            
        End If
        
    Else
        MsgBox "Please open drawing"
    End If
    
End Sub

Function FindTables(draw As SldWorks.DrawingDoc, filter As Variant) As Variant
    
    Dim swTables() As SldWorks.TableAnnotation
    Dim isInit As Boolean
    isInit = False
    
    Dim vSheets As Variant
    
    vSheets = draw.GetViews()
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.View
        Set swSheetView = vViews(0)
        
        Dim vTableAnns As Variant
        vTableAnns = swSheetView.GetTableAnnotations
        
        If Not IsEmpty(vTableAnns) Then
            
            Dim j As Integer
            
            For j = 0 To UBound(vTableAnns)
                
                Dim swTableAnn As SldWorks.TableAnnotation
                Set swTableAnn = vTableAnns(j)
                
                If FilterContains(swTableAnn.Type, filter) Then
                    If isInit Then
                        ReDim Preserve swTables(UBound(swTables) + 1)
                    Else
                        ReDim swTables(0)
                        isInit = True
                    End If
                End If
                
                Set swTables(UBound(swTables)) = swTableAnn
                
            Next
            
        End If
        
    Next
    
    FindTables = swTables
    
End Function

Function FilterContains(val As swTableAnnotationType_e, filter As Variant) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(filter)
        If val = filter(i) Then
            FilterContains = True
            Exit Function
        End If
    Next
    
    FilterContains = False
    
End Function