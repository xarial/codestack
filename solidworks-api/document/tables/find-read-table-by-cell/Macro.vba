Option Compare Text

Const DELIMETER As String = ","

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim tableData As String
        
        Dim swTableAnnotation As SldWorks.TableAnnotation
        Set swTableAnnotation = FindTableByContent(swDraw, "a", 0, 0)
                
        Dim i As Integer
        Dim j As Integer
        
        For i = 0 To swTableAnnotation.RowCount - 1
            
            If i > 0 Then
                tableData = tableData & vbLf
            End If
            
            For j = 0 To swTableAnnotation.ColumnCount - 1
                If j > 0 Then
                    tableData = tableData & DELIMETER
                End If
                tableData = tableData & swTableAnnotation.Text(i, j)
            Next
            
        Next
        
        Debug.Print tableData
        
    Else
        Err.Raise "Open drawing"
    End If
    
End Sub

Function FindTableByContent(draw As SldWorks.DrawingDoc, searchCellVal As String, cellRow As Integer, cellColumn As Integer) As SldWorks.TableAnnotation
    
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
                
                Dim cellVal As String
                cellVal = swTableAnn.Text(cellRow, cellColumn)
                
                If cellVal Like searchCellVal Then
                    Set FindTableByContent = swTableAnn
                    Exit Function
                End If
                
            Next
            
        End If
        
    Next
    
    Err.Raise vbError, "", "Failed to find the table annotation"
    
End Function