Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swTableAnnotation As SldWorks.TableAnnotation

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then

        Set swSelMgr = swModel.SelectionManager
        
        Dim tableData() As String
        
        Set swTableAnnotation = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swTableAnnotation Is Nothing Then
            
            ReDim tableData(swTableAnnotation.RowCount - 1, swTableAnnotation.ColumnCount - 1)
            
            Dim i As Integer
            Dim j As Integer
            
            For i = 0 To swTableAnnotation.RowCount - 1
                
                For j = 0 To swTableAnnotation.ColumnCount - 1
                    tableData(i, j) = swTableAnnotation.Text(i, j)
                Next
                
            Next
        Else
            MsgBox "Please select table"
        End If
    Else
        MsgBox "Please open model"
    End If
End Sub
