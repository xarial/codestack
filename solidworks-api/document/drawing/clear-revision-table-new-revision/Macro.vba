Const ALL_DRAWINGS As Boolean = False

Const INITIAL_REV As String = "001"
Dim COLUMNS As Variant

Dim swApp As SldWorks.SldWorks

Sub main()

    'Fill the revision table columns
    COLUMNS = Array("Sample Zone", "", "Description", "", "Admin")

    Set swApp = Application.SldWorks

    Dim swDraw As SldWorks.DrawingDoc

    If ALL_DRAWINGS Then
        
        Dim vDocs As Variant
        vDocs = swApp.GetDocuments
        
        Dim i As Integer
        
        For i = 0 To UBound(vDocs)
            If TypeOf vDocs(i) Is SldWorks.DrawingDoc Then
                Set swDraw = vDocs(i)
                Debug.Print "Processing " & swDraw.GetTitle()
                ProcessDrawing swDraw
            End If
        Next
    
    Else
            
        Set swDraw = swApp.ActiveDoc
        
        If Not swDraw Is Nothing Then
        
            ProcessDrawing swDraw
                       
        Else
            Err.Raise vbError, "", "Drawing is not open"
        End If

    End If
    
    
End Sub

Sub ProcessDrawing(draw As SldWorks.DrawingDoc)
    
    Dim vSheets As Variant
    vSheets = draw.GetSheetNames
    
    Dim swSheet As SldWorks.sheet
    Set swSheet = draw.sheet(CStr(vSheets(0)))
    
    Dim swRevTable As SldWorks.RevisionTableAnnotation

    Set swRevTable = swSheet.RevisionTable
    
    If Not swRevTable Is Nothing Then
            
        ClearRevisionTable swRevTable
    
        AddRevision swRevTable, INITIAL_REV, COLUMNS
        
        draw.SetSaveFlag
        
    Else
        'NOTE: API will not work with the revision tables on any but first sheet
        Err.Raise vbError, "", "No revision table is found on the first sheet of" & draw.GetTitle
    End If
    
End Sub

Sub ClearRevisionTable(swRevTable As SldWorks.RevisionTableAnnotation)
    
    Dim swTableAnn As SldWorks.TableAnnotation
    
    Set swTableAnn = swRevTable
            
    Dim i As Integer
    
    For i = swTableAnn.RowCount - 1 To 0 Step -1
        
        Dim revId As Long
        revId = swRevTable.GetIdForRowNumber(i)
        
        If revId <> 0 Then
                
            If False = swRevTable.DeleteRevision(revId, True) Then
                Err.Raise vbError, "", "Failed to delete revision"
            End If
        End If
        
    Next
    
End Sub

Function AddRevision(swRevTable As SldWorks.RevisionTableAnnotation, revName As String, rowData As Variant) As Boolean
    
    Dim i As Integer
    Dim swTableAnn As SldWorks.TableAnnotation
    
    Set swTableAnn = swRevTable
    
    Dim revId As Long
    revId = swRevTable.AddRevision(revName)
    
    Dim rowIndex As Integer
    
    rowIndex = swRevTable.GetRowNumberForId(revId)
            
    If rowIndex <> -1 Then
        For i = 0 To UBound(rowData)
                    
            If rowData(i) <> "" Then
                
                swTableAnn.Text(rowIndex, i) = rowData(i)
            
            End If
                    
        Next
        AddRevision = True
    Else
        AddRevision = False
    End If
    
End Function