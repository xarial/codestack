Const CSV_FILE_PATH As String = "D:\spline-data.csv"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swSkMgr As SldWorks.SketchManager
    Set swSkMgr = swModel.SketchManager
    
    If Not swSkMgr.ActiveSketch Is Nothing Then
        
        Dim vPts As Variant
        vPts = ReadCsvFile(CSV_FILE_PATH, True)
        
        DrawSpline swSkMgr, vPts
        
    Else
        Err.Raise vbError, "", "Please activate sketch"
    End If
    
End Sub

Sub DrawSpline(skMgr As SldWorks.SketchManager, vPoints As Variant)
    
    skMgr.AddToDB = True
    
    Dim dSplinePts() As Double
    ReDim dSplinePts((UBound(vPoints) + 1) * 3 - 1)
    
    Dim i As Integer
    
    For i = 0 To UBound(vPoints)
        
        Dim vPt As Variant
        vPt = vPoints(i)
        
        Dim x As Double
        Dim y As Double
        Dim z As Double
        
        If UBound(vPt) >= 0 Then
            x = vPt(0)
        End If
        
        If UBound(vPt) >= 1 Then
            y = vPt(1)
        End If
        
        If UBound(vPt) >= 2 Then
            z = vPt(2)
        End If
        
        dSplinePts(i * 3) = x
        dSplinePts(i * 3 + 1) = y
        dSplinePts(i * 3 + 2) = z
        
    Next
    
    Dim swSkSegment As SldWorks.SketchSegment
    
    Set swSkSegment = skMgr.CreateSpline2(dSplinePts, False)
    
    If swSkSegment Is Nothing Then
        Err.Raise vbError, "", "Failed to create spline"
    End If
    
    skMgr.AddToDB = False
    
End Sub

Function ReadCsvFile(filePath As String, firstRowHeader As Boolean) As Variant
    
    'rows x columns
    Dim vTable() As Variant
    
    Dim fileName As String
    Dim tableRow As String
    Dim fileNo As Integer

    fileNo = FreeFile
    
    Open filePath For Input As #fileNo
    
    Dim isFirstRow As Boolean
        
    isFirstRow = True
    isTableInit = False
    
    Do While Not EOF(fileNo)
        
        Line Input #fileNo, tableRow
            
        If Not isFirstRow Or Not firstRowHeader Then
            
            Dim vCells As Variant
            vCells = Split(tableRow, ",")
            
            Dim i As Integer
            
            Dim dCells() As Double
            ReDim dCells(UBound(vCells))
            
            For i = 0 To UBound(vCells)
                dCells(i) = CDbl(vCells(i))
            Next
                    
            If (Not vTable) = -1 Then
                ReDim vTable(0)
            Else
                ReDim Preserve vTable(UBound(vTable) + 1)
            End If
                    
            vTable(UBound(vTable)) = dCells
            
        End If
        
        If isFirstRow Then
            isFirstRow = False
        End If
    
    Loop
    
    Close #fileNo
    
    ReadCsvFile = vTable
    
End Function