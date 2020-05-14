Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
        
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swSketch As SldWorks.sketch
        
        Set swSketch = swModel.SketchManager.ActiveSketch
        
        If Not swSketch Is Nothing Then
            
            Dim vPoints As Variant
            Dim inputFile As String
            inputFile = InputBox("Specify the full path to CSV file")
            
            If inputFile <> "" Then
            
                vPoints = ReadCsvFile(inputFile, True)
                DrawPoints swModel, vPoints
            
            End If
            
        Else
            MsgBox "Please open 2D or 3D Sketch"
        End If
        
    Else
        MsgBox "Please open the model"
    End If
        
End Sub

Sub DrawPoints(model As SldWorks.ModelDoc2, vPoints As Variant)
    
    model.SketchManager.AddToDB = True
    
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
        
        model.SketchManager.CreatePoint x, y, z
        
    Next
    
    model.SketchManager.AddToDB = False
    
End Sub

Function ReadCsvFile(filePath As String, firstRowHeader As Boolean) As Variant
    
    'rows x columns
    Dim vTable() As Variant
    
    On Error GoTo Error
    
    Dim fileName As String
    Dim tableRow As String
    Dim fileNo As Integer

    fileNo = FreeFile
    
    Open filePath For Input As #fileNo
    
    Dim isFirstRow As Boolean
    Dim isTableInit As Boolean
    
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
            
            Dim lastRowIndex As Integer
            
            If Not isTableInit Then
                lastRowIndex = 0
                isTableInit = True
                ReDim Preserve vTable(lastRowIndex)
            Else
                lastRowIndex = UBound(vTable, 1) + 1
                ReDim Preserve vTable(lastRowIndex)
            End If
            
            vTable(lastRowIndex) = dCells
            
        End If
        
        If isFirstRow Then
            isFirstRow = False
        End If
    
    Loop
    
    Close #fileNo
    
    ReadCsvFile = vTable
    
    Exit Function
    
Error:

    ReadCsvFile = Empty
    
End Function