Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks

    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    Dim swSheet As SldWorks.Sheet
    
    Set swSheet = swDraw.GetCurrentSheet()
    
    Dim vSheetPrps As Variant
    vSheetPrps = swSheet.GetProperties2()
    
    Dim paperSize As swDwgPaperSizes_e
    paperSize = vSheetPrps(0)
    
    Dim templateIn As swDwgTemplates_e
    templateIn = vSheetPrps(1)
    
    Dim scale1 As Double
    scale1 = vSheetPrps(2)
    
    Dim scale2 As Double
    scale2 = vSheetPrps(3)
    
    Dim firstAngle As Boolean
    firstAngle = vSheetPrps(4)
    
    Dim width As Double
    width = vSheetPrps(5)
    
    Dim height As Double
    height = vSheetPrps(6)
    
    Dim prpViewName As String
    prpViewName = vSheetPrps(7)
    
    Dim template As String
    
    template = swSheet.GetTemplateName()
    
    Dim zoneLeftMargin As Double
    zoneLeftMargin = swSheet.GetZoneMargin(swZoneMargin_e.swZoneLeftMargin)
    
    Dim zoneRightMargin As Double
    zoneRightMargin = swSheet.GetZoneMargin(swZoneMargin_e.swZoneRightMargin)
    
    Dim zoneTopMargin As Double
    zoneTopMargin = swSheet.GetZoneMargin(swZoneMargin_e.swZoneTopMargin)
    
    Dim zoneBottomMargin As Double
    zoneBottomMargin = swSheet.GetZoneMargin(swZoneMargin_e.swZoneBottomMargin)
    
    Dim zoneRows As Long
    Dim zoneCols As Long
    
    swSheet.GetZoneSizeDistribution zoneRows, zoneCols
    
    Dim sheetName As String
    sheetName = swSheet.GetName() & " Copy"
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames()
    
    If False <> swDraw.NewSheet4(sheetName, paperSize, templateIn, scale1, scale2, firstAngle, template, width, height, prpViewName, zoneLeftMargin, zoneRightMargin, zoneTopMargin, zoneBottomMargin, zoneRows, zoneCols) Then
        
        Dim swNewSheet As SldWorks.Sheet
        Set swNewSheet = swDraw.Sheet(sheetName)
        
        Dim orderedSheetNames() As String
        ReDim orderedSheetNames(UBound(vSheetNames) + 1)
        
        Dim i As Integer
        Dim j As Integer
        
        j = -1
        
        For i = 0 To UBound(vSheetNames)
            
            j = j + 1
            
            orderedSheetNames(j) = vSheetNames(i)
            
            If LCase(CStr(vSheetNames(i))) = LCase(swSheet.GetName()) Then
                j = j + 1
                orderedSheetNames(j) = swNewSheet.GetName()
            End If
            
        Next
        
        If False = swDraw.ReorderSheets(orderedSheetNames) Then
            Err.Raise vbError, "", "Failed to reorder sheets"
        End If
        
    End If

End Sub