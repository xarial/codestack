Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
try_:
    On Error GoTo catch_:
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If swDraw Is Nothing Then
        Err.Raise vbError, "", "Please open drawing"
    End If
    
    ExportDrawingDimensions swDraw
    
    GoTo finally_

catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Sub ExportDrawingDimensions(draw As SldWorks.DrawingDoc)
    
    Dim vSheets As Variant
    vSheets = draw.GetViews

    Dim fileNmb As Integer
    fileNmb = FreeFile
    
    Dim filePath As String
    filePath = draw.GetPathName
    
    If filePath = "" Then
        Err.Raise vbError, "", "Please save drawing document"
    End If
    
    filePath = Left(filePath, InStrRev(filePath, ".") - 1) & "-dimensions.csv"
    
    Open filePath For Output As #fileNmb
    
    Dim header As String
    header = Join("Name", "Owner", "Type", "X", "Y", "Value", "Grid Ref", "Tolerance", "Min", "Max")

    Print #fileNmb, header
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim j As Integer
        
        For j = 0 To UBound(vViews)
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            ExportViewDimensions swView, draw, fileNmb
            
        Next
        
    Next

    Close #fileNmb
    
End Sub

Sub ExportViewDimensions(view As SldWorks.view, draw As SldWorks.DrawingDoc, fileNmb As Integer)
    
    Dim swDispDim As SldWorks.DisplayDimension
    Set swDispDim = view.GetFirstDisplayDimension5
    
    Dim swSheet As SldWorks.Sheet
    
    Set swSheet = view.Sheet
    
    If swSheet Is Nothing Then
        Set swSheet = draw.Sheet(view.name)
    End If
    
    While Not swDispDim Is Nothing
        
        Dim swAnn As SldWorks.Annotation
        Set swAnn = swDispDim.GetAnnotation
        
        Dim vPos As Variant
        vPos = swAnn.GetPosition()
        
        Dim swDim As SldWorks.dimension
        Set swDim = swDispDim.GetDimension2(0)
                
        Dim drwZone As String
        drwZone = swSheet.GetDrawingZone(vPos(0), vPos(1))
        vPos = GetPositionInDrawingUnits(vPos, draw)
        
        Dim tolType As String
        Dim minVal As Double
        Dim maxVal As Double
        
        GetDimensionTolerance draw, swDim, tolType, minVal, maxVal
        
        OutputDimensionData fileNmb, swDim.FullName, view.name, GetDimensionType(swDispDim), CDbl(vPos(0)), CDbl(vPos(1)), _
                CDbl(swDim.GetValue3(swInConfigurationOpts_e.swThisConfiguration, Empty)(0)), _
                drwZone, tolType, minVal, maxVal
        
        Set swDispDim = swDispDim.GetNext5
        
    Wend
    
End Sub

Function GetPositionInDrawingUnits(pos As Variant, draw As SldWorks.DrawingDoc) As Variant
    
    Dim dPt(1) As Double
    dPt(0) = ConvertToUserUnits(draw, CDbl(pos(0)), swLengthUnit)
    dPt(1) = ConvertToUserUnits(draw, CDbl(pos(1)), swLengthUnit)
    
    GetPositionInDrawingUnits = dPt
    
End Function

Function ConvertToUserUnits(model As SldWorks.ModelDoc2, val As Double, unitType As swUserUnitsType_e) As Double
    
    Dim swUserUnit As SldWorks.UserUnit
    Set swUserUnit = model.GetUserUnit(unitType)
    
    Dim convFactor As Double
    convFactor = swUserUnit.GetConversionFactor()
    
    ConvertToUserUnits = val * convFactor
    
End Function


Function GetDimensionType(dispDim As SldWorks.DisplayDimension) As String

    Select Case dispDim.Type2
        Case swDimensionType_e.swAngularDimension
            GetDimensionType = "Angular"
        Case swDimensionType_e.swArcLengthDimension
            GetDimensionType = "ArcLength"
        Case swDimensionType_e.swChamferDimension
            GetDimensionType = "Chamfer"
        Case swDimensionType_e.swDiameterDimension
            GetDimensionType = "Diameter"
        Case swDimensionType_e.swDimensionTypeUnknown
            GetDimensionType = "Unknown"
        Case swDimensionType_e.swHorLinearDimension
            GetDimensionType = "HorLinear"
        Case swDimensionType_e.swHorOrdinateDimension
            GetDimensionType = "HorOrdinate"
        Case swDimensionType_e.swLinearDimension
            GetDimensionType = "Linear"
        Case swDimensionType_e.swOrdinateDimension
            GetDimensionType = "Ordinate"
        Case swDimensionType_e.swRadialDimension
            GetDimensionType = "Radial"
        Case swDimensionType_e.swScalarDimension
            GetDimensionType = "Scalar"
        Case swDimensionType_e.swVertLinearDimension
            GetDimensionType = "VertLinear"
        Case swDimensionType_e.swVertOrdinateDimension
            GetDimensionType = "VertOrdinate"
        Case swDimensionType_e.swZAxisDimension
            GetDimensionType = "ZAxis"
    End Select
    
End Function

Sub GetDimensionTolerance(draw As SldWorks.DrawingDoc, swDim As SldWorks.dimension, ByRef tolType As String, ByRef minVal As Double, ByRef maxVal As Double)

    Dim swTol As SldWorks.DimensionTolerance
    Set swTol = swDim.Tolerance
    
    Select Case swTol.Type
        Case swTolType_e.swTolBASIC
            tolType = "Basic"
        Case swTolType_e.swTolBILAT
            tolType = "Bilat"
        Case swTolType_e.swTolBLOCK
            tolType = "Block"
        Case swTolType_e.swTolFIT
            tolType = "Fit"
        Case swTolType_e.swTolFITTOLONLY
            tolType = "FitTolOnly"
        Case swTolType_e.swTolFITWITHTOL
            tolType = "FitWithTol"
        Case swTolType_e.swTolGeneral
            tolType = "General"
        Case swTolType_e.swTolLIMIT
            tolType = "Limit"
        Case swTolType_e.swTolMAX
            tolType = "Max"
        Case swTolType_e.swTolMETRIC
            tolType = "Metric"
        Case swTolType_e.swTolMIN
            tolType = "Min"
        Case swTolType_e.swTolNONE
            tolType = "None"
        Case swTolType_e.swTolSYMMETRIC
            tolType = "Symmetric"
    End Select

    swTol.GetMinValue2 minVal
    swTol.GetMaxValue2 maxVal
    
    Dim unitType As swUserUnitsType_e
    
    If swDim.GetType() = swDimensionParamType_e.swDimensionParamTypeDoubleAngular Then
        unitType = swUserUnitsType_e.swAngleUnit
    Else
        unitType = swUserUnitsType_e.swLengthUnit
    End If
    
    minVal = ConvertToUserUnits(draw, minVal, unitType)
    maxVal = ConvertToUserUnits(draw, maxVal, unitType)
    
End Sub

Sub OutputDimensionData(fileNmb As Integer, dimName As String, owner As String, dimType As String, x As Double, y As Double, value As Double, gridRef As String, tol As String, min As Double, max As Double)
    
    Dim line As String
    line = Join(dimName, owner, dimType, x, y, value, gridRef, tol, min, max)

    Print #fileNmb, line
    
End Sub

Function Join(ParamArray parts() As Variant) As String
    
    Dim res As String
    
    If Not IsEmpty(parts) Then
        Dim i As Integer
        For i = 0 To UBound(parts)
            res = res & IIf(i = 0, "", ", ") & parts(i)
        Next
    End If
    
    Join = res
    
End Function