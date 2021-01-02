Const BASE_VIEWS_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Dim scaleMap As Variant
    scaleMap = Array("0-0.1;*;1:1", "0.1-0.2;0.05-0.1;1:2")
    
    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc

try:
    
    On Error GoTo catch
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        RescaleViews swDraw, swDraw.GetCurrentSheet(), scaleMap
        
    Else
        Err.Raise vbError, "", "Please open the drawing document"
    End If
    
    GoTo finally
    
catch:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical
finally:

End Sub

Sub RescaleViews(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, scaleMap As Variant)
    
    Dim vViews As Variant
    vViews = GetSheetViews(draw, sheet)
    
    Dim i As Integer
    
    For i = 0 To UBound(vViews)
        
        Dim swView As SldWorks.view
        Set swView = vViews(i)
        
        Dim width As Double
        Dim height As Double
        GetViewGeometrySize swView, width, height
        
        Debug.Print swView.Name & " : " & width & " x " & height
        
        Dim j As Integer
        
        For j = 0 To UBound(scaleMap)
            
            Dim minWidth As Double
            Dim maxWidth As Double
            Dim minHeight As Double
            Dim maxHeight As Double
            Dim viewScale As Variant
            
            ExtractParameters CStr(scaleMap(j)), minWidth, maxWidth, minHeight, maxHeight, viewScale
            
            If width >= minWidth And width <= maxWidth And height >= minHeight And height <= maxHeight Then
                Debug.Print swView.Name & " matches " & CStr(scaleMap(j))
                If Not BASE_VIEWS_ONLY Or swView.GetBaseView() Is Nothing Then
                    Debug.Print "Setting scale of " & swView.Name & " to " & viewScale(0) & ":" & viewScale(1)
                    swView.ScaleRatio = viewScale
                Else
                    Debug.Print "Skipping " & swView.Name & " view as it is not a base view"
                End If
                
            Else
                Debug.Print swView.Name & " doesn't match " & CStr(scaleMap(j))
            End If
            
        Next
        
    Next
    
    draw.EditRebuild
    
End Sub

Function GetSheetViews(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet) As Variant

    Dim vSheets As Variant
    vSheets = draw.GetViews()
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
    
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        If UCase(swSheetView.Name) = UCase(sheet.GetName()) Then
            
            If UBound(vViews) > 0 Then
                
                Dim swViews() As SldWorks.view
                
                ReDim swViews(UBound(vViews) - 1)
                
                Dim j As Integer
                
                For j = 1 To UBound(vViews)
                    Set swViews(j - 1) = vViews(j)
                Next
                
                GetSheetViews = swViews
                Exit Function
                
            End If
            
        End If
        
    Next
    
End Function

Sub GetViewGeometrySize(view As SldWorks.view, ByRef width As Double, ByRef height As Double)
    
    Dim borderWidth As Double
    borderWidth = GetViewBorderWidth(view)
    
    Dim vOutline As Variant
    vOutline = view.GetOutline()
    
    Dim viewScale As Double
    viewScale = view.ScaleRatio(1) / view.ScaleRatio(0)
    
    width = (vOutline(2) - vOutline(0) - borderWidth * 2) * viewScale
    height = (vOutline(3) - vOutline(1) - borderWidth * 2) * viewScale
    
End Sub

Function GetViewBorderWidth(view As SldWorks.view) As Double
    
    Const VIEW_BORDER_RATIO = 0.02
    
    Dim width As Double
    Dim height As Double
    
    view.sheet.GetSize width, height
    
    Dim minSize As Double
    
    If width < height Then
        minSize = width
    Else
        minSize = height
    End If
    
    GetViewBorderWidth = minSize * VIEW_BORDER_RATIO
    
End Function

Sub ExtractParameters(params As String, ByRef minWidth As Double, ByRef maxWidth As Double, ByRef minHeight As Double, ByRef maxHeight As Double, ByRef viewScale As Variant)

    Dim vParamsData As Variant
    vParamsData = Split(params, ";")
    
    ExtractSizeBounds CStr(vParamsData(0)), minWidth, maxWidth
    ExtractSizeBounds CStr(vParamsData(1)), minHeight, maxHeight
    
    Dim scaleData As Variant
    scaleData = Split(vParamsData(2), ":")
    
    Dim dViewScale(1) As Double
    dViewScale(0) = CDbl(Trim(scaleData(0)))
    dViewScale(1) = CDbl(Trim(scaleData(1)))
    
    viewScale = dViewScale
    
End Sub

Sub ExtractSizeBounds(boundParam As String, ByRef min As Double, ByRef max As Double)
    
    If Trim(boundParam) = "*" Then
        min = 0
        max = 1000000
    Else
        Dim minMax As Variant
        minMax = Split(boundParam, "-")
        min = CDbl(Trim(minMax(0)))
        max = CDbl(Trim(minMax(1)))
    End If
    
End Sub