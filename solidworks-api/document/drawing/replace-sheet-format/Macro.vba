Const REMOVE_MODIFIED_NOTES As Boolean = True
Const FILTER_ANY As String = "*"

Dim swApp As SldWorks.SldWorks

Dim REPLACE_MAP As Variant

Sub main()

    REPLACE_MAP = Array("*|*|D:\new-format.slddrt")

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames
    
    Dim i As Integer
    
    Dim activeSheet As String
    activeSheet = swDraw.GetCurrentSheet().GetName
    
    For i = 0 To UBound(vSheetNames)
        
        Dim sheetName As String
        sheetName = CStr(vSheetNames(i))
        
        Dim swSheet As SldWorks.sheet
        Set swSheet = swDraw.sheet(sheetName)
        
        Dim targetSheetFormatFileName As String
        targetSheetFormatFileName = GetReplaceSheetFormat(swSheet)
        
        swDraw.ActivateSheet sheetName
        
        ReplaceSheetFormat swDraw, swSheet, targetSheetFormatFileName

    Next
    
    swDraw.ActivateSheet activeSheet
    
End Sub

Function GetReplaceSheetFormat(sheet As SldWorks.sheet) As String
    
    Dim curTemplateName As String
    curTemplateName = sheet.GetTemplateName()
    
    Dim curSize As Integer
    curSize = sheet.GetSize(-1, -1)
    
    Dim i As Integer
    
    For i = 0 To UBound(REPLACE_MAP)
        
        Dim map As String
        map = REPLACE_MAP(i)
        
        Dim mapParams As Variant
        mapParams = Split(map, "|")
        
        Dim mapPaperSize As Integer
        Dim srcTemplateName As String
        
        If Trim(mapParams(0)) <> FILTER_ANY Then
            mapPaperSize = CInt(Trim(mapParams(0)))
        Else
            mapPaperSize = -1
        End If
        
        If Trim(mapParams(1)) <> FILTER_ANY Then
            srcTemplateName = CStr(Trim(mapParams(1)))
        Else
            srcTemplateName = ""
        End If
        
        If (mapPaperSize = -1 Or mapPaperSize = curSize) And (srcTemplateName = "" Or LCase(srcTemplateName) = LCase(curTemplateName)) Then
            
            Dim targetTemplateName As String

            targetTemplateName = CStr(Trim(mapParams(2)))
        
            If targetTemplateName = "" Then
                Err.Raise vbError, "", "Target template is not specified"
            End If
        
            GetReplaceSheetFormat = targetTemplateName
            Exit Function
            
        End If
        
    Next
    
    Err.Raise vbError, "", "Failed find the sheet format mathing current sheet"
    
End Function

Sub ReplaceSheetFormat(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, targetSheetFormatFile As String)
    
    Debug.Print "Replacing '" & sheet.GetName() & "' with '" & targetSheetFormatFile & "'"
    
    Dim vProps As Variant
    vProps = sheet.GetProperties()
    
    Dim paperSize As Integer
    Dim templateType As Integer
    Dim scale1 As Double
    Dim scale2 As Double
    Dim firstAngle As Boolean
    Dim width As Double
    Dim height As Double
    Dim custPrpView As String
    
    paperSize = CInt(vProps(0))
    templateType = CInt(vProps(1))
    scale1 = CDbl(vProps(2))
    scale2 = CDbl(vProps(3))
    firstAngle = CBool(vProps(4))
    width = CDbl(vProps(5))
    height = CDbl(vProps(6))
    custPrpView = sheet.CustomPropertyView
    
    If False = draw.SetupSheet5(sheet.GetName(), paperSize, templateType, scale1, scale2, firstAngle, targetSheetFormatFile, width, height, custPrpView, REMOVE_MODIFIED_NOTES) Then
        Err.Raise vbError, "", "Failed to set the sheet format"
    End If
    
End Sub