Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByRef lpOutput As Any, ByRef lpDevMode As Any) As Long

Const PRINTER_NAME As String = "Microsoft Print To PDF"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
        
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    Dim outFilePath As String
    
    outFilePath = swModel.GetPathName()
    
    If outFilePath = "" Then
        Err.Raise vbError, "", "Drawing is not saved"
    End If
    
    outFilePath = Left(outFilePath, Len(outFilePath) - Len(".slddrw")) & ".pdf"
    
    Dim swPageSetup As SldWorks.PageSetup
    
    Set swPageSetup = swModel.PageSetup

    Dim origUsePageSetup As Integer
    Dim origPrinterPaperSize As Integer
    Dim origOrientation As Integer
    
    origUsePageSetup = swModel.Extension.UsePageSetup
    origPrinterPaperSize = swPageSetup.PrinterPaperSize
    origOrientation = swPageSetup.orientation

    Dim paperSize As String
    Dim orientation As swPageSetupOrientation_e

    GetSize swDraw.GetCurrentSheet(), paperSize, orientation

    swPageSetup.PrinterPaperSize = GetPaper(PRINTER_NAME, paperSize)
    swPageSetup.orientation = orientation

    swModel.Extension.UsePageSetup = swPageSetupInUse_e.swPageSetupInUse_Document

    Dim swPrintSpec As SldWorks.PrintSpecification
    
    Set swPrintSpec = swModel.Extension.GetPrintSpecification
    swPrintSpec.ConvertToHighQuality = True
    swPrintSpec.PrintToFile = True
    swPrintSpec.ScaleMethod = swPrintSelectionScaleFactor_e.swPrintAll
    swPrintSpec.AddPrintRange 1, swDraw.GetSheetCount()

    swModel.Extension.PrintOut4 PRINTER_NAME, outFilePath, swPrintSpec

    swPrintSpec.RestoreDefaults
    swPrintSpec.ResetPrintRange

    swModel.Extension.UsePageSetup = origUsePageSetup
    swPageSetup.PrinterPaperSize = origPrinterPaperSize
    swPageSetup.orientation = origOrientation
          
End Sub

Sub GetSize(sheet As SldWorks.sheet, ByRef paperSize As String, ByRef orientation As swPageSetupOrientation_e)

    Select Case sheet.GetSize(0, 0)
        Case swDwgPaperA0size:
            paperSize = "A0"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperA1size:
            paperSize = "A1"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperA2size:
            paperSize = "A2"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperA3size:
            paperSize = "A3"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperA4size:
            paperSize = "A4"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperA4sizeVertical:
            paperSize = "A4"
            orientation = swPageSetupOrient_Portrait
        Case swDwgPaperAsize:
            paperSize = "A"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperAsizeVertical:
            paperSize = "A"
            orientation = swPageSetupOrient_Portrait
        Case swDwgPaperBsize:
            paperSize = "B"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperCsize:
            paperSize = "C"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperDsize:
            paperSize = "D"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPaperEsize:
            paperSize = "E"
            orientation = swPageSetupOrient_Landscape
        Case swDwgPapersUserDefined:
            Err.Raise vbError, "", "Size is not supported"
    End Select
    
End Sub

Function GetPaper(printerName As String, paperName As String) As Integer
    
    Const DC_PAPERNAMES As Integer = &H10
    Const DC_PAPERS As Integer = &H2
    
    Dim papersCount As Integer
    papersCount = DeviceCapabilities(printerName, "", DC_PAPERS, ByVal vbNullString, 0)
    
    If papersCount > 0 Then
    
        Dim papersCodes() As Integer
        ReDim papersCodes(papersCount - 1)
        
        DeviceCapabilities printerName, "", DC_PAPERS, papersCodes(0), 0
        
        Dim papersNames As String
        papersNames = String$(64 * papersCount, 0)
        DeviceCapabilities printerName, "", DC_PAPERNAMES, ByVal papersNames, 0
      
        Dim i As Integer
        
        For i = 0 To papersCount
            If LCase(ParsePaperName(papersNames, 64 * i + 1)) = LCase(paperName) Then
                GetPaper = papersCodes(i)
            End If
        Next
    Else
        Err.Raise vbError, "", "No sizes available for the specified printer"
    End If
    
End Function

Function ParsePaperName(papersNames As String, offset As Integer) As String

    Dim paperName As String
    
    paperName = Mid(papersNames, offset, 64)
    
    Dim nullCharIndex As Integer
    nullCharIndex = InStr(paperName, vbNullChar)
    
    If nullCharIndex <> 0 Then
        paperName = Left$(paperName, nullCharIndex - 1)
    End If
     
    ParsePaperName = paperName
    
End Function