#Const ARGS = False 'True to use arguments from Toolbar+ or Batch+ instead of the constant

Const TOKEN_LAYER = "Layer: "
Const TOKEN_DESCRIPTION = "Description: "
Const TOKEN_COLOR = "Color: "
Const TOKEN_PRINTABLE = "Printable: "
Const TOKEN_STYLE = "Style: "
Const TOKEN_VISIBLE = "Visible: "
Const TOKEN_THICKNESS = "Thickness: "

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    Dim filePath As String
    
    #If ARGS Then
                
        Dim macroRunner As Object
        Set macroRunner = CreateObject("CadPlus.MacroRunner.Sw")
        
        Dim param As Object
        Set param = macroRunner.PopParameter(swApp)
        
        Dim vArgs As Variant
        vArgs = param.Get("Args")
        
        filePath = CStr(vArgs(0))
        
    #Else
        filePath = swDraw.GetPathName
        If filePath <> "" Then
            filePath = Left(filePath, InStrRev(filePath, ".") - 1) & "_Layers.txt"
        Else
            Err.Raise vbError, "", "If output file path is not specified file must be saved"
        End If
    #End If
    
    If Not swDraw Is Nothing Then
        ExportLayers swDraw, filePath
    Else
        Err.Raise vbError, "", "Open drawing"
    End If
    
End Sub

Sub ExportLayers(draw As SldWorks.DrawingDoc, filePath As String)
    
    Dim swLayerMgr As SldWorks.LayerMgr
    
    Set swLayerMgr = draw.GetLayerManager
    
    Dim vLayers As Variant
    vLayers = swLayerMgr.GetLayerList

    Dim fileNmb As Integer
    fileNmb = FreeFile
    
    Open filePath For Output As #fileNmb
        
    Dim i As Integer
    
    For i = 0 To UBound(vLayers)
        
        Dim layerName As String
        layerName = CStr(vLayers(i))
        
        Dim swLayer As SldWorks.Layer
        Set swLayer = swLayerMgr.GetLayer(layerName)
        
        Dim RGBHex As String
        RGBHex = Right("000000" & Hex(swLayer.Color), 6)
        
        Print #fileNmb, TOKEN_LAYER & swLayer.Name
        Print #fileNmb, "    " & TOKEN_DESCRIPTION & swLayer.Description
        Print #fileNmb, "    " & TOKEN_COLOR & CInt("&H" & Mid(RGBHex, 5, 2)) & " " & CInt("&H" & Mid(RGBHex, 3, 2)) & " " & CInt("&H" & Mid(RGBHex, 1, 2))
        Print #fileNmb, "    " & TOKEN_PRINTABLE & swLayer.Printable
        Print #fileNmb, "    " & TOKEN_STYLE & swLayer.Style
        Print #fileNmb, "    " & TOKEN_VISIBLE & swLayer.Visible
        Print #fileNmb, "    " & TOKEN_THICKNESS & swLayer.Width
        Print #fileNmb, ""
        
    Next
        
    Close #fileNmb
    
End Sub