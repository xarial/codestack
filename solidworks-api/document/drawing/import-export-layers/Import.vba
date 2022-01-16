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
        ImportLayers swDraw, filePath
    Else
        Err.Raise vbError, "", "Open drawing"
    End If
    
End Sub

Sub ImportLayers(draw As SldWorks.DrawingDoc, filePath As String)
    
    Dim swLayerMgr As SldWorks.LayerMgr
    
    Set swLayerMgr = draw.GetLayerManager
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        
        Dim swCurrentLayer As SldWorks.Layer
        
        Set file = fso.OpenTextFile(filePath)
                
        Do Until file.AtEndOfStream
                
            Dim line As String
                
            line = file.ReadLine
            
            Dim value As String
            
            If IsToken(line, TOKEN_LAYER, value) Then
                
                Set swCurrentLayer = swLayerMgr.GetLayer(value)
                
                If swCurrentLayer Is Nothing Then
                    swLayerMgr.AddLayer value, "", RGB(255, 255, 255), swLineStyles_e.swLineCENTER, swLineWeights_e.swLW_CUSTOM
                    Set swCurrentLayer = swLayerMgr.GetLayer(value)
                End If
                
                If swCurrentLayer Is Nothing Then
                    Err.Raise vbError, "", "Failed to access layer " & value
                End If
                
            Else
                
                If swCurrentLayer Is Nothing Then
                    Err.Raise vbError, "", "Current layer is not set"
                End If
                
                If IsToken(line, TOKEN_DESCRIPTION, value) Then
                    swCurrentLayer.Description = value
                ElseIf IsToken(line, TOKEN_COLOR, value) Then
                    Dim vRgb As Variant
                    vRgb = Split(value, " ")
                    swCurrentLayer.Color = RGB(CInt(Trim(CStr(vRgb(0)))), CInt(Trim(CStr(vRgb(1)))), CInt(Trim(CStr(vRgb(2)))))
                ElseIf IsToken(line, TOKEN_PRINTABLE, value) Then
                    swCurrentLayer.Printable = CBool(value)
                ElseIf IsToken(line, TOKEN_STYLE, value) Then
                    swCurrentLayer.Style = CInt(value)
                ElseIf IsToken(line, TOKEN_VISIBLE, value) Then
                    swCurrentLayer.Visible = CBool(value)
                ElseIf IsToken(line, TOKEN_THICKNESS, value) Then
                    swCurrentLayer.Width = CInt(value)
                End If
                
            End If
            
        Loop
        
        file.Close
        
    Else
        Err.Raise vbError, "", "File does not exist"
    End If
    
End Sub

Function IsToken(txt As String, token As String, ByRef value As String) As Boolean
    
    txt = Trim(txt)
    
    If LCase(Left(txt, Len(token))) = LCase(token) Then
        value = Trim(Right(txt, Len(txt) - Len(token)))
        IsToken = True
    Else
        value = ""
        IsToken = False
    End If
    
End Function