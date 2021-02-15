Const CUT_LIST_PRPS_TRANSFER As Long = swCutListTransferOptions_e.swCutListTransferOptions_FileProperties
Const OUT_DIR As String = ""

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
    
    Dim i As Integer
    
    For i = 0 To UBound(vBodies)
        
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        
        If False <> swBody.Select2(False, Nothing) Then
            
            Dim outFilePath As String
            outFilePath = GetOutFilePath(swPart, swBody, OUT_DIR)
            
            Dim errs As Long
            Dim warns As Long
            
            If False <> swPart.SaveToFile3(outFilePath, swSaveAsOptions_e.swSaveAsOptions_Silent, CUT_LIST_PRPS_TRANSFER, False, "", errs, warns) Then
                swApp.CloseDoc outFilePath
            Else
                Err.Raise vbError, "", "Failed to save body " & swBody.Name & " to file " & outFilePath & ". Error code: " & errs
            End If
            
        Else
            Err.Raise vbError, "", "Failed to select body " & swBody.Name
        End If
    Next
    
End Sub

Function GetOutFilePath(model As SldWorks.ModelDoc2, body As SldWorks.Body2, outDir As String) As String
    
    If outDir = "" Then
        outDir = model.GetPathName()
        If outDir = "" Then
            Err.Raise vbError, "", "Output directory cannot be composed as file was never saved"
        End If
        
        outDir = Left(outDir, InStrRev(outDir, "\") - 1)
    End If
    
    If Right(outDir, 1) = "\" Then
        outDir = Left(outDir, Len(outDir) - 1)
    End If
    
    GetOutFilePath = ReplaceInvalidPathSymbols(outDir & "\" & body.Name & ".sldprt")
    
End Function

Function ReplaceInvalidPathSymbols(path As String) As String
    
    Const REPLACE_SYMB As String = "_"
    
    Dim res As String
    res = Right(path, Len(path) - Len("X:\"))
    
    Dim drive As String
    drive = Left(path, Len("X:\"))
    
    Dim invalidSymbols As Variant
    invalidSymbols = Array("/", ":", "*", "?", """", "<", ">", "|")
    
    Dim i As Integer
    For i = 0 To UBound(invalidSymbols)
        Dim invalidSymb As String
        invalidSymb = CStr(invalidSymbols(i))
        res = Replace(res, invalidSymb, REPLACE_SYMB)
    Next
    
    ReplaceInvalidPathSymbols = drive + res
    
End Function