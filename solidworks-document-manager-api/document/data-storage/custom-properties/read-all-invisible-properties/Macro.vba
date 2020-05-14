Const SW_DM_KEY As String = "Your license key"

Const FILE_PATH As String = "C:\SampleModel.SLDPRT"

Dim swDmClassFactory As SwDocumentMgr.swDmClassFactory
Dim swDmApp As SwDocumentMgr.SwDMApplication

Sub main()

    Set swDmClassFactory = CreateObject("SwDocumentMgr.SwDMClassFactory")
    
    If Not swDmClassFactory Is Nothing Then
        
        Set swDmApp = swDmClassFactory.GetApplication(SW_DM_KEY)
        Dim swDmDoc As SwDocumentMgr.SwDMDocument19
        Set swDmDoc = OpenDocument(FILE_PATH, True)
        
        Dim vPrpNames As Variant
        vPrpNames = swDmDoc.GetInvisibleCustomPropertyNames()
        
        If Not IsEmpty(vPrpNames) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vPrpNames)
            
                Dim prpName As String
                prpName = vPrpNames(i)
                
                Dim prpType As SwDmCustomInfoType
            
                Dim prpTypeName As String
                
                Dim prpVal As String
                prpVal = swDmDoc.GetInvisibleCustomProperty(prpName, prpType)
                
                Select Case prpType
                    Case SwDmCustomInfoType.swDmCustomInfoDate
                        prpTypeName = "Date"
                    Case SwDmCustomInfoType.swDmCustomInfoNumber
                        prpTypeName = "Number"
                    Case SwDmCustomInfoType.swDmCustomInfoText
                        prpTypeName = "Text"
                    Case SwDmCustomInfoType.swDmCustomInfoYesOrNo
                        prpTypeName = "YesNo"
                    Case SwDmCustomInfoType.swDmCustomInfoUnknown
                        prpTypeName = "Unknown"
                End Select
            
                Debug.Print prpName & ": " & prpVal & " [" & prpTypeName & "]"
            Next
            
        End If
        
    Else
        MsgBox "Document Manager SDK is not installed"
    End If
    
End Sub

Function OpenDocument(filePath As String, readOnly As Boolean) As SwDocumentMgr.SwDMDocument19
    
    Dim openErr As SwDmDocumentOpenError
    
    Dim docType As SwDocumentMgr.SwDmDocumentType
    
    Dim ext As String
    ext = LCase(Right(filePath, Len(".SLDXXX")))
    
    Select Case ext
        Case ".sldprt"
            docType = swDmDocumentPart
        Case ".sldasm"
            docType = swDmDocumentAssembly
        Case ".slddrw"
            docType = swDmDocumentDrawing
    End Select
    
    Dim swDmDoc As SwDocumentMgr.SwDMDocument19
    
    Set swDmDoc = swDmApp.GetDocument(filePath, docType, readOnly, openErr)
    
    If swDmDoc Is Nothing Then
        Err.Raise vbError, "", "Failed to open document: " & openErr
    End If
    
    Set OpenDocument = swDmDoc
    
End Function