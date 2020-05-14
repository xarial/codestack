Const FILE_PATH As String = "FILE PATH"

Const LIC_KEY As String = "YOUR LICENSE KEY"

Dim swDmApp As SwDocumentMgr.SwDMApplication4

Sub main()

    Dim swClassFact As SwDocumentMgr.SwDMClassFactory
    
    Set swClassFact = New SwDocumentMgr.SwDMClassFactory
    
    Set swDmApp = swClassFact.GetApplication(LIC_KEY)
    
    Dim filesColl As Collection
    Set filesColl = New Collection
    
    CollectExternalReferences FILE_PATH, filesColl
    
    Dim i As Integer
    
    Debug.Print "External References:"
    
    For i = 1 To filesColl.Count
        Debug.Print filesColl(i)
    Next
    
End Sub

Function CollectExternalReferences(filePath As String, coll As Collection)
    
    If Not Contains(coll, filePath) Then
        coll.Add filePath
    End If
    
    Dim swDmDoc As SwDocumentMgr.SwDMDocument19
    
    Dim searchOpts As SwDocumentMgr.SwDMSearchOption
    Set searchOpts = swDmApp.GetSearchOptionObject
    searchOpts.SearchFilters = SwDmSearchFilters.SwDmSearchExternalReference + SwDmSearchFilters.SwDmSearchRootAssemblyFolder + SwDmSearchFilters.SwDmSearchSubfolders + SwDmSearchFilters.SwDmSearchInContextReference
    
    Set swDmDoc = OpenDocument(filePath)
    
    If Not swDmDoc Is Nothing Then
        
        Dim vBrokenRefs As Variant
        Dim vVirtComps As Variant
        Dim vTimeStamps As Variant
        Dim vFilePaths As Variant
        
        vFilePaths = swDmDoc.GetAllExternalReferences4(searchOpts, vBrokenRefs, vVirtComps, vTimeStamps)
        
        If Not IsEmpty(vFilePaths) Then
            Dim i As Integer
            
            For i = 0 To UBound(vFilePaths)
                Dim childFilePath As String
                childFilePath = vFilePaths(i)
                CollectExternalReferences childFilePath, coll
            Next
            
        End If
        
    Else
        Debug.Print "Failed to open document: " & filePath
    End If
    
End Function

Function OpenDocument(filePath As String) As SwDocumentMgr.SwDMDocument19
    
    Dim err As SwDmDocumentOpenError
    
    Dim docType As SwDocumentMgr.SwDmDocumentType
    
    Dim ext As String
    ext = LCase(Right(filePath, 6))
    
    Select Case ext
        Case "sldprt"
            docType = swDmDocumentPart
        Case "sldasm"
            docType = swDmDocumentAssembly
        Case "slddrw"
            docType = swDmDocumentDrawing
    End Select
    
    Dim swDmDoc As SwDocumentMgr.SwDMDocument19
    
    Set swDmDoc = swDmApp.GetDocument(filePath, docType, True, err)
    
    Set OpenDocument = swDmDoc
    
End Function

Function Contains(coll As Collection, item As String) As Boolean
    
    Dim i As Integer
    
    For i = 1 To coll.Count
        If LCase(coll.item(i)) = LCase(item) Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function