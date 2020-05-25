Const SW_DM_KEY As String = "<Your License Key>"

Sub main()
End Sub

Function ConnectToDm() As SwDocumentMgr.SwDMApplication

    Dim swDmClassFactory As SwDocumentMgr.swDmClassFactory
    Dim swDmApp As SwDocumentMgr.SwDMApplication
    
    Set swDmClassFactory = CreateObject("SwDocumentMgr.SwDMClassFactory")
        
    If Not swDmClassFactory Is Nothing Then
        Set swDmApp = swDmClassFactory.GetApplication(SW_DM_KEY)
        Set ConnectToDm = swDmApp
    Else
        Err.Raise vbError, "", "Document Manager SDK is not installed"
    End If
    
End Function

Function OpenDocument(swDmApp As SwDocumentMgr.SwDMApplication, path As String, readOnly As Boolean) As SwDocumentMgr.SwDMDocument10
    
    Dim ext As String
    ext = LCase(Right(path, Len(path) - InStrRev(path, ".")))
    
    Dim docType As SwDmDocumentType
    
    Select Case ext
        Case "sldprt"
            docType = swDmDocumentPart
        Case "sldasm"
            docType = swDmDocumentAssembly
        Case "slddrw"
            docType = swDmDocumentDrawing
        Case Else
            Err.Raise vbError, "", "Unsupported file type: " & ext
    End Select
    
    Dim swDmDoc As SwDocumentMgr.SwDMDocument10
    Dim openDocErr As SwDmDocumentOpenError
    Set swDmDoc = swDmApp.GetDocument(path, docType, readOnly, openDocErr)
    
    If swDmDoc Is Nothing Then
        Err.Raise vbError, "", "Failed to open document: '" & path & "'. Error Code: " & openDocErr
    End If
    
    Set OpenDocument = swDmDoc
    
End Function

Public Function GETSWPRP(fileName As String, prpNames As Variant, Optional confName As String = "") As Variant
    
    Dim swDmApp As SwDocumentMgr.SwDMApplication
    Dim swDmDoc As SwDocumentMgr.SwDMDocument10
    
try_:
    On Error GoTo catch_
    
    Dim vNames As Variant
            
    If TypeName(prpNames) = "Range" Then
        vNames = RangeToArray(prpNames)
    Else
        vNames = Array(CStr(prpNames))
    End If
    
    Set swDmApp = ConnectToDm()
    Set swDmDoc = OpenDocument(swDmApp, fileName, True)
    
    Dim res() As String
    Dim i As Integer
    ReDim res(UBound(vNames))
    
    Dim prpType As SwDmCustomInfoType
    
    If confName = "" Then
        For i = 0 To UBound(vNames)
            res(i) = swDmDoc.GetCustomProperty(CStr(vNames(i)), prpType)
        Next
    Else
        Dim swDmConf As SwDocumentMgr.SwDMConfiguration10
        Set swDmConf = swDmDoc.ConfigurationManager.GetConfigurationByName(confName)
        If Not swDmConf Is Nothing Then
            For i = 0 To UBound(vNames)
                res(i) = swDmConf.GetCustomProperty(CStr(vNames(i)), prpType)
            Next
        Else
            Err.Raise vbError, "", "Failed to get configuration '" & confName & "' from '" & fileName & "'"
        End If
    End If
    
    GETSWPRP = res
    
    GoTo finally_
    
catch_:
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
finally_:
    If Not swDmDoc Is Nothing Then
        swDmDoc.CloseDoc
    End If

End Function

Public Function SETSWPRP(fileName As String, prpNames As Variant, prpVals As Variant, Optional confName As String = "")
    
    Dim swDmApp As SwDocumentMgr.SwDMApplication
    Dim swDmDoc As SwDocumentMgr.SwDMDocument10
    
try_:
    On Error GoTo catch_
    
    If TypeName(prpNames) <> TypeName(prpVals) Then
        Err.Raise vbError, "", "Property name and value must be of the same type, e.g. either range or cell"
    End If
    
    Dim vNames As Variant
    Dim vVals As Variant
        
    If TypeName(prpNames) = "Range" Then
        
        vNames = RangeToArray(prpNames)
        
        vVals = RangeToArray(prpVals)
        
        If UBound(vNames) <> UBound(vVals) Then
            Err.Raise vbError, "", "Number of cells in the name and value are not equal"
        End If
    Else
        vNames = Array(CStr(prpNames))
        vVals = Array(CStr(prpVals))
    End If
    
    Set swDmApp = ConnectToDm()
    Set swDmDoc = OpenDocument(swDmApp, fileName, False)
    
    Dim i As Integer
    
    If confName = "" Then
        For i = 0 To UBound(vNames)
            swDmDoc.AddCustomProperty CStr(vNames(i)), swDmCustomInfoText, CStr(vVals(i))
            swDmDoc.SetCustomProperty CStr(vNames(i)), CStr(vVals(i))
        Next
    Else
        Dim swDmConf As SwDocumentMgr.SwDMConfiguration10
        Set swDmConf = swDmDoc.ConfigurationManager.GetConfigurationByName(confName)
        
        If Not swDmConf Is Nothing Then
            For i = 0 To UBound(vNames)
                swDmConf.AddCustomProperty CStr(vNames(i)), swDmCustomInfoText, CStr(vVals(i))
                swDmConf.SetCustomProperty CStr(vNames(i)), CStr(vVals(i))
            Next
        Else
            Err.Raise vbError, "", "Failed to get configuration '" & confName & "' from '" & fileName & "'"
        End If
    End If
    
    swDmDoc.Save
    
    SETSWPRP = "OK"
    
    GoTo finally_
    
catch_:
    Debug.Print Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
finally_:
    If Not swDmDoc Is Nothing Then
        swDmDoc.CloseDoc
    End If
    
End Function

Private Function RangeToArray(vRange As Variant) As Variant
    
    If TypeName(vRange) = "Range" Then
        Dim excelRange As range
        Set excelRange = vRange
        
        Dim i As Integer
        
        Dim valsArr() As String
        ReDim valsArr(excelRange.Cells.Count - 1)
        
        i = 0
        
        For Each cell In excelRange.Cells
            valsArr(i) = cell.Value
            i = i + 1
        Next
        
        RangeToArray = valsArr
        
    Else
        Err.Raise vbError, "", "Value is not a Range"
    End If
    
End Function