'#Const TEST = True

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
       
    Dim macroOper As IMacroOperation
    Set macroOper = GetMacroOperation()
    
    Dim vArgs As Variant
    vArgs = macroOper.Arguments
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = macroOper.model
    
    If Not swModel Is Nothing Then
    
        If swModel.GetType() = swDocumentTypes_e.swDocPART Then
        
            Dim swPart As SldWorks.PartDoc
            
            Set swPart = swModel
            
            Dim vCutLists As Variant
            vCutLists = GetCutLists(swPart)
            
            Dim i As Integer
            Dim swBody As SldWorks.Body2
            
            Dim customVarValProv As IMacroCustomVariableValueProvider
            Set customVarValProv = New CustomVariableValueProvider
            
            Dim resFilePaths() As String
            Dim inputBodies() As SldWorks.Body2
            
            For i = 0 To UBound(vCutLists)
            
                Dim swCutList As SldWorks.Feature
                Set swCutList = vCutLists(i)
                
                Dim j As Integer
                
                For j = 0 To UBound(vArgs)
                
                    Dim macroArg As IMacroArgument
                    Set macroArg = vArgs(j)
                    
                    Dim fileName As String
                    fileName = macroArg.GetValue(customVarValProv, swCutList)
                    
                    Dim filePath As String
                    filePath = GetDirectory(swModel.GetPathName) & fileName
                    
                    If (Not resFilePaths) = -1 Then
                        ReDim resFilePaths(0)
                        ReDim inputBodies(0)
                    Else
                        ReDim Preserve resFilePaths(UBound(resFilePaths) + 1)
                        ReDim Preserve inputBodies(UBound(inputBodies) + 1)
                    End If
                    
                    Dim swBodyFolder As SldWorks.BodyFolder
                    Set swBodyFolder = swCutList.GetSpecificFeature2
                                        
                    If swBodyFolder.GetBodyCount() > 0 Then
                        Set swBody = swBodyFolder.GetBodies()(0)
                    Else
                        Set swBody = Nothing
                    End If
                    
                    resFilePaths(UBound(resFilePaths)) = filePath
                    Set inputBodies(UBound(inputBodies)) = swBody
                    
                Next
                
            Next
            
            Dim vResFiles As Variant
            vResFiles = macroOper.SetResultFiles(resFilePaths)
            
            For i = 0 To UBound(vResFiles)
                
                Dim resFile As IMacroOperationResultFile
                Set resFile = vResFiles(i)
                Set swBody = inputBodies(i)
                
                Dim ext As String
                ext = GetExtension(resFile.path)
                
                TryExportBody swModel, swBody, resFile, macroOper
                 
            Next
            
        Else
            Err.Raise vbError, "", "Only parts are supported"
        End If
        
    Else
        Err.Raise vbError, "", "Open model"
    End If
    
End Sub

Sub TryExportBody(model As SldWorks.ModelDoc2, body As SldWorks.Body2, resFile As IMacroOperationResultFile, macroOper As MacroOperation)

try_:
    On Error GoTo catch_
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    swSelMgr.SuspendSelectionList
    
    Dim swBodies(0) As SldWorks.Body2
    Set swBodies(0) = body
    
    If swSelMgr.AddSelectionListObjects(swBodies, Nothing) = 1 Then
        
        Dim filePath As String
        filePath = resFile.path
        
        Dim errs As Long
        Dim warns As Long
        Dim dir As String
        
        dir = GetDirectory(filePath)
        
        CreateDirectories dir
        
        If False <> model.Extension.SaveAs2(filePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, "", False, errs, warns) Then
            resFile.Status = MacroOperationResultFileStatus_e_Succeeded
        Else
            Err.Raise vbError, "", "Failed to export '" & body.Name & "' to '" & filePath & "'. Error code: " & errs
        End If
    Else
        Err.Raise vbError, "", "Failed to select " & body.Name
    End If

    GoTo finally_
catch_:
    macroOper.ReportIssue Err.Description, MacroIssueType_e_Error
    resFile.Status = MacroOperationResultFileStatus_e_Failed
finally_:

    swSelMgr.ResumeSelectionList2 False
    
End Sub

Sub TryExportFlatPattern(model As SldWorks.ModelDoc2, body As SldWorks.Body2, resFile As IMacroOperationResultFile, macroOper As MacroOperation)

try_:
    On Error GoTo catch_
    
    Dim expData(0) As FlatPatternExportDataCom
    Set expData(0) = New FlatPatternExportDataCom
    
    Set expData(0).body = body
    expData(0).Options = FlatPatternOptionsCom_e.FlatPatternOptionsCom_e_BendLines
    expData(0).OutFilePath = resFile.path
    
    Dim vRes As Variant
    vRes = swCadPlus.FlatPatternExport.BatchExportFlatPatterns(model, expData)
    
    Dim res As FlatPatternExportResult
    Set res = vRes(0)
    
    If False = res.Succeeded Then
        Err.Raise vbError, "", res.Error
    End If
    
    resFile.Status = MacroOperationResultFileStatus_e_Succeeded
    
    GoTo finally_
catch_:
    macroOper.ReportIssue Err.Description, MacroIssueType_e_Error
    resFile.Status = MacroOperationResultFileStatus_e_Failed
finally_:

End Sub

Function GetMacroOperation(Optional dummy As Variant = Empty) As IMacroOperation
    
    Dim macroOper As IMacroOperation
    
    #If TEST Then
        Dim swCadPlusFact As Object
        Set swCadPlusFact = CreateObject("CadPlusFactory.Sw")
        
        Set swCadPlus = swCadPlusFact.Create(swApp, False)
        
        Dim args(1) As String
        args(0) = "MFGs\STEP\{ path [FileNameWithoutExtension] }-{ cutListPrp [Description] }.step"
        Set macroOper = swCadPlus.CreateMacroOperation(swApp.ActiveDoc, "", args)
    #Else
        Dim macroOprMgr As IMacroOperationManager
        Set macroOprMgr = CreateObject("CadPlus.MacroOperationManager")
        
        Set macroOper = macroOprMgr.PopOperation(swApp)
    #End If
    
    Set GetMacroOperation = macroOper
    
End Function

Function GetExtension(path As String) As String
    GetExtension = Right(path, Len(path) - InStrRev(path, "."))
End Function

Function GetDirectory(path As String)
    GetDirectory = Left(path, InStrRev(path, "\"))
End Function

Sub CreateDirectories(path As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(path) Then
        Exit Sub
    End If

    CreateDirectories fso.GetParentFolderName(path)
    
    fso.CreateFolder path
    
End Sub

Function GetCutLists(part As SldWorks.PartDoc) As Variant

    Dim swFeat As SldWorks.Feature
    
    Dim swCutLists() As SldWorks.Feature
    
    Set swFeat = part.FirstFeature
    
    While Not swFeat Is Nothing
        
        If swFeat.GetTypeName2 <> "HistoryFolder" Then
        
            ProcessFeature swFeat, swCutLists
            
            TraverseSubFeatures swFeat, swCutLists
        
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
    GetCutLists = swCutLists
    
End Function

Sub TraverseSubFeatures(parentFeat As SldWorks.Feature, cutLists() As SldWorks.Feature)
    
    Dim swChildFeat As SldWorks.Feature
    Set swChildFeat = parentFeat.GetFirstSubFeature
    
    While Not swChildFeat Is Nothing
        ProcessFeature swChildFeat, cutLists
        Set swChildFeat = swChildFeat.GetNextSubFeature()
    Wend
    
End Sub

Sub ProcessFeature(feat As SldWorks.Feature, cutLists() As SldWorks.Feature)
    
    If feat.GetTypeName2() = "SolidBodyFolder" Then
        Dim swBodyFolder As SldWorks.BodyFolder
        Set swBodyFolder = feat.GetSpecificFeature2
        swBodyFolder.UpdateCutList
    ElseIf feat.GetTypeName2() = "CutListFolder" Then
        
        If Not Contains(cutLists, feat) Then
            If (Not cutLists) = -1 Then
                ReDim cutLists(0)
            Else
                ReDim Preserve cutLists(UBound(cutLists) + 1)
            End If
            
            Set cutLists(UBound(cutLists)) = feat
        End If
        
    End If
    
End Sub

Function Contains(arr As Variant, item As Object) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        If arr(i) Is item Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function