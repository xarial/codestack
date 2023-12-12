'#Const TEST = True

Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Boolean

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
        
    Dim macroOper As IMacroOperation
    Set macroOper = GetMacroOperation()
    
    Dim vArgs As Variant
    vArgs = macroOper.Arguments
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = macroOper.model
        
    Dim resFilePaths() As String
    ReDim resFilePaths(UBound(vArgs))
        
    For i = 0 To UBound(vArgs)
        
        Dim macroArg As IMacroArgument
        Set macroArg = vArgs(i)
        
        Dim filePath As String
        filePath = macroArg.GetValue()
        
        If PathIsRelative(filePath) Then
            
            Dim modelPath As String
            modelPath = swModel.GetPathName
            
            If modelPath <> "" Then
                filePath = GetDirectory(modelPath) & filePath
            Else
                Err.Raise vbError, "", "Cannot use relative path for an unsaved model"
            End If
        
        End If
               
        resFilePaths(i) = filePath
        
    Next
    
    Dim vResFiles As Variant
    vResFiles = macroOper.SetResultFiles(resFilePaths)
    
    For i = 0 To UBound(vResFiles)
        
        Dim resFile As IMacroOperationResultFile
        Set resFile = vResFiles(i)
                
        TryExport swModel, resFile, macroOper
                 
    Next
    
End Sub

Sub TryExport(model As SldWorks.ModelDoc2, resFile As IMacroOperationResultFile, macroOper As MacroOperation)

try_:
    On Error GoTo catch_
    
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
        Err.Raise vbError, "", "Failed to export to '" & filePath & "'. Error code: " & errs
    End If

    GoTo finally_
catch_:
    macroOper.ReportIssue Err.Description, MacroIssueType_e_Error
    resFile.Status = MacroOperationResultFileStatus_e_Failed
finally_:
    
End Sub

Function GetMacroOperation() As IMacroOperation
    
    Dim macroOper As IMacroOperation
    
    #If TEST Then
        Dim swCadPlusFact As Object
        Set swCadPlusFact = CreateObject("CadPlusFactory.Sw")
        
        Set swCadPlus = swCadPlusFact.Create(swApp, False)
        
        Dim args(1) As String
        args(0) = "MFGs\STEP\{ path [FileNameWithoutExtension] }.step"
        args(1) = "MFGs\IGES\{ path [FileNameWithoutExtension] }.igs"
        Set macroOper = swCadPlus.CreateMacroOperation(swApp.ActiveDoc, "", args)
    #Else
        Dim macroOprMgr As IMacroOperationManager
        Set macroOprMgr = CreateObject("CadPlus.MacroOperationManager")
        
        Set macroOper = macroOprMgr.PopOperation(swApp)
    #End If
    
    Set GetMacroOperation = macroOper
    
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