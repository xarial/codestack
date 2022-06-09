Const ALL_CONFIGS As Boolean = False
Const OUT_FOLDER As String = ""
Const STEP_VERSION As Long = 214 '203 or 214

Dim OUT_NAME_TEMPLATES As Variant

Dim swApp As SldWorks.SldWorks

Sub main()
        
    Dim origStepVersion As Long
        
    OUT_NAME_TEMPLATES = Array("PDFs\<_FileName_>_<_ConfName_>_<PartNo>.pdf", "IMGs\<_FileName_>_<_ConfName_>_<PartNo>.jpg")
    
    Set swApp = Application.SldWorks
    
try_:
    On Error GoTo catch_
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Err.Raise vbError, "", "Please open document"
    End If
    
    If swModel.GetPathName() = "" Then
        Err.Raise vbError, "", "Please save the model"
    End If
    
    Dim outFolder As String
    
    If Not TryGetOutDirFromArguments(outFolder) Then
        outFolder = OUT_FOLDER
    End If
    
    ReadOptions origStepVersion
    SetupOptions STEP_VERSION
    
    ExportFile swModel, OUT_NAME_TEMPLATES, ALL_CONFIGS, outFolder
    
    GoTo finally_
    
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

    SetupOptions origStepVersion

End Sub

Sub ReadOptions(ByRef stepVersion As Long)

    stepVersion = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP)
    
End Sub

Sub SetupOptions(stepVersion As Long)
    
    If False = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP, stepVersion) Then
        Err.Raise vbError, "", "Failed to set Step Export version to " & stepVersion
    End If
    
End Sub

Sub ExportFile(model As SldWorks.ModelDoc2, vOutNameTemplates As Variant, allConfigs As Boolean, outFolder As String)
    
    Dim i As Integer
    Dim j As Integer
    
    Dim curConf As String
    
    If model.GetType() = swDocumentTypes_e.swDocDRAWING Then
        Dim swDraw As SldWorks.DrawingDoc
        Set swDraw = model
        curConf = swDraw.GetCurrentSheet().GetName
    Else
        curConf = model.ConfigurationManager.ActiveConfiguration.Name
    End If
    
    Dim vConfs As Variant
    
    If allConfigs Then
        If model.GetType() = swDocumentTypes_e.swDocDRAWING Then
            vConfs = model.GetSheetNames()
        Else
            vConfs = model.GetConfigurationNames()
        End If
    Else
        Dim sConfs(0) As String
        sConfs(0) = curConf
        vConfs = sConfs
    End If
    
    For i = 0 To UBound(vConfs)
    
        If model.GetType() = swDocumentTypes_e.swDocDRAWING Then
            curConf = swDraw.ActivateSheet(CStr(vConfs(i)))
        Else
            model.ShowConfiguration2 CStr(vConfs(i))
        End If
                
        For j = 0 To UBound(vOutNameTemplates)
            
            Dim errs As Long
            Dim warns As Long
        
            Dim outNameTemplate As String
            outNameTemplate = vOutNameTemplates(j)
            
            Dim outFilePath As String
            outFilePath = ComposeOutFileName(outNameTemplate, model, outFolder)

            Dim outDir As String
            outDir = Left(outFilePath, InStrRev(outFilePath, "\"))
    
            CreateDirectories outDir
    
            If False = model.Extension.SaveAs(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errs, warns) Then
                Err.Raise vberrror, "", "Failed to export to " & outFilePath
            End If
            
        Next
        
    Next
    
    If model.GetType() = swDocumentTypes_e.swDocDRAWING Then
        curConf = swDraw.ActivateSheet(curConf)
    Else
        model.ShowConfiguration2 curConf
    End If
    
End Sub

Function ComposeOutFileName(template As String, model As SldWorks.ModelDoc2, outFolder As String) As String

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<[^>]*>"
    
    Dim regExMatches As Object
    Set regExMatches = regEx.Execute(template)
    
    Dim i As Integer
    
    Dim outFileName As String
    outFileName = template
    
    For i = regExMatches.Count - 1 To 0 Step -1
        
        Dim regExMatch As Object
        Set regExMatch = regExMatches.Item(i)
                    
        Dim tokenName As String
        tokenName = Mid(regExMatch.Value, 2, Len(regExMatch.Value) - 2)
        
        outFileName = Left(outFileName, regExMatch.FirstIndex) & ResolveToken(tokenName, model) & Right(outFileName, Len(outFileName) - (regExMatch.FirstIndex + regExMatch.Length))
    Next
    
    ComposeOutFileName = ReplaceInvalidPathSymbols(GetFullPath(model, outFileName, outFolder))
    
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

Function ResolveToken(token As String, model As SldWorks.ModelDoc2) As String
    
    Const FILE_NAME_TOKEN As String = "_FileName_"
    Const CONF_NAME_TOKEN As String = "_ConfName_"
    
    Select Case LCase(token)
        Case LCase(FILE_NAME_TOKEN)
            ResolveToken = GetFileNameWithoutExtension(model.GetPathName)
        Case LCase(CONF_NAME_TOKEN)
            If model.GetType() = swDocumentTypes_e.swDocDRAWING Then
                Dim swDraw As SldWorks.DrawingDoc
                Set swDraw = model
                ResolveToken = swDraw.GetCurrentSheet().GetName
            Else
                ResolveToken = model.ConfigurationManager.ActiveConfiguration.Name
            End If
        Case Else
            
            Dim swCustPrpMgr As SldWorks.CustomPropertyManager
            Dim resVal As String
            resVal = ""
            
            If model.GetType() <> swDocumentTypes_e.swDocDRAWING Then
                Set swCustPrpMgr = model.Extension.CustomPropertyManager(model.ConfigurationManager.ActiveConfiguration.Name)
                swCustPrpMgr.Get2 token, "", resVal
            End If
            
            If resVal = "" Then
                Set swCustPrpMgr = model.Extension.CustomPropertyManager("")
                swCustPrpMgr.Get2 token, "", resVal
            End If
            
            ResolveToken = resVal
    End Select
    
End Function

Function GetFileNameWithoutExtension(path As String) As String
    
    GetFileNameWithoutExtension = Mid(path, InStrRev(path, "\") + 1, InStrRev(path, ".") - InStrRev(path, "\") - 1)
    
End Function

Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
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

Function GetFullPath(model As SldWorks.ModelDoc2, path As String, outFolder As String)
    
    GetFullPath = path
        
    If IsPathRelative(path) Then
        
        If Left(path, 1) <> "\" Then
            path = "\" & path
        End If
        
        If outFolder = "" Then
        
            Dim modelPath As String
            Dim modelDir As String
            
            modelPath = model.GetPathName
            
            modelDir = Left(modelPath, InStrRev(modelPath, "\") - 1)
            
            outFolder = modelDir
        Else
            If Right(outFolder, 1) = "\" Then
                outFolder = Left(outFolder, Len(outFolder) - 1)
            End If
        End If
        
        GetFullPath = outFolder & path
        
    End If
    
End Function

Function IsPathRelative(path As String)
    IsPathRelative = Mid(path, 2, 1) <> ":" And Not IsPathUnc(path)
End Function

Function IsPathUnc(path As String)
    IsPathUnc = Left(path, 2) = "\\"
End Function

Function TryGetOutDirFromArguments(ByRef outDir As String) As Boolean

try_:

    On Error GoTo catch_

    Dim macroRunner As Object
    Set macroRunner = CreateObject("CadPlus.MacroRunner.Sw")
    
    Dim param As Object
    Set param = macroRunner.PopParameter(swApp)
    
    Dim vArgs As Variant
    vArgs = param.Get("Args")
    
    outDir = CStr(vArgs(0))
    TryGetOutDirFromArguments = True
    GoTo finally_
    
catch_:
    TryGetOutDirFromArguments = False
finally_:

End Function