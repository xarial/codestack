Enum SheetMetalOptions_e
    ExportFlatPatternGeometry = 1
    IncludeHiddenEdges = 2
    ExportBendLines = 4
    IncludeSketches = 8
    MergeCoplanarFaces = 16
    ExportLibraryFeatures = 32
    ExportFormingTools = 64
    ExportBoundingBox = 2048
End Enum

Const SKIP_EXISTING_FILES As Boolean = False

Const OPEN_APP_PATH As String = "" 'Path to the application to open the file (e.g. eDrawings), empty string to not open the file, * to open with default application

Const OUT_NAME_TEMPLATE As String = "DXFs\<_FileName_>_<_FeatureName_>_<_ConfName_>_<$CLPRP:Description>.dxf"

Const FLAT_PATTERN_OPTIONS As Integer = SheetMetalOptions_e.ExportBendLines + SheetMetalOptions_e.ExportFlatPatternGeometry

Dim swApp As SldWorks.SldWorks

Sub main()
        
    Set swApp = Application.SldWorks
    
try_:
    On Error GoTo catch_
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Err.Raise vbError, "", "Please open assembly or part document"
    End If
    
    If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
    
        Dim swAssy As SldWorks.AssemblyDoc
        
        Set swAssy = swModel
            
        swAssy.ResolveAllLightWeightComponents True
            
        Dim vComps As Variant
        vComps = GetDistinctSheetMetalComponents(swAssy)
        
        Dim i As Integer
        
        For i = 0 To UBound(vComps)
        
            Dim swComp As SldWorks.Component2
            Set swComp = vComps(i)
            
            ProcessSheetMetalModel swAssy, swComp.GetModelDoc2(), swComp.ReferencedConfiguration
        
        Next
        
    ElseIf swModel.GetType() = swDocumentTypes_e.swDocPART Then
        
        Dim swPart As SldWorks.PartDoc
        Set swPart = swApp.ActiveDoc
        
        ProcessSheetMetalModel swPart, swPart, swPart.ConfigurationManager.ActiveConfiguration.Name
        
    Else
        Err.Raise vbError, "", "Only assembly and part documents are supported"
    End If
    
    swApp.SendMsgToUser2 "Operation completed", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    
    GoTo finally_
    
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Function GetDistinctSheetMetalComponents(assy As SldWorks.AssemblyDoc) As Variant
    
    Dim vComps As Variant
    vComps = assy.GetComponents(False)
        
    Dim i As Integer
    
    Dim swSheetMetalComps() As SldWorks.Component2
    
    For i = 0 To UBound(vComps)
        
        Dim swComp As SldWorks.Component2
        Set swComp = vComps(i)
                
        If False = swComp.IsSuppressed() Then
            
            If Not ContainsComponent(swSheetMetalComps, swComp) Then
                
                If IsSheetMetalComponent(swComp) Then
                    If (Not swSheetMetalComps) = -1 Then
                        ReDim swSheetMetalComps(0)
                    Else
                        ReDim Preserve swSheetMetalComps(UBound(swSheetMetalComps) + 1)
                    End If
                    
                    Set swSheetMetalComps(UBound(swSheetMetalComps)) = swComp
                End If
            
            End If
            
        End If
        
    Next
    
    If (Not swSheetMetalComps) = -1 Then
        GetDistinctSheetMetalComponents = Empty
    Else
        GetDistinctSheetMetalComponents = swSheetMetalComps
    End If
    
End Function

Function IsSheetMetalComponent(comp As SldWorks.Component2) As Boolean
    
    Dim vBodies As Variant
    vBodies = comp.GetBodies3(swBodyType_e.swSolidBody, Empty)
    
    If Not IsEmpty(vBodies) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            
            If False <> swBody.IsSheetMetal() Then
                IsSheetMetalComponent = True
                Exit Function
            End If
            
        Next
    End If
    
    IsSheetMetalComponent = False
    
End Function

Function ContainsComponent(comps As Variant, swComp As SldWorks.Component2) As Boolean

    Dim i As Integer
    
    For i = 0 To UBound(comps)
        Dim swThisComp As SldWorks.Component2
        Set swThisComp = comps(i)
        
        If swThisComp.GetPathName() = swComp.GetPathName() And swThisComp.ReferencedConfiguration = swComp.ReferencedConfiguration Then
            ContainsComponent = True
            Exit Function
        End If
    Next
    
    ContainsComponent = False

End Function

Function ComposeOutFileName(template As String, rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String, flatPatternFeat As SldWorks.Feature, cutListFeat As SldWorks.Feature) As String

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
        
        outFileName = Left(outFileName, regExMatch.FirstIndex) & ResolveToken(tokenName, rootModel, sheetMetalModel, conf, flatPatternFeat, cutListFeat) & Right(outFileName, Len(outFileName) - (regExMatch.FirstIndex + regExMatch.Length))
    Next
    
    ComposeOutFileName = ReplaceInvalidPathSymbols(GetFullPath(rootModel, outFileName))
    
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

Function ResolveToken(token As String, rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String, flatPatternFeat As SldWorks.Feature, cutListFeat As SldWorks.Feature) As String
    
    Const FILE_NAME_TOKEN As String = "_FileName_"
    Const ASSM_FILE_NAME_TOKEN As String = "_AssmFileName_"
    Const FEAT_NAME_TOKEN As String = "_FeatureName_"
    Const CONF_NAME_TOKEN As String = "_ConfName_"
    
    Const PRP_TOKEN As String = "$PRP:"
    Const CUT_LIST_PRP_TOKEN As String = "$CLPRP:"
    Const ASM_PRP_TOKEN As String = "$ASSMPRP:"
    
    Select Case LCase(token)
        Case LCase(FILE_NAME_TOKEN)
            ResolveToken = GetFileNameWithoutExtension(sheetMetalModel.GetPathName)
        Case LCase(FEAT_NAME_TOKEN)
            ResolveToken = flatPatternFeat.Name
        Case LCase(CONF_NAME_TOKEN)
            ResolveToken = conf
        Case LCase(ASSM_FILE_NAME_TOKEN)
            If rootModel.GetPathName() = "" Then
                Err.Raise vbError, "", "Assembly must be saved to use " & ASSM_FILE_NAME_TOKEN
            End If
            ResolveToken = GetFileNameWithoutExtension(rootModel.GetPathName())
        Case Else
            
            Dim prpName As String
                        
            If Left(token, Len(PRP_TOKEN)) = PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(PRP_TOKEN))
                ResolveToken = GetModelPropertyValue(sheetMetalModel, conf, prpName)
            ElseIf Left(token, Len(ASM_PRP_TOKEN)) = ASM_PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(ASM_PRP_TOKEN))
                ResolveToken = GetModelPropertyValue(rootModel, rootModel.ConfigurationManager.ActiveConfiguration.Name, prpName)
            ElseIf Left(token, Len(CUT_LIST_PRP_TOKEN)) = CUT_LIST_PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(CUT_LIST_PRP_TOKEN))
                ResolveToken = GetPropertyValue(cutListFeat.CustomPropertyManager, prpName)
            Else
                Err.Raise vbError, "", "Unrecognized token: " & token
            End If
            
    End Select
    
End Function

Function GetModelPropertyValue(model As SldWorks.ModelDoc2, confName As String, prpName As String) As String
    
    Dim prpVal As String
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)
    prpVal = GetPropertyValue(swCustPrpMgr, prpName)
    
    If prpVal = "" Then
        Set swCustPrpMgr = model.Extension.CustomPropertyManager("")
        prpVal = GetPropertyValue(swCustPrpMgr, prpName)
    End If
    
    GetModelPropertyValue = prpVal
    
End Function

Function GetPropertyValue(custPrpMgr As SldWorks.CustomPropertyManager, prpName As String) As String
    Dim resVal As String
    custPrpMgr.Get2 prpName, "", resVal
    GetPropertyValue = resVal
End Function

Function GetFileNameWithoutExtension(path As String) As String
    GetFileNameWithoutExtension = Mid(path, InStrRev(path, "\") + 1, InStrRev(path, ".") - InStrRev(path, "\") - 1)
End Function

Function GetCutListFeatures(model As SldWorks.ModelDoc2) As Variant
    GetCutListFeatures = GetFeaturesByType(model, "CutListFolder")
End Function

Function GetFlatPatternFeatures(model As SldWorks.ModelDoc2) As Variant
    GetFlatPatternFeatures = GetFeaturesByType(model, "FlatPattern")
End Function

Sub ProcessSheetMetalModel(rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String)

    Dim activeConfName As String

    Dim error As ErrObject

try_:
    
    On Error GoTo catch_
    
    activeConfName = sheetMetalModel.ConfigurationManager.ActiveConfiguration.Name
    
    If LCase(activeConfName) <> LCase(conf) Then
        If False = sheetMetalModel.ShowConfiguration2(conf) Then
            Err.Raise vbError, "", "Failed to activate configuration"
        End If
    End If
        
    Dim vCutListFeats As Variant
    vCutListFeats = GetCutListFeatures(sheetMetalModel)
    
    If Not IsEmpty(vCutListFeats) Then
        
        Dim vFlatPatternFeats As Variant
        vFlatPatternFeats = GetFlatPatternFeatures(sheetMetalModel)
        
        If Not IsEmpty(vFlatPatternFeats) Then
            
            Dim swProcessedCutListsFeats() As SldWorks.Feature
            
            Dim i As Integer
    
            For i = 0 To UBound(vFlatPatternFeats)
                
                Dim swFlatPatternFeat As SldWorks.Feature
                Dim swFlatPattern As SldWorks.FlatPatternFeatureData
                
                Set swFlatPatternFeat = vFlatPatternFeats(i)
                
                Set swFlatPattern = swFlatPatternFeat.GetDefinition
                
                Dim swFixedEnt As SldWorks.Entity
                
                Set swFixedEnt = swFlatPattern.FixedFace2
                
                Dim swBody As SldWorks.Body2
                
                If TypeOf swFixedEnt Is SldWorks.Face2 Then
                    Dim swFixedFace As SldWorks.Face2
                    Set swFixedFace = swFixedEnt
                    Set swBody = swFixedFace.GetBody
                ElseIf TypeOf swFixedEnt Is SldWorks.Edge Then
                    Dim swFixedEdge As SldWorks.Edge
                    Set swFixedEdge = swFixedEnt
                    Set swBody = swFixedEdge.GetBody
                ElseIf TypeOf swFixedEnt Is SldWorks.Vertex Then
                    Dim swFixedVert As SldWorks.Vertex
                    Set swFixedVert = swFixedEnt
                    Set swBody = swFixedVert.GetBody
                End If
                
                Dim swCutListFeat As SldWorks.Feature
                Set swCutListFeat = FindCutListFeature(vCutListFeats, swBody)
                
                If Not swCutListFeat Is Nothing Then
                    
                    Dim isUnique As Boolean
                                        
                    If (Not swProcessedCutListsFeats) = -1 Then
                        isUnique = True
                    ElseIf Not ContainsSwObject(swProcessedCutListsFeats, swCutListFeat) Then
                        isUnique = True
                    Else
                        isUnique = False
                    End If
                    
                    If isUnique Then
                        
                        If (Not swProcessedCutListsFeats) = -1 Then
                            ReDim swProcessedCutListsFeats(0)
                        Else
                            ReDim Preserve swProcessedCutListsFeats(UBound(swProcessedCutListsFeats) + 1)
                        End If
                        
                        Set swProcessedCutListsFeats(UBound(swProcessedCutListsFeats)) = swCutListFeat
                        
                        Dim outFileName As String
                        outFileName = ComposeOutFileName(OUT_NAME_TEMPLATE, rootModel, sheetMetalModel, conf, swFlatPatternFeat, swCutListFeat)
                        
                        If Not SKIP_EXISTING_FILES Or Not FileExists(outFileName) Then
                            ExportFlatPattern sheetMetalModel, swFlatPatternFeat, outFileName, FLAT_PATTERN_OPTIONS
                            If OPEN_APP_PATH <> "" Then
                                OpenFile outFileName
                            End If
                        End If
                    End If
                    
                Else
                    Err.Raise vbError, "", "Failed to find cut-list for flat pattern " & swFlatPatternFeat.Name
                End If
                
            Next
            
        Else
            Err.Raise vbError, "", "No flat pattern features found"
        End If
        
    Else
        Err.Raise vbError, "", "No cut-list items found"
    End If
    
    GoTo finally_
catch_:
    Set error = Err
finally_:

    sheetMetalModel.ShowConfiguration2 activeConfName
    
    If Not error Is Nothing Then
        Err.Raise error.Number, error.Source, error.Description, error.HelpFile, error.HelpContext
    End If
    
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Function FindCutListFeature(vCutListFeats As Variant, body As SldWorks.Body2) As SldWorks.Feature
    
    Dim i As Integer
    
    For i = 0 To UBound(vCutListFeats)
        
        Dim swCutListFeat As SldWorks.Feature
        Set swCutListFeat = vCutListFeats(i)
        
        Dim swBodyFolder As SldWorks.BodyFolder
        Set swBodyFolder = swCutListFeat.GetSpecificFeature2
        
        Dim vBodies As Variant
        
        vBodies = swBodyFolder.GetBodies
        
        If ContainsSwObject(vBodies, body) Then
            Set FindCutListFeature = swCutListFeat
        End If
            
    Next
    
End Function

Function ContainsSwObject(vArr As Variant, obj As Object) As Boolean
    
    If Not IsEmpty(vArr) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vArr)
            
            Dim swObj As Object
            Set swObj = vArr(i)
            
            If swApp.IsSame(swObj, obj) = swObjectEquality.swObjectSame Then
                ContainsSwObject = True
                Exit Function
            End If
        Next
        
    End If
    
    ContainsSwObject = False
    
End Function

Function GetFeaturesByType(model As SldWorks.ModelDoc2, typeName As String) As Variant
    
    Dim swFeats() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        
        If typeName = "CutListFolder" And swFeat.GetTypeName2() = "SolidBodyFolder" Then
            Dim swBodyFolder As SldWorks.BodyFolder
            Set swBodyFolder = swFeat.GetSpecificFeature2
            swBodyFolder.UpdateCutList
        End If
        
        ProcessFeature swFeat, swFeats, typeName

        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
    If (Not swFeats) = -1 Then
        GetFeaturesByType = Empty
    Else
        GetFeaturesByType = swFeats
    End If
    
End Function

Sub ProcessFeature(thisFeat As SldWorks.Feature, featsArr() As SldWorks.Feature, typeName As String)
    
    If thisFeat.GetTypeName2() = typeName Then
    
        If (Not featsArr) = -1 Then
            ReDim featsArr(0)
            Set featsArr(0) = thisFeat
        Else
            Dim i As Integer
            
            For i = 0 To UBound(featsArr)
                If swApp.IsSame(featsArr(i), thisFeat) = swObjectEquality.swObjectSame Then
                    Exit Sub
                End If
            Next
            
            ReDim Preserve featsArr(UBound(featsArr) + 1)
            Set featsArr(UBound(featsArr)) = thisFeat
        End If
    
    End If
    
    Dim swSubFeat As SldWorks.Feature
    Set swSubFeat = thisFeat.GetFirstSubFeature
        
    While Not swSubFeat Is Nothing
        ProcessFeature swSubFeat, featsArr, typeName
        Set swSubFeat = swSubFeat.GetNextSubFeature
    Wend
        
End Sub

Sub ExportFlatPattern(part As SldWorks.PartDoc, flatPattern As SldWorks.Feature, outFilePath As String, opts As SheetMetalOptions_e)
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = part
    
    Dim error As ErrObject
    Dim hide As Boolean

try_:
    
    On Error GoTo catch_

    If False = swModel.Visible Then
        hide = True
        swModel.Visible = True
    End If
    
    swApp.ActivateDoc3 swModel.GetPathName(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, 0
    
    swModel.FeatureManager.EnableFeatureTree = False
    swModel.FeatureManager.EnableFeatureTreeWindow = False
    swModel.ActiveView.EnableGraphicsUpdate = False
    
    Dim outDir As String
    outDir = Left(outFilePath, InStrRev(outFilePath, "\"))
    
    CreateDirectories outDir
    
    Dim modelPath As String
    
    modelPath = part.GetPathName
    
    If modelPath = "" Then
        Err.Raise vbError, "", "Part document must be saved"
    End If
    
    If False <> flatPattern.Select2(False, -1) Then
        If False = part.ExportToDWG2(outFilePath, modelPath, swExportToDWG_e.swExportToDWG_ExportSheetMetal, True, Empty, False, False, opts, Empty) Then
            Err.Raise vbError, "", "Failed to export flat pattern"
        End If
    Else
        Err.Raise vbError, "", "Failed to select flat-pattern"
    End If
    
    GoTo finally_
    
catch_:
    Set error = Err
finally_:

    swModel.FeatureManager.EnableFeatureTree = True
    swModel.FeatureManager.EnableFeatureTreeWindow = True
    swModel.ActiveView.EnableGraphicsUpdate = True
    
    If hide Then
        swApp.CloseDoc swModel.GetTitle
    End If
    
    If Not error Is Nothing Then
        Err.Raise error.Number, error.Source, error.Description, error.HelpFile, error.HelpContext
    End If
    
End Sub

Sub OpenFile(filePath As String)
                    
    If OPEN_APP_PATH = "*" Then
        Shell """" & filePath & """"
    Else
        Shell """" & OPEN_APP_PATH & """" & " " & """" & filePath & """"
    End If
    
End Sub

Sub CreateDirectories(path As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(path) Then
        Exit Sub
    End If

    CreateDirectories fso.GetParentFolderName(path)
    
    fso.CreateFolder path
    
End Sub

Function GetFullPath(model As SldWorks.ModelDoc2, path As String)
    
    GetFullPath = path
        
    If IsPathRelative(path) Then
        
        If Left(path, 1) <> "\" Then
            path = "\" & path
        End If
        
        Dim modelPath As String
        Dim modelDir As String
        
        modelPath = model.GetPathName
        
        modelDir = Left(modelPath, InStrRev(modelPath, "\") - 1)
        
        GetFullPath = modelDir & path
        
    End If
    
End Function

Function IsPathRelative(path As String)
    IsPathRelative = Mid(path, 2, 1) <> ":" And Not IsPathUnc(path)
End Function

Function IsPathUnc(path As String)
    IsPathUnc = Left(path, 2) = "\\"
End Function