Const SW_VERSION As Integer = -1 'save into previous version

Const PREFIX As String = ""
Const SUFFIX As String = "_PREV"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swAdvancedSaveAsOpts As SldWorks.AdvancedSaveAsOptions
        
        Set swAdvancedSaveAsOpts = swModel.Extension.GetAdvancedSaveAsOptions(swSaveWithReferencesOptions_e.swSaveWithReferencesOptions_None)
                
        swAdvancedSaveAsOpts.SaveAsPreviousVersion = GetVersionNumber(SW_VERSION)
        swAdvancedSaveAsOpts.SaveAllAsCopy = True
                
        Dim vIds As Variant
        Dim vNames As Variant
        Dim vPaths As Variant
        
        swAdvancedSaveAsOpts.GetItemsNameAndPath vIds, vNames, vPaths
        
        Dim i As Integer
        
        For i = 0 To UBound(vNames)
            vNames(i) = ComposeName(CStr(vNames(i)))
        Next
        
        swAdvancedSaveAsOpts.ModifyItemsNameAndPath vIds, vNames, vPaths
                
        Dim errs As Long
        Dim warns As Long
        
        Dim path As String
        path = swModel.GetPathName
                
        If path <> "" Then
            
            Dim dir As String
            
            dir = GetDirectory(path)
            
            Dim fileName As String
            
            fileName = ComposeName(GetFileName(path))
            
            If False = swModel.Extension.SaveAs3(dir & fileName, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, swAdvancedSaveAsOpts, errs, warns) Then
                Err.Raise vbError, "", "Failed to save model: " & errs
            End If
        
        Else
            Err.Raise vbError, "", "Active model is never saved"
        End If
        
    Else
        Err.Raise vbError, "", "Open model"
    End If
    
End Sub

Function GetVersionNumber(swVers As Integer) As Integer
    
    Dim revNmb As Integer
    revNmb = CInt(Split(swApp.RevisionNumber, ".")(0))
    
    Const SW_2022_REVISION As Integer = 30
    Const SW_2022_VERSION As Integer = 15000
    Const SW_VERSION_STEP As Integer = 1000
    
    Dim versOffset As Integer
    
    versOffset = revNmb + swVers - SW_2022_REVISION
    
    If versOffset >= 0 Then
        GetVersionNumber = SW_2022_VERSION + SW_VERSION_STEP * versOffset
    Else
        Err.Raise vbError, "", "Minimum supported version is SOLIDWORKS 2022"
    End If
    
End Function

Function ComposeName(fileName As String) As String

    Dim ext As String
    ext = Right(fileName, Len(".SLDXXX"))
            
    ComposeName = PREFIX & Left(fileName, Len(fileName) - Len(ext)) & SUFFIX & ext

End Function

Function GetFileName(path As String) As String
    
    Dim fileName As String
    
    fileName = Right(path, Len(path) - InStrRev(path, "\"))
    
    GetFileName = fileName
    
End Function

Function GetDirectory(path As String)
    GetDirectory = Left(path, InStrRev(path, "\"))
End Function