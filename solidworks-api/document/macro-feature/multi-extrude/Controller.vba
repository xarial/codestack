Enum PageModes_e
    Insert
    Edit
End Enum

Const BASE_NAME As String = "MultiBoss-Extrude"

Dim swApp As SldWorks.SldWorks
Dim WithEvents swPage As PropertyPage
Dim swPreviewBodies As Variant

Dim swCurEditFeature As SldWorks.Feature
Dim swCurEditFeatureDef As SldWorks.MacroFeatureData

Private Sub Class_Initialize()
    Set swApp = Application.SldWorks
    Set swPage = New PropertyPage
End Sub

Public Sub InsertExtrude()
    swPage.Show PageModes_e.Insert
End Sub

Public Sub EditExtrude(feat As SldWorks.Feature)
    
    Set swCurEditFeature = feat
    Set swCurEditFeatureDef = feat.GetDefinition
    
    swCurEditFeatureDef.AccessSelections swApp.ActiveDoc, Nothing
    
    Dim vSelection As Variant
    swCurEditFeatureDef.GetSelections3 vSelection, Empty, Empty, Empty, Empty
    
    Dim vParamValues As Variant
    swCurEditFeatureDef.GetParameters Empty, Empty, vParamValues
    
    swPage.Show PageModes_e.Edit, vSelection, vParamValues
    
    Preview vSelection, vParamValues
    
End Sub
    
Private Sub swPage_Closed(mode As Integer, vSketches As Variant, vDepths As Variant, isCancelled As Boolean)
    
    HidePreview
    
    Select Case mode
        Case PageModes_e.Insert
            If Not isCancelled Then
                InsertMacroFeature vSketches, vDepths
            End If
        Case PageModes_e.Edit
            If Not isCancelled Then
                ModifyMacroFeature vSketches, vDepths
            Else
                swCurEditFeatureDef.ReleaseSelectionAccess
            End If
    End Select
    
End Sub

Sub InsertMacroFeature(vSketches As Variant, vDepths As Variant)
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim curMacroPath As String
    curMacroPath = swApp.GetCurrentMacroPathName
    
    Dim vMethods(8) As String
                    
    Const MACRO_FEATURE_MODULE_NAME As String = "MacroFeature"
    
    vMethods(0) = curMacroPath: vMethods(1) = MACRO_FEATURE_MODULE_NAME: vMethods(2) = "swmRebuild"
    vMethods(3) = curMacroPath: vMethods(4) = MACRO_FEATURE_MODULE_NAME: vMethods(5) = "swmEditDefinition"
    vMethods(6) = curMacroPath: vMethods(7) = MACRO_FEATURE_MODULE_NAME: vMethods(8) = "swmSecurity"
    
    Dim iconsDir As String
    iconsDir = swApp.GetCurrentMacroPathFolder() & "\Icons\"
    
    Dim icons(8) As String
    icons(0) = iconsDir & "extrude_20x20.bmp"
    icons(1) = iconsDir & "extrude-suppressed_20x20.bmp"
    icons(2) = iconsDir & "extrude_20x20.bmp"
    
    icons(3) = iconsDir & "extrude_32x32.bmp"
    icons(4) = iconsDir & "extrude-suppressed_32x32.bmp"
    icons(5) = iconsDir & "extrude_32x32.bmp"
    
    icons(6) = iconsDir & "extrude_40x40.bmp"
    icons(7) = iconsDir & "extrude-suppressed_40x40.bmp"
    icons(8) = iconsDir & "extrude_40x40.bmp"

    Dim vParamNames As Variant
    Dim vParamTypes As Variant
    Dim vParamValues As Variant
    
    CreateParameters vDepths, vParamNames, vParamTypes, vParamValues
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FeatureManager.InsertMacroFeature3(BASE_NAME, "", vMethods, _
        vParamNames, vParamTypes, vParamValues, Empty, Empty, Empty, _
        icons, swMacroFeatureOptions_e.swMacroFeatureEmbedMacroFile)
    
    If swFeat Is Nothing Then
        MsgBox "Failed to create feature"
    End If
    
End Sub

Sub ModifyMacroFeature(vSketches As Variant, vDepths As Variant)
    
    Dim vParamNames As Variant
    Dim vParamTypes As Variant
    Dim vParamValues As Variant
    
    CreateParameters vDepths, vParamNames, vParamTypes, vParamValues
    swCurEditFeatureDef.SetParameters vParamNames, vParamTypes, vParamValues
    
    Dim swSelMarks() As Long
    Dim swViews() As SldWorks.View
    
    ReDim swSelMarks(UBound(vSketches))
    ReDim swViews(UBound(vSketches))
    
    swCurEditFeatureDef.SetSelections2 vSketches, swSelMarks, swViews
    
    swCurEditFeature.ModifyDefinition swCurEditFeatureDef, swApp.ActiveDoc, Nothing
    
End Sub

Sub CreateParameters(vDepths As Variant, ByRef vParamNames As Variant, ByRef vParamTypes As Variant, ByRef vParamValues As Variant)
        
    Dim sParamNames() As String
    Dim iParamTypes() As Long
    Dim dParamValues() As String
        
    ReDim sParamNames(UBound(vDepths))
    ReDim iParamTypes(UBound(vDepths))
    ReDim dParamValues(UBound(vDepths))
    
    Dim i As Integer
    For i = 0 To UBound(vDepths)
        sParamNames(i) = "DEPTH" & i + 1
        iParamTypes(i) = swMacroFeatureParamType_e.swMacroFeatureParamTypeDouble
        dParamValues(i) = CStr(vDepths(i))
    Next
    
    vParamNames = sParamNames
    vParamTypes = iParamTypes
    vParamValues = dParamValues
    
End Sub

Private Sub swPage_DataChanged(vSketches As Variant, vDepths As Variant)
    Preview vSketches, vDepths
End Sub

Sub Preview(vSketches As Variant, vDepths As Variant)
    
    HidePreview
    
    swPreviewBodies = Geometry.CreateBodiesFromSketches(vSketches, vDepths)
    
    If Not IsEmpty(swPreviewBodies) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(swPreviewBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = swPreviewBodies(i)
            swBody.Display3 swApp.ActiveDoc, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
        Next
        
    End If
    
End Sub

Sub HidePreview()
    
    Dim i As Integer
    
    If Not IsEmpty(swPreviewBodies) Then
                
        For i = 0 To UBound(swPreviewBodies)
            Set swPreviewBodies(i) = Nothing
        Next
                
    End If
    
End Sub