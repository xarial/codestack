Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
try:
        
    On Error GoTo catch
        
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim outFilePath As String
        outFilePath = GetOutputFilePath(swModel)
        
        Dim swSketch As SldWorks.Feature
        Set swSketch = MergeSelectedSketchSegments(swModel)
        
        ExportSketchToIges swModel, swSketch, outFilePath
        
        If False <> swSketch.Select2(False, -1) Then
            If False = swModel.Extension.DeleteSelection2(swDeleteSelectionOptions_e.swDelete_Absorbed) Then
                err.Raise vbError, "", "Failed to delete temp sketch " & swSketch.Name
            End If
        Else
            err.Raise vbError, "", "Failed to select merged sketch " & swSketch.Name & " to delete"
        End If
                
    Else
        err.Raise vbError, "", "Please open model"
    End If
    
    GoTo finally
    
catch:
    Debug.Print err.Number
    swApp.SendMsgToUser2 err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally:

End Sub

Function GetOutputFilePath(model As SldWorks.ModelDoc2) As String
    
    Dim outFilePath As String
    
    Dim index As Integer
    index = -1
    
    Do
        Dim suffix As String
        
        index = index + 1
        
        If index > 0 Then
            suffix = "(" & index & ")"
        End If
        
        outFilePath = model.GetPathName()
        
        If outFilePath = "" Then
            err.Raise vbError, "", "Source file is not saved to disk"
        End If
        
        outFilePath = Left(outFilePath, InStrRev(outFilePath, ".") - 1) & suffix & ".igs"
        
    Loop While Dir(outFilePath) <> ""
    
    GetOutputFilePath = outFilePath
    
End Function

Function MergeSelectedSketchSegments(model As SldWorks.ModelDoc2) As SldWorks.Feature
    
    Dim swSketch As SldWorks.sketch
        
    If Not model.SketchManager.ActiveSketch Is Nothing Then
        err.Raise vbError, "", "Close active sketch"
    End If
    
    Dim vSkSegs As Variant
    vSkSegs = GetSelectedSketchSegments(model)
    
    If Not IsEmpty(vSkSegs) Then
    
        model.ClearSelection2 True
        model.SketchManager.Insert3DSketch True
    
        If model.Extension.MultiSelect2(vSkSegs, False, Nothing) = UBound(vSkSegs) + 1 Then
        
            model.SketchManager.SketchUseEdge3 False, False
            
            Set swTargetSketch = model.SketchManager.ActiveSketch
            
            model.SketchManager.ActiveSketch.RelationManager.DeleteAllRelations
                    
            model.SketchManager.Insert3DSketch True
            
            Set MergeSelectedSketchSegments = swTargetSketch
            
        Else
            err.Raise vbError, "", "Failed to select sketches"
        End If
    
    Else
        err.Raise vbError, "", "No sketch segments selected"
    End If
    
End Function

Function GetSelectedSketchSegments(model As SldWorks.ModelDoc2) As Variant
        
    Dim swSketchSegs() As SldWorks.SketchSegment
        
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        Dim objType As swSelectType_e
        objType = swSelMgr.GetSelectedObjectType3(i, -1)
        
        Dim swSelObj As Object
        
        Set swSelObj = swSelMgr.GetSelectedObject6(i, -1)
        
        If objType = swSelectType_e.swSelSKETCHES Then
            
            Dim swFeat As SldWorks.Feature
            Set swFeat = swSelObj
                       
            Dim swSketch As SldWorks.sketch
            Set swSketch = swFeat.GetSpecificFeature2
            
            Dim vSegs As Variant
            vSegs = swSketch.GetSketchSegments
            
            Dim j As Integer
            
            If Not IsEmpty(vSegs) Then
                                
                For j = 0 To UBound(vSegs)
                                
                    If (Not swSketchSegs) = -1 Then
                        ReDim swSketchSegs(0)
                    Else
                        ReDim Preserve swSketchSegs(UBound(swSketchSegs) + 1)
                    End If
                            
                    Set swSketchSegs(UBound(swSketchSegs)) = vSegs(j)
                
                Next
                
            End If
        ElseIf objType = swSelectType_e.swSelEXTSKETCHSEGS Or objType = swSelectType_e.swSelSKETCHSEGS Then
            Dim swSkSeg As SldWorks.SketchSegment
            Set swSkSeg = swSelObj
            If (Not swSketchSegs) = -1 Then
                ReDim swSketchSegs(0)
            Else
                ReDim Preserve swSketchSegs(UBound(swSketchSegs) + 1)
            End If
                    
            Set swSketchSegs(UBound(swSketchSegs)) = swSkSeg
        End If
    Next
    
    If (Not swSketchSegs) = -1 Then
        GetSelectedSketchSegments = Empty
    Else
        GetSelectedSketchSegments = swSketchSegs
    End If
        
End Function

Sub ExportSketchToIges(model As SldWorks.ModelDoc2, sketch As SldWorks.Feature, outFilePath As String)

    If False <> sketch.Select2(False, -1) Then
        
        model.EditCopy
        
        Dim partTemplate As String
        partTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
        
        If partTemplate = "" Then
            err.Raise vbError, "", "Failed to find the default part template"
        End If
        
        Dim swPart As SldWorks.ModelDoc2
        Set swPart = swApp.NewDocument(partTemplate, swDwgPaperSizes_e.swDwgPapersUserDefined, 0, 0)
        swPart.Paste
        
        Dim errs As Long
        Dim warns As Long
        
        Dim expAsWireFrame As Boolean
        Dim expSketchEnts As Boolean
        
        expAsWireFrame = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swIGESExportAsWireframe)
        expSketchEnts = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swIGESExportSketchEntities)
        
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swIGESExportAsWireframe, True
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swIGESExportSketchEntities, True
        
        Dim expRes As Boolean
        
        expRes = swPart.Extension.SaveAs3(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, Nothing, errs, warns)
        
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swIGESExportAsWireframe, expAsWireFrame
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swIGESExportSketchEntities, expSketchEnts
        
        If False = expRes Then
            err.Raise vbError, "", "Failed to export file"
        End If
            
    Else
        err.Raise vbError, "", "Failed to select sketch to export" & sketch.Name
    End If
        
    swApp.CloseDoc swPart.GetTitle()
        
End Sub