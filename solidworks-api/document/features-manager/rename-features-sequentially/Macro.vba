Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Dim FEATS_FILTER As Variant

Sub main()

    'FEATS_FILTER = Array("Mate*")

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc

try_:
    
    On Error GoTo catch_
    
    If Not swModel Is Nothing Then
        
        swModel.FeatureManager.EnableFeatureTree = False
        swModel.FeatureManager.EnableFeatureTreeWindow = False
        
        Dim vComps As Variant
        
        vComps = GetSelectedComponents(swModel.SelectionManager)
        
        If Not IsEmpty(vComps) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vComps)
                
                Dim swComp As SldWorks.Component2
                Set swComp = vComps(i)
                ProcessFeatureTree swComp.FirstFeature, swComp
                
            Next
        
        Else
            ProcessFeatureTree swModel.FirstFeature, swModel
        End If
        
    Else
        Err.Raise vbError, "", "Please open model"
    End If
    
    GoTo finally_
    
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:
    
    If Not swModel Is Nothing Then
        swModel.FeatureManager.EnableFeatureTree = True
        swModel.FeatureManager.EnableFeatureTreeWindow = True
    End If

End Sub

Sub ProcessFeatureTree(firstFeat As SldWorks.Feature, owner As Object)
    
    Dim passedOrigin As Boolean
    passedOrigin = False

    Dim featNamesTable As Object
    Dim processedFeats() As SldWorks.Feature
    
    Set featNamesTable = CreateObject("Scripting.Dictionary")
        
    featNamesTable.CompareMode = vbTextCompare 'case insensitive
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = firstFeat
    
    While Not swFeat Is Nothing
        
        If passedOrigin Then
        
            If Not Contains(processedFeats, swFeat) Then
                
                If (Not processedFeats) = -1 Then
                    ReDim processedFeats(0)
                Else
                    ReDim Preserve processedFeats(UBound(processedFeats) + 1)
                End If
                
                Set processedFeats(UBound(processedFeats)) = swFeat
        
                RenameFeature swFeat, featNamesTable, owner
            End If
            
            Dim swSubFeat As SldWorks.Feature
            Set swSubFeat = swFeat.GetFirstSubFeature
            
            While Not swSubFeat Is Nothing
                
                If Not Contains(processedFeats, swSubFeat) Then
                    If (Not processedFeats) = -1 Then
                        ReDim processedFeats(0)
                    Else
                        ReDim Preserve processedFeats(UBound(processedFeats) + 1)
                    End If
                    
                    Set processedFeats(UBound(processedFeats)) = swSubFeat
                    RenameFeature swSubFeat, featNamesTable, owner
                End If
                
                Set swSubFeat = swSubFeat.GetNextSubFeature
                
            Wend
        
        End If
        
        If swFeat.GetTypeName2() = "OriginProfileFeature" Then
            passedOrigin = True
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Wend
    
End Sub

Sub RenameFeature(feat As SldWorks.Feature, featNamesTable As Object, owner As Object)

    If MatchesFilter(feat) Then
    
        Dim baseFeatName As String
        
        If TryGetBaseName(feat.name, baseFeatName) Then
            
            Dim nextIndex As Integer
                
            If featNamesTable.Exists(baseFeatName) Then
                nextIndex = featNamesTable.item(baseFeatName) + 1
                featNamesTable.item(baseFeatName) = nextIndex
            Else
                nextIndex = 1
                featNamesTable.Add baseFeatName, nextIndex
            End If
            
            Dim newName As String
            newName = baseFeatName & nextIndex
            
            If LCase(feat.name) <> LCase(newName) Then
            
                ResolveFeatureNameConflict owner, newName
                
                Debug.Print "Renaming '" & feat.name & "' to '" & newName & "'"
                
                feat.name = newName
            
            End If
            
        End If
        
    End If

End Sub

Function MatchesFilter(feat As SldWorks.Feature) As Boolean

    Dim typeName As String
    typeName = feat.GetTypeName2()
    
    If typeName <> "Reference" And typeName <> "ReferencePattern" Then
        If Not IsEmpty(FEATS_FILTER) Then
            Dim i As Integer
            For i = 0 To UBound(FEATS_FILTER)
                If typeName Like CStr(FEATS_FILTER(i)) Then
                    MatchesFilter = True
                    Exit Function
                End If
            Next
            
            MatchesFilter = False
            
        Else
            MatchesFilter = True
        End If
    Else
        MatchesFilter = False
    End If
    
End Function

Function TryGetBaseName(name As String, ByRef baseName As String)
    
    TryGetBaseName = False
    baseName = ""
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "(.+?)(\d+)$"
    
    Dim regExMatches As Object
    Set regExMatches = regEx.Execute(name)
    
    If regExMatches.Count = 1 Then
        
        If regExMatches(0).SubMatches.Count = 2 Then
            
            baseName = regExMatches(0).SubMatches(0)
            TryGetBaseName = True
            
        End If
        
    End If
    
End Function

Sub ResolveFeatureNameConflict(owner As Object, name As String)
    
    Const INDEX_OFFSET As Integer = 100
    Dim index As Integer
    
    Dim swFeatMgr As SldWorks.FeatureManager
    
    Dim swFeat As SldWorks.Feature
        
    If TypeOf owner Is SldWorks.Component2 Then
        
        Dim swComp As SldWorks.Component2
        Set swComp = owner
        
        Dim swRefModel As SldWorks.ModelDoc2
        Set swRefModel = swComp.GetModelDoc2
        
        If Not swRefModel Is Nothing Then
            Set swFeatMgr = swRefModel.FeatureManager
            Set swFeat = swComp.FeatureByName(name)
        Else
            Err.Raise vbError, "", "Component model is not loaded"
        End If
        
    ElseIf TypeOf owner Is SldWorks.ModelDoc2 Then
        
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = owner
        Set swFeatMgr = swModel.FeatureManager
        Set swFeat = swModel.FeatureByName(name)
        
    Else
        Err.Raise vbError, "", "Not supported owner"
    End If
    
    If Not swFeat Is Nothing Then
        
        Dim baseName As String
        
        If TryGetBaseName(name, baseName) Then
            
            Dim newName As String
            newName = baseName & (INDEX_OFFSET + index)
            
            While False <> swFeatMgr.IsNameUsed(swNameType_e.swFeatureName, newName)
                index = index + 1
                newName = baseName & (INDEX_OFFSET + index)
            Wend
            
            swFeat.name = newName
            
        Else
            Exit Sub
        End If
    
    End If
    
End Sub

Function Contains(vArr As Variant, item As Object) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(vArr)
        If vArr(i) Is item Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function

Function GetSelectedComponents(selMgr As SldWorks.SelectionMgr) As Variant

    Dim swComps() As SldWorks.Component2

    Dim i As Integer
    
    For i = 1 To selMgr.GetSelectedObjectCount2(-1)
                
        Dim swComp As SldWorks.Component2
    
        Set swComp = selMgr.GetSelectedObjectsComponent4(i, -1)
        
        If Not swComp Is Nothing Then
            
            If (Not swComps) = -1 Then
                ReDim swComps(0)
                Set swComps(0) = swComp
            Else
                If Not Contains(swComps, swComp) Then
                    ReDim Preserve swComps(UBound(swComps) + 1)
                    Set swComps(UBound(swComps)) = swComp
                End If
            End If
                        
        End If
    
    Next

    If (Not swComps) = -1 Then
        GetSelectedComponents = Empty
    Else
        GetSelectedComponents = swComps
    End If

End Function