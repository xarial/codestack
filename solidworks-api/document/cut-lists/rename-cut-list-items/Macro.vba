Const NAME_TEMPLATE = "SM_{0}x{1}x{2}"
Dim PROPERTIES As Variant

Dim swApp As SldWorks.SldWorks

Sub Init(Optional dummy As Variant = Empty)
    PROPERTIES = Array("Bounding Box Length", "Bounding Box Width", "Sheet Metal Thickness")
End Sub

Sub main()

try_:
    On Error GoTo catch_
    
    Init
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim vCutLists As Variant
        vCutLists = GetCutLists(swModel)
        
        Dim i As Integer
        
        For i = 0 To UBound(vCutLists)
            
            Dim swCutListFeat As SldWorks.Feature
            Set swCutListFeat = vCutLists(i)
            
            Dim vPrpVals As Variant
            vPrpVals = ReadProperties(swCutListFeat.CustomPropertyManager, PROPERTIES)
            
            Dim featName As String
            
            featName = FormatString(NAME_TEMPLATE, vPrpVals)

            If swCutListFeat.Name <> featName Then
                                
                If featName <> "" Then
                
                    Dim index As Integer
                    index = 0
                    
                    While swModel.FeatureManager.IsNameUsed(swNameType_e.swFeatureName, featName)
                        index = index + 1
                        featName = FormatString(NAME_TEMPLATE, vPrpVals) + CStr(index)
                    Wend
                    
                    swCutListFeat.Name = featName
                Else
                    Debug.Print "Empty name for " & swCutListFeat.Name
                End If
            End If
            
        Next
        
    Else
        MsgBox "Please open the document"
    End If
    
    GoTo finally_

catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Function ReadProperties(custPrpMgr As SldWorks.CustomPropertyManager, prpNames As Variant) As Variant
    
    Dim prpValues() As String
    
    ReDim prpValues(UBound(prpNames))
    
    Dim i As Integer
    
    For i = 0 To UBound(prpNames)
        Dim resVal As String
        custPrpMgr.Get2 CStr(prpNames(i)), "", resVal
        prpValues(i) = resVal
    Next
    
    ReadProperties = prpValues
    
End Function

Function GetCutLists(model As SldWorks.ModelDoc2) As Variant
    
    GetCutLists = GetFeaturesByType(model, "CutListFolder")

End Function

Function GetFeaturesByType(model As SldWorks.ModelDoc2, typeName As String) As Variant
    
    Dim swFeats() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        
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

Function FormatString(inputStr As String, params As Variant)
    
    Dim resStr As String
    resStr = inputStr
    
    Dim i As Integer
    
    For i = 0 To UBound(params)
        resStr = Replace(resStr, "{" & i & "}", CStr(params(i)))
    Next
    
    FormatString = resStr
    
End Function