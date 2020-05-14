Const NAME_TEMPLATE = "SM_{0}x{1}x{2}"
Dim PROPERTIES As Variant

Dim swApp As SldWorks.SldWorks

Sub Init(Optional dummy As Variant = Empty)
    PROPERTIES = Array("Bounding Box Length", "Bounding Box Width", "Sheet Metal Thickness")
End Sub

Sub main()

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
            swCutListFeat.Name = FormatString(NAME_TEMPLATE, vPrpVals)
        Next
        
    Else
        MsgBox "Please open the document"
    End If
    
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
    
    Dim swCutListFeats() As SldWorks.Feature
    Dim isInit As Boolean
    isInit = False
    
    Dim swFeat As SldWorks.Feature
    Dim swBodyFolder As SldWorks.BodyFolder
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        
        If swFeat.GetTypeName2 = "CutListFolder" Then
            
            If Not isInit Then
                isInit = True
                ReDim swCutListFeats(0)
            Else
                ReDim Preserve swCutListFeats(UBound(swCutListFeats) + 1)
            End If
            
            Set swCutListFeats(UBound(swCutListFeats)) = swFeat
            
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
    GetCutLists = swCutListFeats

End Function

Function FormatString(inputStr As String, params As Variant)
    
    Dim resStr As String
    resStr = inputStr
    
    Dim i As Integer
    
    For i = 0 To UBound(params)
        resStr = Replace(resStr, "{" & i & "}", CStr(params(i)))
    Next
    
    FormatString = resStr
    
End Function