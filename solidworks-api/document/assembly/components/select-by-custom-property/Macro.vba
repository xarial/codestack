Type PropertyFilter
    PropertyName As String
    Patterns As Variant
    Inclusive As Boolean
End Type

Const FILE_TYPE As Integer = swDocumentTypes_e.swDocPART

Dim PROPERTY_FILTERS() As PropertyFilter

Dim swApp As SldWorks.SldWorks

Sub main()

    ReDim PROPERTY_FILTERS(1)
    
    PROPERTY_FILTERS(0).PropertyName = "Type"
    PROPERTY_FILTERS(0).Patterns = Array(".*MadeToStock.*", ".*PurchasedToStock.*")
    PROPERTY_FILTERS(0).Inclusive = True
    
    PROPERTY_FILTERS(1).PropertyName = "StockNumber"
    PROPERTY_FILTERS(1).Patterns = Array(".+")
    PROPERTY_FILTERS(1).Inclusive = False

    Set swApp = Application.SldWorks
    
    Dim swAssy As SldWorks.AssemblyDoc
    
    Set swAssy = swApp.ActiveDoc
    
    If Not swAssy Is Nothing Then
        
        Dim vComps As Variant
        vComps = FilterComponents(swAssy)
        
        If Not IsEmpty(vComps) Then
        
            Dim swModel As SldWorks.ModelDoc2
            Set swModel = swAssy
            
            If UBound(vComps) + 1 <> swModel.Extension.MultiSelect2(vComps, False, Nothing) Then
                Err.Raise vbError, , "Failed to select components"
            End If
        
        End If
    
    Else
        Err.Raise vbError, "", "Open assembly"
    End If
    
End Sub

Function FilterComponents(assy As SldWorks.AssemblyDoc) As Variant
    
    Dim vComps As Variant
    vComps = assy.GetComponents(False)
    
    Dim swFilteredComps() As SldWorks.Component2
        
    Dim i As Integer
    
    For i = 0 To UBound(vComps)
        
        Dim swComp As SldWorks.Component2
        Set swComp = vComps(i)
        
        If IsFiltered(swComp) Then
        
            If (Not swFilteredComps) = -1 Then
                ReDim swFilteredComps(0)
            Else
                ReDim Preserve swFilteredComps(UBound(swFilteredComps) + 1)
            End If
                    
            Set swFilteredComps(UBound(swFilteredComps)) = swComp
            
        End If
        
    Next
    
    If (Not swFilteredComps) = -1 Then
        FilterComponents = Empty
    Else
        FilterComponents = swFilteredComps
    End If
    
End Function

Function IsFiltered(comp As SldWorks.Component2) As Boolean
    
    If False = comp.IsSuppressed() Then
        Dim swRefModel As SldWorks.ModelDoc2
        Set swRefModel = comp.GetModelDoc2
        
        If Not swRefModel Is Nothing Then
            If swRefModel.GetType() = FILE_TYPE Then
                                
                Dim i As Integer
                
                For i = 0 To UBound(PROPERTY_FILTERS)
                    Dim prpFilter As PropertyFilter
                    prpFilter = PROPERTY_FILTERS(i)
                    
                    If prpFilter.Inclusive <> MatchesAnyPropertyValue(swRefModel, prpFilter.PropertyName, comp.ReferencedConfiguration, prpFilter.Patterns) Then
                        IsFiltered = False
                        Exit Function
                    End If
                    
                Next
                
                IsFiltered = True
                
            Else
                IsFiltered = False
            End If
            
        Else
            Err.Raise vbError, "", "Referenced model of '" & comp.Name2 & "' is not loaded. Make sure component is not lightweight"
        End If
        
    Else
        IsFiltered = False
    End If
    
End Function

Function MatchesAnyPropertyValue(model As SldWorks.ModelDoc2, prpName As String, confName As String, vals As Variant) As Boolean
        
    Dim prpVal As String
                
    prpVal = GetCustomPropertyValue(model, prpName, confName)
    
    If prpVal = "" Then
        prpVal = GetCustomPropertyValue(model, prpName, "")
    End If
    
    Dim i As Integer
    
    For i = 0 To UBound(vals)
        If IsMatch(prpVal, CStr(vals(i))) Then
            MatchesAnyPropertyValue = True
            Exit Function
        End If
    Next
    
    MatchesAnyPropertyValue = False
        
End Function

Function GetCustomPropertyValue(model As SldWorks.ModelDoc2, prpName As String, confName As String) As String

    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)

    Dim prpVal As String
    Dim prpResVal As String
    Dim wasResolved As Boolean
    Dim isLinked As Boolean
    
    Dim res As Long
    res = swCustPrpMgr.Get6(prpName, cached, prpVal, prpResVal, wasResolved, isLinked)
    
    GetCustomPropertyValue = prpResVal
    
End Function

Function IsMatch(text As String, pattern As String) As Boolean

    If pattern = "" Then
        IsMatch = text = ""
        Exit Function
    End If

    With CreateObject("VBScript.RegExp")
        .pattern = pattern
        .ignorecase = True
        
        If .Test(text) Then
            IsMatch = True
        Else
            IsMatch = False
        End If
        
    End With

End Function