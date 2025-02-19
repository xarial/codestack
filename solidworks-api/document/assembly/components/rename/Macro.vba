Const SUFFIX As String = "_Renamed"
Const PRP_NAME As String = ""

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
    
    If False <> swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefUpdateCompNames) Then
        Err.Raise vbError, "", "Uncheck option System Options->External References->Updated component names when documents are replaced"
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr

    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim vComps As Variant
        vComps = GetSelectedComponents(swModel)
        
        If Not IsEmpty(vComps) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vComps)
            
                Dim swComp As SldWorks.Component2
            
                Set swComp = vComps(i)
            
                Dim newCompName As String
                newCompName = GetNewComponentName(swComp)
                
                Dim compInst As Integer
                compInst = GetComponentInstance(swComp)
                
                If False <> swComp.Select4(False, Nothing, False) Then
                    swComp.Name2 = newCompName
                    
                    If LCase(swComp.Name2) <> LCase(newCompName & "-" & compInst) Then
                        Err.Raise vbError, "", "Failed to rename component '" & swComp.Name2 & "' to '" & newCompName & "'"
                    End If
                Else
                    Err.Raise vbError, "", "Failed to select component '" & swComp.Name2 & "'"
                End If
                
            Next
        
        Else
            Err.Raise vbError, "", "Please select components to rename"
        End If
    
    Else
        MsgBox "Please open assembly document"
    End If
    
End Sub

Function GetNewComponentName(comp As SldWorks.Component2) As String

    Dim compName As String
    
    If PRP_NAME <> "" Then
        compName = GetCustomPropertyValue(comp, PRP_NAME)
        
        If compName = "" Then
            Err.Raise vbError, "", "Failed to get custom proeprty value from '" & comp.Name2 & "'"
        End If
    Else
        compName = comp.Name2
        
        If Not comp.GetParent() Is Nothing Then
            'if not root remove the sub-assemblies name
            compName = Right(compName, Len(compName) - InStrRev(compName, "/"))
        End If
        
        If comp.IsVirtual() Then
            'if virtual remove the context assembly name
            compName = Left(compName, InStr(compName, "^") - 1)
        Else
            'remove the index name
            compName = Left(compName, InStrRev(compName, "-") - 1)
        End If
    End If
    
    Dim newCompName As String
    newCompName = compName & SUFFIX
    
    GetNewComponentName = newCompName

End Function

Function GetComponentInstance(comp As SldWorks.Component2) As Integer

    Dim instId As Integer
    Dim compName As String
    compName = comp.Name2
    instId = CInt(Right(compName, Len(compName) - InStrRev(compName, "-")))
    
    GetComponentInstance = instId
    
End Function

Function GetCustomPropertyValue(comp As SldWorks.Component2, prpName As String) As String

    Dim swRefDoc As SldWorks.ModelDoc2
    Set swRefDoc = comp.GetModelDoc2
    
    If Not swRefDoc Is Nothing Then
        Dim prpVal As String
        
        swRefDoc.Extension.CustomPropertyManager(comp.ReferencedConfiguration).Get3 prpName, False, "", prpVal
        
        If prpVal = "" Then
            swRefDoc.Extension.CustomPropertyManager("").Get3 prpName, False, "", prpVal
        End If
        
        GetCustomPropertyValue = prpVal
        
    Else
        Err.Raise vbError, "", "Reference document of '" & swComp.Name2 & "' is not loaded"
    End If

End Function

Function GetSelectedComponents(model As SldWorks.ModelDoc2) As Variant
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    Dim swComps() As SldWorks.Component2

    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
                
        Dim swComp As SldWorks.Component2
    
        Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
        
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