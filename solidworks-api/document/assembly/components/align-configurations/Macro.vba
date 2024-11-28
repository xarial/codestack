Const ROOT_CONFS_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swAssy As SldWorks.AssemblyDoc
    
    Set swAssy = swModel
    
    If Not swAssy Is Nothing Then
    
        Dim vComps As Variant
        vComps = GetSelectedRootComponents(swAssy)
        
        If Not IsEmpty(vComps) Then
        
            Dim vConfs As Variant
            vConfs = swModel.GetConfigurationNames
            
            Dim i As Integer
            
            For i = 0 To UBound(vConfs)
                
                Dim swConf As SldWorks.Configuration
                Set swConf = swModel.GetConfigurationByName(CStr(vConfs(i)))
                
                If swConf.GetParent() Is Nothing Or Not ROOT_CONFS_ONLY Then
                                    
                    Dim confParams() As String
                    Dim confParamVals() As String
                    
                    ReDim confParams(UBound(vComps))
                    ReDim confParamVals(UBound(vComps))
                    
                    Dim j As Integer
                           
                    For j = 0 To UBound(vComps)
                
                        Dim swComp As SldWorks.Component2
                        Set swComp = vComps(j)
                        
                        If HasConfiguration(swComp, swConf.Name) Then
                                                   
                           confParams(j) = "$CONFIGURATION@" & GetComponentNameForParameter(swComp)
                           confParamVals(j) = swConf.Name
                        
                        Else
                            Err.Raise vbError, "", swComp.Name2 & " does not contain configuration " & swConf.Name
                        End If
                        
                    Next
                    
                    swConf.SetParameters (confParams), (confParamVals)
                    
                End If
                
            Next
        
        Else
            Err.Raise vbError, "", "Select components to process"
        End If
    
    Else
        Err.Raise vbError, "", "Open assembly"
    End If
    
End Sub

Function GetSelectedRootComponents(assm As SldWorks.AssemblyDoc) As Variant
    
    Dim swComps() As SldWorks.Component2
    
    Dim swSelMgr As SldWorks.SelectionMgr
        
    Set swSelMgr = assm.SelectionManager
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
            
        Dim swComp As SldWorks.Component2
        Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
        
        If Not swComp Is Nothing Then
            
            If swComp.GetParent() Is Nothing Then
            
                If (Not swComps) = -1 Then
                    ReDim swComps(0)
                Else
                    ReDim Preserve swComps(UBound(swComps) + 1)
                End If
                
                Set swComps(UBound(swComps)) = swComp
                
            Else
                Err.Raise vbError, "", "Only top level components are supported"
            End If
            
        End If
        
    Next
    
    If (Not swComps) = -1 Then
        GetSelectedRootComponents = Empty
    Else
        GetSelectedRootComponents = swComps
    End If
    
End Function

Function GetComponentNameForParameter(comp As SldWorks.Component2) As String
    
    Dim instId As Integer
    Dim compName As String
    compName = comp.Name2
    instId = CInt(Right(compName, Len(compName) - InStrRev(compName, "-")))
    compName = Left(compName, InStrRev(compName, "-") - 1)
    
    GetComponentNameForParameter = compName & "<" & instId & ">"
    
End Function

Function HasConfiguration(comp As SldWorks.Component2, confName As String) As Boolean
    
    Dim swRefModel As SldWorks.ModelDoc2
    Set swRefModel = comp.GetModelDoc2
    
    Dim vConfs As Variant
    
    If Not swRefModel Is Nothing Then
        vConfs = swRefModel.GetConfigurationNames
    Else
        vConfs = swApp.GetConfigurationNames(comp.GetPathName())
    End If
    
    HasConfiguration = Contains(vConfs, confName)
    
End Function

Function Contains(vArr As Variant, item As String) As Boolean

    Contains = False

    If Not IsEmpty(vArr) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vArr)
            If LCase(CStr(vArr(i))) = LCase(item) Then
                Contains = True
                Exit Function
            End If
        Next
        
    End If
    
End Function