Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
try_:
    On Error GoTo catch_
    
    If Not swModel Is Nothing Then
    
        If swModel.GetType() <> swDocumentTypes_e.swDocASSEMBLY Then
            Err.Raise vbError, "", "Active document is not an assembly"
        End If
        
        If False = swModel.IsOpenedViewOnly Then
            Err.Raise vbError, "", "Active assembly is not opened in Large Design Review mode"
        End If
        
        Dim vComps As Variant
        
        vComps = GetSelectedComponents(swModel)
        
        OpenComponents swModel, vComps
        
        GoTo finally_
        
    Else
        Err.Raise vbError, "", "Please open assembly document"
    End If

catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:
 
End Sub

Sub OpenComponents(model As SldWorks.ModelDoc2, vComps As Variant)
    
    If Not IsEmpty(vComps) Then
            
        Dim i As Integer
        
        For i = 0 To UBound(vComps)
            
            Dim swComp As SldWorks.Component2
            Set swComp = vComps(i)
        
            Dim compPath As String
            compPath = GetComponentPath(model, swComp)
            
            Dim swDocSpec As SldWorks.DocumentSpecification
            Set swDocSpec = swApp.GetOpenDocSpec(compPath)
            
            swDocSpec.ViewOnly = True
            
            Dim swRefModel As SldWorks.ModelDoc2
            Set swRefModel = swApp.OpenDoc7(swDocSpec)
            
            If swRefModel Is Nothing Then
                Err.Raise vbError, "", "Failed to open component. Error code: " & swDocSpec.Error
            End If
            
        Next
        
    Else
        Err.Raise vbError, "", "No component selected"
    End If
    
End Sub

Function GetSelectedComponents(model As SldWorks.ModelDoc2) As Variant
    
    Dim swComps() As SldWorks.Component2
    Dim isArrInit As Boolean
    isArrInit = False
    
    Dim i As Integer
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        Dim swComp As SldWorks.Component2
        Set swComp = swSelMgr.GetSelectedObject6(i, -1)
        
        If Not swComp Is Nothing Then
            
            Dim unique As Boolean
            unique = False
            
            If False = isArrInit Then
                isArrInit = True
                ReDim swComps(0)
                unique = True
            Else
                unique = Not Contains(swComps, swComp)
                If True = unique Then
                    ReDim Preserve swComps(UBound(swComps) + 1)
                End If
            End If
                
            If True = unique Then
                Set swComps(UBound(swComps)) = swComp
            End If
            
        End If
        
    Next
    
    If isArrInit Then
        GetSelectedComponents = swComps
    Else
        GetSelectedComponents = Empty
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

Function GetComponentPath(assm As SldWorks.AssemblyDoc, comp As SldWorks.Component2) As String
    
    Dim pathParts As Variant
    pathParts = Split(comp.GetPathName(), "\")
    
    Dim assmFolder As String
    assmFolder = assm.GetPathName
    assmFolder = Left(assmFolder, InStrRev(assmFolder, "\") - 1)

    Dim i As Integer
    
    Dim curRelPath As String
    
    For i = UBound(pathParts) To 1 Step -1
        
        curRelPath = pathParts(i) & IIf(curRelPath <> "", "\", "") & curRelPath
        Dim path As String
        path = assmFolder & "\" & curRelPath
        
        If Dir(path) <> "" Then
            GetComponentPath = path
            Exit Function
        End If
        
    Next
    
    GetComponentPath = comp.GetPathName
    
End Function