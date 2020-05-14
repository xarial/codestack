Const VIRTUAL_ONLY As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr

    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
        
            Dim swAssy As SldWorks.AssemblyDoc
            Set swAssy = swModel
            
            Dim vComps As Variant
            vComps = swAssy.GetComponents(False)
            
            Dim i As Integer
            
            For i = 0 To UBound(vComps)
                
                Dim swComp As SldWorks.Component2
                Set swComp = vComps(i)
                
                If swComp.IsVirtual Or Not VIRTUAL_ONLY Then
                
                    Dim compName As String
                    
                    compName = swComp.Name2
                    
                    If Not swComp.GetParent() Is Nothing Then
                        'if not root remove the sub-assemblies name
                        compName = Right(compName, Len(compName) - InStrRev(compName, "/"))
                    End If
                    
                    If swComp.IsVirtual() Then
                        'if virtual remove the context assembly name
                        compName = Left(compName, InStr(compName, "^") - 1)
                    Else
                        'remove the index name
                        compName = Left(compName, InStrRev(compName, "-") - 1)
                    End If
                    
                    swComp.ComponentReference = compName
                
                End If
                
            Next

        Else
            MsgBox "Active document is not an assembly"
        End If
    
    Else
        MsgBox "Please open assembly document"
    End If
    
End Sub