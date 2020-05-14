Const SUFFIX As String = "_Renamed"

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr

    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Set swSelMgr = swModel.SelectionManager
        
        Dim swComp As SldWorks.Component2
        
        Set swComp = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swComp Is Nothing Then
        
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
            
            Dim newCompName As String
            newCompName = compName & SUFFIX
            
            swComp.Name2 = newCompName
            
        Else
            MsgBox "Please select component to rename"
        End If
    
    Else
        MsgBox "Please open assembly document"
    End If
    
End Sub