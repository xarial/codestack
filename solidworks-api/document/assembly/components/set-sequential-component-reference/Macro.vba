Const INPUT_SEED As Boolean = False
Const SEED As Integer = 1

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swSelMgr As SldWorks.SelectionMgr
        
        Set swSelMgr = swModel.SelectionManager
        
        Dim i As Integer
        
        Dim nextRef As Integer
        
        If INPUT_SEED Then
            Dim seedStr As String
            seedStr = InputBox("Specify the start seed number")
            If seedStr <> "" Then
                nextRef = CInt(seedStr)
            Else
                End
            End If
        Else
            nextRef = SEED
        End If
        
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
            Dim swComp As SldWorks.Component2
            Set swComp = swSelMgr.GetSelectedObjectsComponent3(i, -1)
            
            If swComp Is Nothing Then
                Err.Raise vbError, "", "Object selected at index " & i & " does not belong to component"
            End If
            
            swComp.ComponentReference = nextRef
            
            nextRef = nextRef + 1
            
        Next
        
    Else
        Err.Raise vbError, "", "Open assembly"
    End If
    
End Sub