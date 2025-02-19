Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swComp As SldWorks.Component2
    
    Set swComp = swSelMgr.GetSelectedObjectsComponent4(1, -1)
    
    If Not swComp Is Nothing Then
        
        Dim swTransform As SldWorks.MathTransform
        Set swTransform = swComp.Transform2
        Dim vMatrix As Variant
        vMatrix = swTransform.ArrayData
        
        Dim i As Integer
        
        Dim matrixText As String
        
        For i = 0 To UBound(vMatrix)
            If i > 0 Then
                matrixText = matrixText & " "
            End If
            matrixText = matrixText & CDbl(vMatrix(i))
        Next
        
        SetClipboard matrixText
        
    Else
        Err.Raise vbError, "", "Select component"
    End If
    
End Sub

Sub SetClipboard(text As String)

    Dim dataObject As Object
    Set dataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObject.SetText text
    dataObject.PutInClipboard
    Set dataObject = Nothing
    
End Sub