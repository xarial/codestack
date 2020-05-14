Const FILE_NAME As String = "SimpleBox.SLDPRT"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Dim path As String
    path = swApp.GetCurrentMacroPathFolder() & "\" & FILE_NAME
    
    Set swModel = swApp.GetOpenDocumentByName(path)
    
    Dim wasVisible As Boolean
    
    If Not swModel Is Nothing Then
        wasVisible = swModel.Visible
    End If
    
    Set swModel = swApp.OpenDoc6(path, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
    
    If Not swModel Is Nothing Then
        swApp.ActivateDoc3 swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, 0
    End If
    
    MsgBox "Was Visible: " & wasVisible
    
    If False = wasVisible Then
        swApp.CloseDoc swModel.GetTitle
    End If
    
End Sub

