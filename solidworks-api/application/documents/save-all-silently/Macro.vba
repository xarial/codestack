Const SHOW_ERROR As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swFrame As SldWorks.Frame
    Set swFrame = swApp.Frame()
    
    Dim vModelWnds As Variant
    vModelWnds = swFrame.ModelWindows
    
    If Not IsEmpty(vModelWnds) Then
        
        Dim i As Integer
        
        Dim savedCount As Integer
        Dim failedCount As Integer
        savedCount = 0
        failedCount = 0
        
        For i = 0 To UBound(vModelWnds)
            
            Dim swModelWnd As SldWorks.ModelWindow
            Set swModelWnd = vModelWnds(i)
            Dim swModel As SldWorks.ModelDoc2
            Set swModel = swModelWnd.ModelDoc
            
            If swModel.GetSaveFlag() Then
                
                Dim errs As Long
                Dim warns As Long
                
                If False = swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errs, warns) Then
                    failedCount = failedCount + 1
                    Debug.Print "Failed to save " & swModel.GetTitle() & ": " & errs
                Else
                    savedCount = savedCount + 1
                    Debug.Print "Saved " & swModel.GetTitle
                End If
                
            End If
            
        Next
        
        swFrame.SetStatusBarText "Saved " & savedCount & " document(s). Failed: " & failedCount & " document(s)"
        
        If failedCount > 0 And SHOW_ERROR Then
            swApp.SendMsgToUser2 "Some of the files failed to save automatically", swMessageBoxIcon_e.swMbWarning, swMessageBoxBtn_e.swMbOk
        End If
        
    End If
    
End Sub