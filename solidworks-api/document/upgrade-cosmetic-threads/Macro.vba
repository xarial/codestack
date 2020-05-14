Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks

    Dim allowUpgrade As Boolean
    allowUpgrade = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swEnableAllowCosmeticThreadsUpgrade)

try:
    On Error GoTo catch
    
    Set swModel = swApp.ActiveDoc
        
    If Not swModel Is Nothing Then
                
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swEnableAllowCosmeticThreadsUpgrade, True
        
        If False = swModel.Extension.UpgradeLegacyCThreads() Then
            Debug.Print "Thread is not upgraded"
        End If
            
    Else
        Err.Raise vbError, "", "Please open document"
    End If
    
    GoTo finally
    
catch:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally:
    
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swEnableAllowCosmeticThreadsUpgrade, allowUpgrade

End Sub