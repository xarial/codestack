Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim curVal As Boolean
    curVal = False <> swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference)
    
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, Not curVal
    
End Sub