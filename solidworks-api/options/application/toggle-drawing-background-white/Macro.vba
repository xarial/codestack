'This macro toggles between 2 colors for the background of drawings.
'It uses a system option, so every drawing you open will get the choosen color
'This can be usefull if you want to make screen captures on a white background.

'Here you can set the 2 colors you want to toggle between
Const Color1 As Variant = 16777215 'color white
Const Color2 As Variant = 14411494 'color grey (default color for drawing background)


Sub main()

try_:

    On Error GoTo catch_

    Dim swApp As Object
    Set swApp = Application.SldWorks
    
    Dim swModel As ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    'Get the color on first use (look in Immediate window CTRL + G)
    Dim Color As Variant
    Color = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsDrawingsPaper)
    Debug.Print "Color : " + CStr(Color)
    
     
    If Color <> Color1 Then
       Color = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsDrawingsPaper, Color1)
    Else
       Color = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsDrawingsPaper, Color2)
    End If
    
    swModel.ForceRebuild
 
GoTo finally_:
    
catch_:

    Debug.Print "Error: " & Err.Number & ":" & Err.Source & ":" & Err.Description
    
finally_:

    Debug.Print "FINISHED Toggle Drawing Background"
    
End Sub
