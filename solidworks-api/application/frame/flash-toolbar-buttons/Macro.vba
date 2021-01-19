Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    'flash line button and show tooltip
    FlashToolbarButton 32873
    
    'only show tooltip for a new file button
    FlashToolbarButton 57600, True
    
End Sub

Sub FlashToolbarButton(buttonId As Long, Optional tooltipOnly As Boolean = False)
    
    swApp.ShowBubbleTooltip buttonId, IIf(tooltipOnly, "", CStr(buttonId)), 0, "", ""
    
End Sub