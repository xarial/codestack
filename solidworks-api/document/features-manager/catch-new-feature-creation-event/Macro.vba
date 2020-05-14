Dim swApp As SldWorks.SldWorks
Dim swEventListener As EventListener

Sub main()

    Set swApp = Application.SldWorks
    
    Set swEventListener = New EventListener
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    swEventListener.SetPart swPart
    
    While swApp.ActiveDoc Is swPart
        DoEvents
    Wend
    
End Sub
