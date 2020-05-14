Dim swApp As SldWorks.SldWorks
Dim swEventsListener As EventsListener

Sub main()

    Set swApp = Application.SldWorks

    Set swEventsListener = New EventsListener
        
    Dim swAssy As SldWorks.AssemblyDoc
    
    Set swAssy = swApp.ActiveDoc
    
    swEventsListener.SetAssembly swAssy
    
    While swApp.ActiveDoc Is swAssy
        DoEvents
    Wend
        
End Sub
