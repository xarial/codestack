Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    Dim addToDbOrig As Boolean
    
    addToDbOrig = swModel.SketchManager.AddToDB 'get the original value
    swModel.SketchManager.AddToDB = True
    
    swModel.SketchManager.CreateLine 0, 0, 0, 0.01, 0.02, 0

    swModel.SketchManager.AddToDB = addToDbOrig 'restore the original value
    
End Sub