Const WIDTH As Double = 0.01
Const LENGTH As Double = 0.01
Const HEIGHT As Double = 0.01

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    If Not swPart Is Nothing Then
    
        Dim swModeler As SldWorks.Modeler
        Set swModeler = swApp.GetModeler
        
        Dim dCenter(2) As Double
        dCenter(0) = 0: dCenter(1) = 0: dCenter(2) = 0
        
        Dim dAxis(2) As Double
        dAxis(0) = 0: dAxis(1) = 0: dAxis(2) = 1
                        
        Dim dBoxData(8) As Double
        dBoxData(0) = dCenter(0): dBoxData(1) = dCenter(1): dBoxData(2) = dCenter(2)
        dBoxData(3) = dAxis(0): dBoxData(4) = dAxis(1): dBoxData(5) = dAxis(2)
        dBoxData(6) = WIDTH: dBoxData(7) = LENGTH: dBoxData(8) = HEIGHT
        
        Dim swBody As SldWorks.Body2
        
        Set swBody = swModeler.CreateBodyFromBox3(dBoxData)
        
        swBody.Display3 swPart, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectable
        
        Stop 'continue to hide the body
        
        Set swBody = Nothing
    Else
        MsgBox "Please open part document"
    End If
    
End Sub