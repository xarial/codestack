Dim swApp
Set swApp = CreateObject("SldWorks.Application")

Dim filePath
filePath = InputBox("Specify the path to the part file")

Dim docSpec
Set docSpec = swApp.GetOpenDocSpec(filePath)
docSpec.ReadOnly = True
docSpec.Silent = True

Dim swModel
Set swModel = swApp.OpenDoc7(docSpec)

Dim swMassPrps
Set swMassPrps = swModel.Extension.CreateMassProperty()

MsgBox "Mass: " & swMassPrps.Mass & vbLf & "Volume: " & swMassPrps.Volume & vbLf & "Surface area: " & swMassPrps.SurfaceArea

swApp.CloseDoc swModel.GetTitle()