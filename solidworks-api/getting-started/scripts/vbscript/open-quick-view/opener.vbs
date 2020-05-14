Dim swApp
Set swApp = CreateObject("SldWorks.Application")
swApp.Visible = True

Dim filePath
filePath = WScript.Arguments.Item(0)

If filePath <> "" then

	Dim docSpec
	Set docSpec = swApp.GetOpenDocSpec(filePath)
	docSpec.ViewOnly = True

	Dim swModel
	Set swModel = swApp.OpenDoc7(docSpec)

	If swModel is Nothing Then
		MsgBox "Failed to open document"
	End If
	
Else
	MsgBox "File path is not specified"
End If