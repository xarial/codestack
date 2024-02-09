Dim dirPath
dirPath = WScript.Arguments.Item(0)

Dim filter
filter = WScript.Arguments.Item(1)

Dim outDir
outDir = WScript.Arguments.Item(2)

Dim outExt
outExt = WScript.Arguments.Item(3)

WScript.Echo "Connecting to SOLIDWORKS"

Dim swApp
Set swApp = CreateObject("SldWorks.Application")
swApp.Visible = True

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim folder
Set folder = fso.GetFolder(dirPath)

dim file

For Each file in folder.Files
    If LCase(fso.GetExtensionName(file.Path)) = LCase(filter) Then

        WScript.Echo  "Opening " & file.Path
        Dim docSpec
        Set docSpec = swApp.GetOpenDocSpec(file.Path)
        docSpec.ReadOnly = True

        Dim swModel
        Set swModel = swApp.OpenDoc7(docSpec)

        If Not swModel is Nothing Then
            Dim outFilePath
            outFilePath = outDir & "\" & fso.GetBaseName(file) & "." & outExt
            WScript.Echo "Exporting " & file.Path & " to " & outFilePath
            swModel.SaveAs outFilePath
            swApp.CloseDoc swModel.GetTitle()
        End If
    End If
Next

swApp.ExitApp