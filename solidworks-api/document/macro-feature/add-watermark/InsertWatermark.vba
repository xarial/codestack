Const WATERMARK_MACRO_NAME As String = "Watermark.swp"
Const SECURITY_NOTE As String = "www.codestack.net"

Const BASE_NAME As String = "Watermark"

Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim watermarkMacroPath As String
        watermarkMacroPath = swApp.GetCurrentMacroPathFolder() & "\" & WATERMARK_MACRO_NAME
        Dim vMethods(8) As String
        Dim moduleName As String
        
        GetMacroEntryPoint swApp, watermarkMacroPath, moduleName
        
        vMethods(0) = watermarkMacroPath: vMethods(1) = moduleName: vMethods(2) = "swmRebuild"
        vMethods(3) = watermarkMacroPath: vMethods(4) = moduleName: vMethods(5) = "swmEditDefinition"
        vMethods(6) = watermarkMacroPath: vMethods(7) = moduleName: vMethods(8) = "swmSecurity"
        
        Dim swFeat As SldWorks.Feature
        
        Set swFeat = swModel.FeatureManager.InsertMacroFeature3(BASE_NAME, "", vMethods, _
            Empty, Empty, Empty, Empty, Empty, Empty, _
            Empty, swMacroFeatureOptions_e.swMacroFeatureEmbedMacroFile + swMacroFeatureOptions_e.swMacroFeatureAlwaysAtEnd)
        
        If Not swFeat Is Nothing Then
            Dim swSecNote As SldWorks.note
            If SECURITY_NOTE <> "" Then
                Set swSecNote = swModel.FeatureManager.InsertSecurityNote(SECURITY_NOTE, swFeat)
            End If
        Else
            MsgBox "Failed to create watermark feature"
        End If
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub

Sub GetMacroEntryPoint(app As SldWorks.SldWorks, macroPath As String, ByRef moduleName As String)
        
    Dim vMethods As Variant
    vMethods = app.GetMacroMethods(macroPath, swMacroMethods_e.swAllMethods)
    
    Dim i As Integer
    
    If Not IsEmpty(vMethods) Then
    
        For i = 0 To UBound(vMethods)
            Dim vData As Variant
            vData = Split(vMethods(i), ".")
            
            If i = 0 Or LCase(vData(1)) = "swmRebuild" Then
                moduleName = vData(0)
            End If
        Next
        
    End If
    
End Sub