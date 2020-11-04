Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Const VK_CONTROL As Long = &H11

Public Enum swPlanes_e
    Front = 1
    Top = 2
    Right = 3
End Enum

Const PLANE As Integer = swPlanes_e.Top
Const SCROLL As Boolean = False

Dim swApp As SldWorks.SldWorks
    
Sub main()
               
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Or _
            swModel.GetType() = swDocumentTypes_e.swDocPART Then
            
            SelectPlane swModel, PLANE
            
        Else
            MsgBox "Only assemblies and parts are supported"
        End If
    Else
        MsgBox "Please open part or assembly"
    End If
    
End Sub

Sub SelectPlane(model As SldWorks.ModelDoc2, planeType As swPlanes_e)

    Dim planeIndex As Integer
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature

    Do While Not swFeat Is Nothing

        If swFeat.GetTypeName = "RefPlane" Then

            planeIndex = planeIndex + 1
            
            If CInt(planeType) = planeIndex Then
                
                Dim defScrollState As Boolean
                defScrollState = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swFeatureManagerEnsureVisible)
                swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swFeatureManagerEnsureVisible, SCROLL
                
                Dim append As Boolean
                append = GetKeyState(VK_CONTROL) < 0
                
                swFeat.Select2 append, 0
                
                swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swFeatureManagerEnsureVisible, defScrollState
                
                Exit Sub

            End If

        End If
    
        Set swFeat = swFeat.GetNextFeature

    Loop
    
End Sub