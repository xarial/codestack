#Const ARGS = True
#Const TEST = FALSE

Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Const VK_CONTROL As Long = &H11

Public Enum swRefGeom_e
    Origin = 4
    Front = 1
    Top = 2
    Right = 3
End Enum

Dim REF_GEOM As swRefGeom_e
Const SCROLL As Boolean = False
Const APPEND_SEL As Boolean = False

Dim swApp As SldWorks.SldWorks
    
Sub main()
    
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2

    Set swModel = swApp.ActiveDoc

#If ARGS Then
    
    Dim macroOper As Object
    Set macroOper = GetMacroOperation()
    
    Dim vArgs As Variant
    vArgs = macroOper.Arguments
    
    Dim arg As Object
    Set arg = vArgs(0)
    
    Dim planeName As String
    planeName = arg.GetValue()
    
    Select Case UCase(planeName)
        Case "ORIGIN"
            REF_GEOM = swRefGeom_e.Origin
        Case "TOP"
            REF_GEOM = swRefGeom_e.Top
        Case "FRONT"
            REF_GEOM = swRefGeom_e.Front
        Case "RIGHT"
            REF_GEOM = swRefGeom_e.Right
        Case Else
            Err.Raise vbError, "", "Not supported argument"
    End Select
#Else
    REF_GEOM = swRefGeom_e.Top
#End If
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Or _
            swModel.GetType() = swDocumentTypes_e.swDocPART Then
            
            Dim swSelMgr As SldWorks.SelectionMgr
            Set swSelMgr = swModel.SelectionManager
                        
            Dim swComp As SldWorks.Component2
            Set swComp = swSelMgr.GetSelectedObjectsComponent3(-1, -1)
            
            If swComp Is Nothing Then
                SelectRefGeom swModel.FirstFeature(), REF_GEOM
            Else
                SelectRefGeom swComp.FirstFeature(), REF_GEOM
            End If
            
        Else
            Err.Raise vbError, "", "Only assemblies and parts are supported"
        End If
    Else
        Err.Raise vbError, "", "Please open part or assembly"
    End If
    
End Sub

Sub SelectRefGeom(firstFeat As SldWorks.Feature, refGeomType As swRefGeom_e)

    Dim refGeomIndex As Integer
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = firstFeat

    Do While Not swFeat Is Nothing

        If swFeat.GetTypeName = "RefPlane" Or swFeat.GetTypeName2() = "OriginProfileFeature" Then

            refGeomIndex = refGeomIndex + 1
            
            If CInt(refGeomType) = refGeomIndex Then
                
                Dim defScrollState As Boolean
                defScrollState = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swFeatureManagerEnsureVisible)
                swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swFeatureManagerEnsureVisible, SCROLL
                
                Dim append As Boolean
                
                If APPEND_SEL Then
                    append = True
                Else
                    append = GetKeyState(VK_CONTROL) < 0
                End If
                
                If refGeomType = Origin Then
                    SelectOrigin swFeat, append
                Else
                    swFeat.Select2 append, -1
                End If
                
                swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swFeatureManagerEnsureVisible, defScrollState
                
                Exit Sub

            End If

        End If
    
        Set swFeat = swFeat.GetNextFeature

    Loop
    
End Sub

Sub SelectOrigin(origFeat As SldWorks.Feature, append As Boolean)
    
    Dim swSketch As SldWorks.Sketch
    Set swSketch = origFeat.GetSpecificFeature2
    
    Dim swSkPoint As SldWorks.SketchPoint
    Set swSkPoint = swSketch.GetSketchPoints2()(0)
    
    swSkPoint.Select4 append, Nothing
    
End Sub

Function GetMacroOperation(Optional dummy As Variant = Empty) As Object
    
    Dim macroOper As Object
    
    #If TEST Then
        Dim swCadPlusFact As Object
        Set swCadPlusFact = CreateObject("CadPlusFactory.Sw")
        
        Set swCadPlus = swCadPlusFact.Create(swApp, False)
        
        Dim ARGS(0) As String
        ARGS(0) = "FRONT"
        Set macroOper = swCadPlus.CreateMacroOperation(swApp.ActiveDoc, "", ARGS)
    #Else
        Dim macroOprMgr As Object
        Set macroOprMgr = CreateObject("CadPlus.MacroOperationManager")
        
        Set macroOper = macroOprMgr.PopOperation(swApp)
    #End If
    
    Set GetMacroOperation = macroOper
    
End Function