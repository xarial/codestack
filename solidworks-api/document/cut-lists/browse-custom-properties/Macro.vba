#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    Dim swBody As SldWorks.Body2
    
    Set swBody = GetSelectedObjectBody(swPart, 1)
    
    Dim swCutListFeat As SldWorks.Feature
    
    Set swCutListFeat = GetCutListFromBody(swPart, swBody)
    
    ShowCutListPropertiesDialog swCutListFeat
    
End Sub

Function GetSelectedObjectBody(model As SldWorks.ModelDoc2, index As Integer) As SldWorks.Body2
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = model.SelectionManager
    
    Dim swSelObj As Object
    
    Set swSelObj = swSelMgr.GetSelectedObject6(index, -1)
    
    Dim swBody As SldWorks.Body2
    
    If Not swSelObj Is Nothing Then
        
        Select Case swSelMgr.GetSelectedObjectType3(index, -1)
            Case swSelectType_e.swSelSOLIDBODIES
                Set swBody = swSelObj
                
            Case swSelectType_e.swSelFACES
                Dim swFace As SldWorks.Face2
                Set swFace = swSelObj
                Set swBody = swFace.GetBody
            
            Case swSelectType_e.swSelEDGES
                Dim swEdge As SldWorks.Edge
                Set swEdge = swSelObj
                Set swBody = swEdge.GetBody
                
            Case swSelectType_e.swSelVERTICES
                Dim swVert As SldWorks.Vertex
                Set swVert = swSelObj
                Set swBody = swVert.GetEdges()(0).GetBody()
                
            Case swSelectType_e.swSelBODYFEATURES
                Dim swFeat As SldWorks.Feature
                Set swFeat = swSelObj
                Set swBody = swFeat.GetFaces()(0).GetBody
                
            Case Else
                Err.Raise vbError, "", "Not supported"
        End Select
        
    Else
        Err.Raise vbError, "", "Select entity"
    End If
    
    Set GetSelectedObjectBody = swBody

End Function

Function GetCutListFromBody(model As SldWorks.ModelDoc2, body As SldWorks.Body2) As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    Dim swBodyFolder As SldWorks.BodyFolder
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        
        If swFeat.GetTypeName2 = "CutListFolder" Then
            
            Set swBodyFolder = swFeat.GetSpecificFeature2
            
            Dim vBodies As Variant
            
            vBodies = swBodyFolder.GetBodies
            
            Dim i As Integer
            
            If Not IsEmpty(vBodies) Then
                For i = 0 To UBound(vBodies)
                    
                    Dim swCutListBody As SldWorks.Body2
                    Set swCutListBody = vBodies(i)
                    
                    If swApp.IsSame(swCutListBody, body) = swObjectEquality.swObjectSame Then
                        Set GetCutListFromBody = swFeat
                        Exit Function
                    End If
                    
                Next
            End If
            
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
    Err.Raise vbError, "", "Failed to find cut-list from body"

End Function

Sub ShowCutListPropertiesDialog(cutListFeat As SldWorks.Feature)
    
    If False <> cutListFeat.Select2(False, -1) Then
        Const CMD_ShowProperties As Long = 51482
        Const WM_COMMAND As Long = &H111
        SendMessage swApp.Frame().GetHWnd(), WM_COMMAND, CMD_ShowProperties, 0
    Else
        Err.Raise vbError, "", "Failed to select cut-list feature"
    End If
    
End Sub