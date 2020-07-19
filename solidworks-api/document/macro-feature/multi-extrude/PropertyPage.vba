Implements PropertyManagerPage2Handler9

Public Event Closed(mode As Integer, vSketches As Variant, vDepths As Variant, isCancelled As Boolean)
Public Event DataChanged(vSketches As Variant, vDepths As Variant)

Dim swPage As PropertyManagerPage2
Dim swGroupBox() As PropertyManagerPageGroup
Dim swSelectionBox() As PropertyManagerPageSelectionbox
Dim swNumberBox() As PropertyManagerPageNumberbox

Const EXTRUDES_COUNT As Integer = 5

Const GroupStartID As Long = 1
Const SelectionBoxStartID As Long = GroupStartID + EXTRUDES_COUNT
Const NumberBoxStartID As Long = SelectionBoxStartID + EXTRUDES_COUNT

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Dim vSelSketches As Variant
Dim vDepthVals As Variant
Dim IsCancel As Boolean
Dim PageMode As Integer

Sub Show(Optional mode As Integer = -1, Optional vSketches As Variant = Empty, Optional vDepths As Variant = Empty)
        
    PageMode = mode
        
    Set swApp = Application.SldWorks
    
    CreatePage
    
    InitPageValues vSketches, vDepths
    
    Const swPropertyManagerPageShowOptions_Default As Integer = 0
    
    swPage.Show2 swPropertyManagerPageShowOptions_Default
    
    Set swModel = swApp.ActiveDoc
        
End Sub

Sub CreatePage()

    Dim errs As Long
    Set swPage = swApp.CreatePropertyManagerPage("MultiBoss-Extrude", _
        swPropertyManager_OkayButton + swPropertyManager_CancelButton, Me, errs)
    
    If Not swPage Is Nothing Then
        
        Dim i As Integer
            
        Dim selMark As Integer
        
        ReDim swGroupBox(EXTRUDES_COUNT - 1)
        ReDim swSelectionBox(EXTRUDES_COUNT - 1)
        ReDim swNumberBox(EXTRUDES_COUNT - 1)
        
        For i = 0 To EXTRUDES_COUNT - 1
            
            Dim grpOtps As Integer
            grpOtps = swAddGroupBoxOptions_e.swGroupBoxOptions_Visible + swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox
            
            If i = 0 Then
                grpOtps = grpOtps + swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded
            End If
            
            Set swGroupBox(i) = swPage.AddGroupBox(GroupStartID + i, "Extrude" & i + 1, grpOtps)
            
            Set swSelectionBox(i) = swGroupBox(i).AddControl2(SelectionBoxStartID + i, _
                swPropertyManagerPageControlType_e.swControlType_Selectionbox, "Region", _
                swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, _
                swAddControlOptions_e.swControlOptions_Enabled + swAddControlOptions_e.swControlOptions_Visible, "Select region to extrude")
    
            Dim filters(0) As Long
            filters(0) = swSelectType_e.swSelSKETCHES
    
            swSelectionBox(i).SingleEntityOnly = True
            swSelectionBox(i).Height = 30
            swSelectionBox(i).SetSelectionFilters filters
            swSelectionBox(i).Mark = 2 ^ i
            
            Set swNumberBox(i) = swGroupBox(i).AddControl2(NumberBoxStartID + i, _
                  swPropertyManagerPageControlType_e.swControlType_Numberbox, "Extrude Depth", _
                  swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, _
                  swAddControlOptions_e.swControlOptions_Enabled + swAddControlOptions_e.swControlOptions_Visible, "")
                  
            swNumberBox(i).SetRange2 swNumberboxUnitType_e.swNumberBox_Length, 0.00001, 1000, False, 0.01, 0.1, 0.001
            swNumberBox(i).Value = 0.01
            
        Next
        
    End If

End Sub

Sub InitPageValues(vSketches As Variant, vDepths As Variant)
    
    If Not IsEmpty(vSketches) Then
        
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = swApp.ActiveDoc
        swModel.ClearSelection2 True
        
        Dim i As Integer
        
        For i = 0 To UBound(vSketches)
            Dim swSketchFeat As SldWorks.Feature
            Set swSketchFeat = vSketches(i)
            swSketchFeat.SelectByMark True, 2 ^ i
            swGroupBox(i).Checked = True
            swNumberBox(i).Value = CDbl(vDepths(i))
        Next
        
    End If
    
End Sub

Sub PropertyManagerPage2Handler9_AfterActivation()
End Sub

Sub PropertyManagerPage2Handler9_AfterClose()
    Set swPage = Nothing
    RaiseEvent Closed(PageMode, vSelSketches, vDepthVals, IsCancel)
End Sub

Function PropertyManagerPage2Handler9_OnActiveXControlCreated(ByVal Id As Long, ByVal Status As Boolean) As Long
    PropertyManagerPage2Handler9_OnActiveXControlCreated = 0
End Function

Sub PropertyManagerPage2Handler9_OnButtonPress(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnCheckboxCheck(ByVal Id As Long, ByVal Checked As Boolean)
End Sub

Sub PropertyManagerPage2Handler9_OnClose(ByVal Reason As Long)
    
    IsCancel = Reason = swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel
        
    If Not IsCancel Then
        CollectData vSelSketches, vDepthVals
    End If
    
End Sub

Sub PropertyManagerPage2Handler9_OnComboboxEditChanged(ByVal Id As Long, ByVal Text As String)
End Sub

Sub PropertyManagerPage2Handler9_OnComboboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)
End Sub
 
Sub PropertyManagerPage2Handler9_OnGroupCheck(ByVal Id As Long, ByVal Checked As Boolean)
    HandleDataChanged
End Sub

Sub PropertyManagerPage2Handler9_OnGroupExpand(ByVal Id As Long, ByVal Expanded As Boolean)
End Sub

Function PropertyManagerPage2Handler9_OnHelp() As Boolean
    PropertyManagerPage2Handler9_OnHelp = True
End Function

Function PropertyManagerPage2Handler9_OnKeystroke(ByVal Wparam As Long, ByVal Message As Long, ByVal Lparam As Long, ByVal Id As Long) As Boolean
End Function

Sub PropertyManagerPage2Handler9_OnListboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)
End Sub

Function PropertyManagerPage2Handler9_OnNextPage() As Boolean
    PropertyManagerPage2Handler9_OnNextPage = True
End Function

Sub PropertyManagerPage2Handler9_OnNumberboxChanged(ByVal Id As Long, ByVal Value As Double)
    HandleDataChanged
End Sub

Sub PropertyManagerPage2Handler9_OnOptionCheck(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnPopupMenuItem(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnPopupMenuItemUpdate(ByVal Id As Long, retVal As Long)
End Sub

Function PropertyManagerPage2Handler9_OnPreview() As Boolean
    PropertyManagerPage2Handler9_OnPreview = True
End Function

Function PropertyManagerPage2Handler9_OnPreviousPage() As Boolean
    PropertyManagerPage2Handler9_OnPreviousPage = True
End Function

Sub PropertyManagerPage2Handler9_OnRedo()
End Sub

Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutCreated(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnSelectionboxCalloutDestroyed(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnSelectionboxFocusChanged(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnSelectionboxListChanged(ByVal Id As Long, ByVal Count As Long)
    HandleDataChanged
End Sub

Sub PropertyManagerPage2Handler9_OnSliderPositionChanged(ByVal Id As Long, ByVal Value As Double)
End Sub

Sub PropertyManagerPage2Handler9_OnSliderTrackingCompleted(ByVal Id As Long, ByVal Value As Double)
End Sub

Function PropertyManagerPage2Handler9_OnSubmitSelection(ByVal Id As Long, ByVal Selection As Object, ByVal SelType As Long, ItemText As String) As Boolean
    PropertyManagerPage2Handler9_OnSubmitSelection = True
End Function

Function PropertyManagerPage2Handler9_OnTabClicked(ByVal Id As Long) As Boolean
    PropertyManagerPage2Handler9_OnTabClicked = True
End Function

Sub PropertyManagerPage2Handler9_OnTextboxChanged(ByVal Id As Long, ByVal Text As String)
End Sub

Sub PropertyManagerPage2Handler9_OnUndo()
End Sub

Sub PropertyManagerPage2Handler9_OnWhatsNew()
End Sub

Sub PropertyManagerPage2Handler9_OnLostFocus(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnGainedFocus(ByVal Id As Long)
End Sub

Sub PropertyManagerPage2Handler9_OnListBoxRMBUp(ByVal Id As Long, ByVal posX As Long, ByVal posY As Long)
End Sub

Function PropertyManagerPage2Handler9_OnWindowFromHandleControlCreated(ByVal Id As Long, ByVal Status As Boolean) As Long
    PropertyManagerPage2Handler9_OnWindowFromHandleControlCreated = 0
End Function

Sub PropertyManagerPage2Handler9_OnNumberboxTrackingCompleted(ByVal Id As Long, ByVal Value As Double)
End Sub

Sub HandleDataChanged()
    
    Dim vCurSketches As Variant
    Dim vCurDepths As Variant
    
    CollectData vCurSketches, vCurDepths
    
    RaiseEvent DataChanged(vCurSketches, vCurDepths)

End Sub

Sub CollectData(ByRef sketches As Variant, ByRef depths As Variant)
    
    Dim swSketches() As SldWorks.Feature
    Dim DepthVals() As Double
    
    Dim i As Integer
        
    For i = 0 To EXTRUDES_COUNT - 1
        
        If False <> swGroupBox(i).Checked Then
            If swSelectionBox(i).ItemCount > 0 Then
                
                Dim selInd As Integer
                selInd = swSelectionBox(i).SelectionIndex(0)
                
                Dim swSketch As SldWorks.Feature
                Set swSketch = swModel.SelectionManager.GetSelectedObject6(selInd, -1)
                
                If (Not swSketches) = -1 Then
                    ReDim swSketches(0)
                    ReDim DepthVals(0)
                Else
                    ReDim Preserve swSketches(UBound(swSketches) + 1)
                    ReDim Preserve DepthVals(UBound(DepthVals) + 1)
                End If
                
                Set swSketches(UBound(swSketches)) = swSketch
                DepthVals(UBound(DepthVals)) = swNumberBox(i).Value
                
            End If
        End If
        
    Next
    
    If (Not swSketches) <> -1 Then
        sketches = swSketches
        depths = DepthVals
    Else
        sketches = Empty
        depths = Empty
    End If
    
End Sub