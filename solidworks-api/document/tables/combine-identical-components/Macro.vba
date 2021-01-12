#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swBomTable As SldWorks.TableAnnotation
    Set swBomTable = swModel.SelectionManager.GetSelectedObject6(1, -1)
    
    CombineIdenticalComponents swModel, swBomTable, 1, swBomTable.RowCount - 1
    
End Sub

Sub CombineIdenticalComponents(model As SldWorks.ModelDoc2, table As SldWorks.BomTableAnnotation, startRowIndex As Integer, entRowIndex As Integer)
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    
    Dim swTableAnnotation As SldWorks.TableAnnotation
    Set swTableAnnotation = table
    
    Dim swAnn As SldWorks.Annotation
    Set swAnn = swTableAnnotation.GetAnnotation()
    
    swSelData.SetCellRange startRowIndex, entRowIndex, 0, 0
    
    swAnn.Select3 False, swSelData
    
    RunCombineIdenticalComponentsCommand
    
End Sub

Sub RunCombineIdenticalComponentsCommand()
    
    Const WM_COMMAND As Long = &H111
        
    Dim swFrame As SldWorks.Frame
        
    Set swFrame = swApp.Frame
        
    Const CMD_COMBINE_IDENTICAL_COMPONENTS As Long = 24378
        
    SendMessage swFrame.GetHWnd(), WM_COMMAND, CMD_COMBINE_IDENTICAL_COMPONENTS, 0
    
End Sub