Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Const SET_CONFIGURATIONS As Boolean = False

Const BOM_QTY_PRP_NAME As String = "UNIT_OF_MEASURE"
Const QTY_PRP_NAME As String = "Qty"

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim swCustPrpMgr As SldWorks.CustomPropertyManager
        
        Set swCustPrpMgr = swModel.Extension.CustomPropertyManager("")
            
        SetQtyCustomProperty swCustPrpMgr
        
        If SET_CONFIGURATIONS Then
            
            Dim vConfNames As Variant
            vConfNames = swModel.GetConfigurationNames
            
            If Not IsEmpty(vConfNames) Then
                
                Dim i As Integer
                
                For i = 0 To UBound(vConfNames)
                    Dim swConfCustPrpMgr As SldWorks.CustomPropertyManager
                    Set swConfCustPrpMgr = swModel.Extension.CustomPropertyManager(CStr(vConfNames(i)))
                    SetQtyCustomProperty swConfCustPrpMgr
                Next
                
            End If
            
        End If
    
    Else
        
        MsgBox "Please open model"
        
    End If
    
End Sub

Sub SetQtyCustomProperty(custPrpMgr As SldWorks.CustomPropertyManager)
    
    Dim bomQtyPrp As String
    custPrpMgr.Get3 BOM_QTY_PRP_NAME, False, "", bomQtyPrp
        
    Debug.Print bomQtyPrp
        
    custPrpMgr.Add2 BOM_QTY_PRP_NAME, swCustomInfoType_e.swCustomInfoText, QTY_PRP_NAME
    custPrpMgr.Set2 BOM_QTY_PRP_NAME, QTY_PRP_NAME
    
End Sub