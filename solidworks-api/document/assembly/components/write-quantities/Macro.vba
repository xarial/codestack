'**********************
'Copyright(C) 2020 Xarial Pty Limited
'Reference: http://localhost:82/solidworks-api/document/assembly/components/write-quantities/
'License: https://www.codestack.net/license/
'**********************

Type BomPosition
    model As SldWorks.ModelDoc2
    Configuration As String
    Quantity As Double
End Type

Const PRP_NAME As String = "Qty"
Const MERGE_CONFIGURATIONS As Boolean = False
Const INCLUDE_BOM_EXCLUDED As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
try_:
    On Error GoTo catch_
    
    Dim swAssy As SldWorks.AssemblyDoc
    
    Set swAssy = swApp.ActiveDoc
    
    If swAssy Is Nothing Then
        Err.Raise vbError, "", "Assembly is not opened"
    End If
    
    swAssy.ResolveAllLightWeightComponents True

    Dim swConf As SldWorks.Configuration
    Set swConf = swAssy.ConfigurationManager.ActiveConfiguration

    Dim bom() As BomPosition
    ComposeFlatBom swConf.GetRootComponent3(True), bom
        
    If (Not bom) <> -1 Then
        WriteBomQuantities bom
    End If
    
    GoTo finally_
catch_:
    MsgBox Err.Description, vbCritical, "Count Components"
finally_:
    
End Sub

Sub ComposeFlatBom(swParentComp As SldWorks.Component2, bom() As BomPosition)
        
    Dim vComps As Variant
    vComps = swParentComp.GetChildren
    
    If Not IsEmpty(vComps) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vComps)
            
            Dim swComp As SldWorks.Component2
            Set swComp = vComps(i)
            
            If swComp.GetSuppression() <> swComponentSuppressionState_e.swComponentSuppressed And (False = swComp.ExcludeFromBOM Or INCLUDE_BOM_EXCLUDED) Then
                
                Dim swRefModel As SldWorks.ModelDoc2
                Set swRefModel = swComp.GetModelDoc2()
                
                If swRefModel Is Nothing Then
                    Err.Raise vbError, "", swComp.GetPathName() & " model is not loaded"
                End If
                
                Dim swRefConf As SldWorks.Configuration
                Set swRefConf = swRefModel.GetConfigurationByName(swComp.ReferencedConfiguration)
                
                Dim bomChildType As Integer
                bomChildType = swRefConf.ChildComponentDisplayInBOM
                
                If bomChildType <> swChildComponentInBOMOption_e.swChildComponent_Promote Then
                
                    Dim bomPos As Integer
                    bomPos = FindBomPosition(bom, swComp)
                    
                    If bomPos = -1 Then
                        
                        If (Not bom) = -1 Then
                            ReDim bom(0)
                        Else
                            ReDim Preserve bom(UBound(bom) + 1)
                        End If
                                            
                        bomPos = UBound(bom)
        
                        Dim refConfName As String
            
                        If MERGE_CONFIGURATIONS Then
                            refConfName = ""
                        Else
                            refConfName = swComp.ReferencedConfiguration
                        End If
        
                        Set bom(bomPos).model = swRefModel
                        bom(bomPos).Configuration = refConfName
                        bom(bomPos).Quantity = GetQuantity(swComp)
                                            
                    Else
                        bom(bomPos).Quantity = bom(bomPos).Quantity + GetQuantity(swComp)
                    End If
                
                End If
                
                If bomChildType <> swChildComponentInBOMOption_e.swChildComponent_Hide Then
                    ComposeFlatBom swComp, bom
                End If
                
            End If
            
        Next
    
    End If
    
End Sub

Function FindBomPosition(bom() As BomPosition, comp As SldWorks.Component2) As Integer
        
    FindBomPosition = -1
    
    Dim i As Integer
    
    If (Not bom) <> -1 Then
        
        Dim refConfName As String
        
        If MERGE_CONFIGURATIONS Then
            refConfName = ""
        Else
            refConfName = comp.ReferencedConfiguration
        End If
        
        For i = 0 To UBound(bom)
            If LCase(bom(i).model.GetPathName()) = LCase(comp.GetPathName()) And LCase(bom(i).Configuration) = LCase(refConfName) Then
                FindBomPosition = i
                Exit Function
            End If
        Next
    End If
    
End Function

Function GetQuantity(comp As SldWorks.Component2) As Double

On Error GoTo err_

    Dim refModel As SldWorks.ModelDoc2
    Set refModel = comp.GetModelDoc2
    
    Dim qtyPrpName As String
    
    qtyPrpName = GetPropertyValue(refModel, comp.ReferencedConfiguration, "UNIT_OF_MEASURE")
    
    If qtyPrpName <> "" Then
        GetQuantity = CDbl(GetPropertyValue(refModel, comp.ReferencedConfiguration, qtyPrpName))
    Else
        GetQuantity = 1
    End If
    
    Exit Function

err_:
    Debug.Print "Failed to extract quantity of " & comp.Name2 & ": " & Err.Description
    GetQuantity = 1

End Function

Function GetPropertyValue(model As SldWorks.ModelDoc2, conf As String, prpName As String) As String
    
    Dim confSpecPrpMgr As SldWorks.CustomPropertyManager
    Dim genPrpMgr As SldWorks.CustomPropertyManager
    
    Set confSpecPrpMgr = model.Extension.CustomPropertyManager(conf)
    Set genPrpMgr = model.Extension.CustomPropertyManager("")
    
    Dim prpResVal As String
    
    confSpecPrpMgr.Get3 prpName, False, "", prpResVal
    
    If prpResVal = "" Then
        genPrpMgr.Get3 prpName, False, "", prpResVal
    End If
    
    GetPropertyValue = prpResVal
    
End Function

Sub WriteBomQuantities(bom() As BomPosition)
    
    Dim i As Integer
    
    If (Not bom) <> -1 Then
        
        For i = 0 To UBound(bom)
            
            Dim refConfName As String
        
            If MERGE_CONFIGURATIONS Then
                refConfName = ""
            Else
                refConfName = bom(i).Configuration
            End If
            
            Dim swRefModel As SldWorks.ModelDoc2
            Set swRefModel = bom(i).model
            
            Dim swCustPrpsMgr As SldWorks.CustomPropertyManager
            Set swCustPrpsMgr = swRefModel.Extension.CustomPropertyManager(refConfName)
            
            swCustPrpsMgr.Add3 PRP_NAME, swCustomInfoType_e.swCustomInfoText, bom(i).Quantity, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
            swCustPrpsMgr.Set2 PRP_NAME, bom(i).Quantity
            
        Next
    End If
    
End Sub