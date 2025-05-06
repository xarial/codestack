---
caption: Select By Custom Property
title: VBA macro to select components based on the custom property
description: VBA macro selects components based on the value of the custom property and type in SOLIDWORKS assembly
---

This VBA macro batch selects components from the active SOLIDWORKS assembly based on the specified filters.

## Configuration

Specify the type of the file in the **FILE_TYPE** constant (either *swDocumentTypes_e.swDocPART* or *swDocumentTypes_e.swDocASSEMBLY*)

Define filters in the **PROPERTY_FILTERS** array:

* Set **PropertyName** property to the name of the target property
* Set **Patterns** property to an array of Regular Expressions to match the proeprty value
* Set **Inclusive** to indicate if result needs to be included or excluded if the property value matches the specified pattern

~~~ vb
Const FILE_TYPE As Integer = swDocumentTypes_e.swDocPART

Dim PROPERTY_FILTERS() As PropertyFilter

Dim swApp As SldWorks.SldWorks

Sub main()

    ReDim PROPERTY_FILTERS(1)
    
    PROPERTY_FILTERS(0).PropertyName = "Type" 'Check the custom property type,
    PROPERTY_FILTERS(0).Patterns = Array(".*MadeToStock.*", ".*PurchasedToStock.*") 'If value of property contains MadeToStock or PurchasedToStock
    PROPERTY_FILTERS(0).Inclusive = True 'Include the result
    
    PROPERTY_FILTERS(1).PropertyName = "StockNumber" 'Also validate the value of the custom property StockNumber ...
    PROPERTY_FILTERS(1).Patterns = Array(".+") 'If it has any value
    PROPERTY_FILTERS(1).Inclusive = False 'Exclude the result
~~~

{% code-snippet { file-name: Macro.vba } %}