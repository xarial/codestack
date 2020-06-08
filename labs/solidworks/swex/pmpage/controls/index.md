---
title: Adding and customizing controls to Property Manager page 
caption: Controls
description: Overview of controls supported by the SwEx.PMPage framework and the customization and decoration options
toc-group-name: labs-solidworks-swex
order: 2
---
Framework will automatically generate the best suitable control for the public property in the data model. For example for all numeric properties the number box control will be generated. For all string properties text box control will be generated. For all complex types group box will be generated.

The style of the controls can be customized via attributes.

## Accessing controls

Access to controls is provided via [IPropertyManagerPageControlEx](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Controls_IPropertyManagerPageControlEx.htm) wrapper interface. Common properties can be accessed via this interface (such as control id, enable or visible flags). Underlying native SOLIDWORKS control can be accessed via [IPropertyManagerPageControlEx::SwControl](https://docs.codestack.net/swex/pmpage/html/P_CodeStack_SwEx_PMPage_Controls_IPropertyManagerPageControlEx_SwControl.htm) property. It returns the pointer to corresponding [IPropertyManagerPageControl](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipropertymanagerpagecontrol.html) which can be cast to specific control such as [IPropertyManagerPageSelectionbox](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipropertymanagerpageselectionbox.html), [IPropertyManagerPageCombobox](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipropertymanagerpagecombobox.html), [IPropertyManagerPageTextbox](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipropertymanagerpagetextbox.html) etc.

All controls can be accessed via [IPropertyManagerPageEx::Controls](https://docs.codestack.net/swex/pmpage/html/P_CodeStack_SwEx_PMPage_Base_IPropertyManagerPageEx_2_Controls.htm) property.