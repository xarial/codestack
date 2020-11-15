---
layout: sw-tool
title: Open selected components in the Large Design Review (LDR) mode
caption: Open components in LDR mode
description: VBA macro to open all selected sub-assemblies and part components in the Large Design Review (LDR) mode and view only mode correspondingly
image: ldr-sub-assembly.svg
labels: [performance, ldr, view only, large design review, open]
group: Assembly
---
This VBA macro can be run from the assembly opened in the Large Design Review (LDR) mode or drawing opened in the Detailing mode. Macro will open all selected components in their own windows, but unlike out-of-the-box functionality assemblies will not be resolved and will preserve the LDR mode.

![Selected sub-assembly components](selected-sub-assemblies.png)

It is then possible to enable the editing in the LDR mode, modify the assembly and update the graphics in the top level assembly.

Preserving the LDR mode on all the steps will significantly improve the performance.

## Drawings

This macro can also work from the drawing opened in the Detailing mode. It is required to select drawing view(s) before running the macro.

![Selected drawing view](selected-drawing-view.png)

To enable drawing support it is required to enable Document Manager API in the macro. Follow [Activating Document Manager](/solidworks-document-manager-api/getting-started/create-connection#activating-document-manager) section for the detailed steps to request Document Manager license key.

Add the reference to **SwDocumentMgr [Year] Type Library** under the **Tools->References** menu in VBA editor. Follow [Document Manager in VBA](/solidworks-document-manager-api/getting-started/create-connection#vba) for more information.

![Document Manager reference](swdm-reference.png)

Set the license key in the **DM_LIC_KEY** variable. Note that this macro only requires **swdocmgr_general** portion of the key. The following format would be sufficient.

~~~ vb
Const DM_LIC_KEY As String = "[CompanyName]:swdocmgr_general-00000-{31 times}"
~~~

This routine is not required if macro will only be used from assemblies.

## Notes and limitations

* Sub-assembly components will be opened in Large Design Review mode while part components will be opened in the View Only mode
* If target parts or assemblies do not have display data stored - error will be thrown
* Components must be selected from the Feature Manager Tree. Entities selected in the graphics area will be ignored
* This VBA macro is using the simplified version of the [Search Routine for Referenced Documents](https://help.solidworks.com/2016/english/SolidWorks/sldworks/c_Search_Routine_for_Referenced_Documents.htm) and only checks active assembly's folder and sub folders before falling back on the cached path of the component. In some cases this may result in incorrect reference loaded (e.g if search folders are used). But this will only apply to the assembly which was copied and cached file paths were never updated.

### Referenced Configurations

This macro will attempt to open the assembly in the referenced configuration of the component, however by default SOLIDWORKS only stores the display data in the active configuration, unless configurations are marked with 'Display Data Mark' flag

![Add display data mark flag to configuration](add-display-data-mark.png){ width=250 }

If the referenced configuration of the component is not marked with the above flag and it is not an active configuration, then it cannot be loaded in Large Design Review. In this case macro will load the default configuration and display the below warning indicating that the graphics of different configuration is loaded.

![Error with invalid configuration](configuration-error.png)

{% code-snippet { file-name: Macro.vba } %}