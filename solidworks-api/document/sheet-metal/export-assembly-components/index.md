---
layout: sw-tool
title: Export flat patterns from SOLIDWORKS assembly components
caption: Export Flat Patterns From Assembly Components
description: VBA macro to export flat patterns from all components of the active assembly
image: assembly-flat-pattern.svg
labels: [where used,parent,component]
group: Assembly
---
This VBA macro allows to export flat patterns to DXF/DWG from all sheet metal components in the active SOLIDWORKS assembly.

Macro enables flexibility in specifying the name of the output file allowing to use placeholders (original file name, feature name, cut-list custom property) combined with the free text and supports specifying sub-folders.

{%youtube id: FtXkdSlekG8 %}

## Configuration

Macro can be configured by modifying the **OUT_NAME_TEMPLATE** and **FLAT_PATTERN_OPTIONS** constants

### Output name template

This constant allows to specify template for the output path of the flat pattern.

This can be either absolute or relative path. If later, result will be saved relative to the assembly directory.

Extension (either .dxf or .dwg) must be specified as the part of naming template

The following placeholders are supported

* <\_FileName\_> - name of the part file (without extension) where the flat pattern resides in
* <\_FeatureName\_> - name of the flat pattern feature
* <\_ConfName\_> - name of the configuration of this flat pattern (i.e. referenced configuration of the component)
* <[PropertyName]> - any name of the cut-list property to read value from, e.g. \<Thickness\> is replaced with the value of cut-list custom property *Thickness*

Placeholders will be resolved for each flat pattern at runtime.

For example the following value will save flat patterns with the name of the part document in the *DXFs* sub-folder in the same folder as main assembly

~~~ vb
Const OUT_NAME_TEMPLATE As String = "DXFs\<_FileName_>.dxf"
~~~

While the following name will save all of the flat patterns as DWG file into the *Output* folder in *D* drive, where the file name will be extracted from the *PartNo* property for each corresponding flat pattern.

~~~ vb
Const OUT_NAME_TEMPLATE As String = "D:\Output\<PartNo>.dwg"
~~~

### Flat pattern options

Options can be configured by specifying the values of **FLAT_PATTERN_OPTIONS**. Use **+** to combine options

![Flat pattern export options](flat-pattern-export-options.png)

For example to export hidden edges, library features and forming tools, use the setting below.

~~~ vb
Const FLAT_PATTERN_OPTIONS As Integer = SheetMetalOptions_e.IncludeHiddenEdges + SheetMetalOptions_e.ExportLibraryFeatures + SheetMetalOptions_e.ExportFormingTools
~~~

> Note, geometry option must always be specified as it is required for the flat pattern export

## Notes

* Macro will ask to resolve lightweight components if any
* Each flat pattern from the multi-body sheet metal part will be exported. Make sure to use either <\_FeatureName\_> or <[PropertyName]> to differentiate between result files
* Macro will process all distinct components (file path + configuration)
* Macro will automatically create folders if required

{% code-snippet { file-name: Macro.vba } %}
