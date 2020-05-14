---
layout: sw-tool
title: Purge components configurations (remove all unused configurations) from SOLIDWORKS assembly
caption: Purge Components Configurations
description: VBA macro to create copies of all selected components in the assembly and purge configurations in each of them
image: /solidworks-api/document/assembly/components/purge-configurations/purged-components-result1.png
labels: [component, replace, purge]
categories: sw-tools
group: Assembly
---
In some cases it might be required to remove (purge) all unused configurations from the components in the assembly. It is in particular useful for the fastener or toolbox components as file can contains thousands of configurations but only few are used in the assembly.

This macro allows to create a copy of all selected components, purge their configurations and replace them in the assembly.

> It is strongly recommended to backup your assembly before using this macro

You can either select components manually or use [advanced component selection tool](https://help.solidworks.com/2016/English/SolidWorks/sldworks/c_Advanced_Component_Selection_SWassy.htm) to select components based on the criteria (e.g. or fasteners or toolbars):

{% include img.html src="advanced-component-selection.png" alt="Selecting all fasteners and toolbox parts in the assembly via Advanced Components Selection tool" align="center" %}

For additional criteria use the [extended advanced selection macro](/solidworks-api/document/assembly/components/advanced-selection/).

## Notes

* Macro will only work with permanent components. Error will be generated for virtual components
* Macro will only work with part based (*.sldprt) components
* Macro will only work wil fully loaded components, suppressed or lightweight components are not supported
* Macro doesn't save the document after processing. Use *Save All* to save all modifications
* Macro will copy all replacement part at the same location as source part
* Component can be selected in the Feature Manager tree or from the graphics view (it is possible to select any entity of the component as well, such as face or edge)
* Design table will be removed if exists
* Macro will not replace existing files and *File already exist* wil be generated if target file already created. Remove all of these files manually. If macro failed, some of the files may be loaded into the memory despite they are not used in the assembly. Use *Close All* command to release those files
* Mates will be reattached

## Options

### Replacement Name

Specify the name of the replacement file by changing the *REPLACEMENT_NAME* constant. Use fre text with the \[title\] and \[conf\] placeholders which will be replaced with title of the source file and the component's referenced configuration respectively. If the *GROUP_BY_CONFIGURATIONS* option is set to True, the \[conf\] placeholder will be replaced by the join of all configuration names separated by _ symbol.

### Grouping Configurations

*GROUP_BY_CONFIGURATIONS* option allows to specify if the components referencing the same document in different configuration should be replaced by single component or new single configuration part should be created for each component irrespectively.

### Examples

{% include img.html src="components-configurations.png" alt="Unused configurations of components" align="center" %}

There are 2 files with multiple configuration

* Part1.sldprt contains 4 configurations: Default, 2, 3 and 4
* Part2.sldprt contains 6 configurations driven by the design table: Default, A, B, C, D, E
* Part1 is placed into the assembly 2 times in configurations Default and 4
* Part2 is placed into the assembly 2 times in configurations A and B

User selects first 3 components and runs the macro. The following results will be produced depending on the specified settings

### Option 1

~~~ vb
Const GROUP_BY_CONFIGURATIONS As Boolean = False
Const REPLACEMENT_NAME As String = "[title]_[conf]"
~~~

As the result 3 new files will be generated with a single configuration: Part1_Default.sldprt, Part1_4.sldprt, Part2_A.sldprt (design table is removed) and all selected component will be replaced. The 4th component will not be changed as it was not selected initially.

{% include img.html src="purged-components-result1.png" alt="Results of components purge" align="center" %}

### Option 2

~~~ vb
Const GROUP_BY_CONFIGURATIONS As Boolean = True
Const REPLACEMENT_NAME As String = "[title]_[conf]_replacement"
~~~

As the result 2 new files will be generated: Part1_Default_4_replacement.sldprt (with 2 configurations), Part2_A_replacement.sldprt (design table is removed) and all selected component will be replaced. The 4th component will not be changed as it was not selected initially.

{% include img.html src="purged-components-result2.png" alt="Second results of components purge" align="center" %}

{% include_relative Macro.vba.codesnippet %}
