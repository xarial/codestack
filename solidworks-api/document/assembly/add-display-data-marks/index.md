---
layout: sw-tool
caption: Add Display Data Marks
title: Macro to add display data marks to configuration used by the main SOLIDWORKS assembly
description: VBA macro to add display data mark for configurations used in Large assembly to be opened in Large Design Review Mode
image: display-data-mark.svg
group: Assembly
---
This VBA macro is useful for the users working with assemblies in the Large Design Review mode.

By default only active configuration is preserved for using the the Large Design Review mode and other configurations of the assembly cannot be activated:

![No display marks in the assembly configurations](configuration-no-display-marks.png)

This macro will traverse all components of the root assembly and find all the used configurations and add the display mark data to all of them.

![Add display data mark command](add-display-data-mark.png)

This will allow to open all sub components in the Large Design Review mode and activate used configurations.

{% code-snippet { file-name: Macro.vba } %}