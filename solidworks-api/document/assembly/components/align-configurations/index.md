---
caption: Align Referenced Configurations
title: VBA macro to align referenced configuration of components to assembly configurations
description: VBA macro aligns referenced configuration of selected components in the SOLIDWORKS assembly to the corresponding assembly configurations
image: modify-configurations.png
---

This VBA macro aligns the referenced configurations of all selected components to the corresponding assembly configuration. For example if assembly has 3 configurations **A**, **B** and **C**, then referenced configurations for all selected components will be set to **A**, **B** and **C** in the respective configuration of the assembly.

![Modify component configurations](modify-configurations.png){ width=600 }

Macro processes all root configurations (or optionally all configurations)

~~~ vb
Const ROOT_CONFS_ONLY As Boolean = False 'Process all assembly configurations
~~~

Multiple components can be selected and processed at the same time. Only top level-components are supported. For aligning configurations for sub-assembly, it is required to activate the sub-assembly in its own window.

Components in the lightweight mode are supported.

{% code-snippet { file-name: Macro.vba } %}