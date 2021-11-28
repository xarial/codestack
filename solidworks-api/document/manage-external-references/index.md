---
layout: sw-tool
title: SOLIDWORKS macro to break, lock or unlock all external references for files and components
caption: Manage All External References For Components
description: Example will break, lock or unlock all external references for the file or components in the active assembly using SOLIDWORKS API
image: break-all-external-references.png
labels: [assembly, external references, solidworks api, utility]
group: Assembly
redirect-from:
  - /2018/03/solidworks-api-assembly-break-all-external-references-for-components.html
  - /solidworks-api/document/assembly/break-components-external-references
  - /solidworks-api/document/assembly/components/break-external-references/
---
This example will break, lock or unlock all external references for the active model or all or selected components in the active assembly using SOLIDWORKS API.

![Command to break all external references](break-all-external-references.png){ width=640 }

## Configuration

Macro can be configured by modifying the value of the constants

~~~ vb
Const MODIFY_ACTION As Integer = ModifyAction_e.UnlockAll 'Action to call on the references in the model. Supported values: BreakAll, LockAll, UnlockAll
Const REFS_SCOPE As Integer = Scope_e.AllComponents 'Scope to run the above action. Supported values: ThisFile, TopLevelComponents, AllComponents, SelectedComponents
~~~

## CAD+

This macro is compatible with [Toolbar+](https://cadplus.xarial.com/toolbar/) and [Batch+](https://cadplus.xarial.com/batch/) tools so the buttons can be added to toolbar and assigned with shortcut for easier access or run in the batch mode.

In order to enable [macro arguments](https://cadplus.xarial.com/toolbar/configuration/arguments/) set the **ARGS** constant to true

~~~ vb
#Const ARGS = True
~~~

In this case it is not required to make copies of the macro to set individual [options for action and scope](#configuration).

Instead specify 2 arguments:

1. Use the **-b**, **-l**, **-u**, to set the action to **Break All**, **Lock All**, **Unlock All** respectively
1. Use the **-f**, **-t**, **-a** to set the scope to **This File**, **Top Level Components**, **All Components** respectively

For example the parameters below will lock all external references of the file itself

~~~
> -l -f
~~~

While the following command will break all external references for all components of the assembly (including sub-components)

~~~
> -b -a
~~~

{% code-snippet { file-name: Macro.vba } %}
