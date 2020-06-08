---
title: Hooking the notifications in SOLIDWORKS PDM API
caption: Hooks
description: Articles and examples explaining how to use event hooks in SOLIDWORKS PDM add-in from API
labels: [hooks, add-in]
order: 6
---
SOLIDWORKS PDM raises multiple events during the operation. Those events include but not limited to

* Check In and Check Out
* Workflow state change
* Data card values change
* File operations: creation, addition, deletion

SOLIDWORKS PDM API provides an access to these hooks via [IEdmAddIn5::OnCmd](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmaddin5~oncmd.html) overload.

This section explains the use of hooks and provides various code examples of add-ins which are using hooks with SOLIDWORKS PDM API.