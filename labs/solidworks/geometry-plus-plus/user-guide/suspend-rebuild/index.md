---
layout: article
title: Suspend SOLIDWORKS rebuild operation using Geometry++
caption: Suspend Rebuild
description: Suspend SOLIDWORKS rebuild operations in part, assembly and drawing to rebuild in batch to improve performance using Geometry++ add-in
image: /labs/solidworks/geometry-plus-plus/user-guide/suspend-rebuild/icon.png
toc_group_name: labs-solidworks-geometry-plus-plus
---
{% include youtube.html id="QW3tYaNAfo0" width=560 height=315 %}

This command allows to temporary suspend rebuild operation while still allowing to modify the dimensions, sketches and feature definitions.

This approach allows to greatly reduce the modelling time by execution rebuild operations in a batch mode.

Command is available in menu, toolbar and command manager tab and acts as a toggle button.

When button is not toggled the suspend rebuild mode is disabled and rebuild operations are performed normally.

![Suspend Rebuild commands in toolbar and command manager](not-suspended-buttons-state.png)

When the button is toggled all rebuild operations are suspended.

![Suspend rebuild enabled](suspended-buttons-state.png)

The status bar displays the information about the number of currently suspended rebuild operations.

![Number of suspended rebuilds in the status bar](status-bar-message.png)

In suspend rebuild mode the changes are not resolved and model will remain unchanged. When editing the definitions of the features and closing the Property Manager Page all the features below the edited feature become disabled (not editable) until model is rebuild.

Once batch editing completed disable the *Suspend Rebuild* button and click *Rebuild (ctrl+B)* or *Regenerate (ctrl+Q)* command to update the model.

> Disclaimer: Although the functionality of suspend rebuild is implemented using the SOLIDWORKS API, suppressed rebuild may cause unexpected behaviour of the model. However there were no reported issues of any damage or corruption. Use on your own risk.
