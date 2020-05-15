---
layout: article
title: Configuring macro buttons in MyToolbar add-in for SOLIDWORKS
caption: Configuration
description: Adding, removing and customizing (tooltips and icons) toolbars and macro buttons in MyToolbar add-in for SOLIDWORKS
labels: [configure,macro button icon,macro button tooltip]
toc-group-name: labs-solidworks-my-toolbar
order: 1
---
## Adding Macro Buttons And Toolbars

Macro buttons can be configured by clicking on 'Configure...' command in MyToolbar menu in SOLIDWORKS

![MyToolbar add-in menu in SOLIDWORKS](my-toolbar-menu.png){ width=350 }

1. MyToolbar sub-menu in SOLIDWORKS Tools menu
1. Button to configure toolbar
1. About MyToolbar application

Configuration dialog is displayed and existing toolbar and buttons can be modified as well as new can be added.

Click on green plus (+) button to add new toolbar or macro button.

Configure the parameters as shown below:

### Configuring Macro Buttons In Toolbar

![Editing macro button in MyToolbar](edit-macro.png){ width=650 }

1. Path to toolbar file. This setting is stored locally to the user. UNC path is supported if required to configure a [shared toolbar](/labs/solidworks/my-toolbar/user-guide/multi-user/)
1. Add new toolbar
1. Add new macro button in toolbar
1. Title of the macro button. It displayed as a bold header in the tooltip when mouse hovers a button in SOLIDWORKS toolbar
1. Description of the macro button. It displayed as a sub-header in the tooltip and in the SOLIDWORKS task bar in the bottom right corner when mouse hovers a button in SOLIDWORKS toolbar
1. Icon of the macro button. Optimal size is between 16x16 and 120x120 in PNG format, however image will be automatically scaled and aligned which allows support for any size (including different width and height). Transparency is supported.
1. Full path to macro to run
1. Macro entry point. This is the subroutine which should be run first when executing the macro. This is a parameterless subroutine (usually named main). MyToolbar will try to automatically find the best suitable subroutine

### Configuring Toolbar

![Editing toolbar in MyToolbar](edit-toolbar.png){ width=650 }

1. Toolbar title to be displayed in the SOLIDWORKS toolbars manager
1. Toolbar tooltip
1. Toolbar icon
1. Preview of toolbar icon (default icon)

### Modifying Commands

Select macro buttons and toolbars to load and edit the parameters. Use context menu or commands box to reorganize the commands as shown below.

![Caption](modifying-commands.png){ width=350 }

1. Move selected macro button to the left or move toolbar up
1. Move selected macro button to the right or move toolbar down
1. Add new macro button left to the selected macro button or add new toolbar above the selected toolbar
1. Add new macro button right to the selected macro button or add new toolbar below the selected toolbar
1. Remove selected macro button or toolbar
1. Save changes and close dialog
1. Close dialog without saving changes

## Saving Changes

If toolbar configuration changed the following message is displayed and SOLIDWORKS needs to be restarted for settings to take effect (unless toolbar is read-only in multi-user environment).

![MyToolbar specification changed notification](toolbar-spec-changed.png){ width=350 }

## Accessing Macro Buttons

Toolbars are available in the SOLIDWORKS toolbars list:

![SOLIDWORKS toolbars list](solidworks-toolbars.png){ width=250 }

and in the SOLIDWORKS menu.

![Macro buttons in SOLIDWORKS menu](my-toolbar-commands-menu.png){ width=450 }

Toolbars can be reorganized and placed into the SOLIDWORKS command manager area. Visit [Customization](/solidworks/my-toolbar/user-guide/customization/) page for additional customization options for the toolbars and commands.

![Macro buttons in SOLIDWORKS toolbar](my-toolbar-commands.png){ width=350 }