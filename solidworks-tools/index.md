---
layout: default
title: Library of macros and scripts to automate SOLIDWORKS
caption: "'Goodies'"
description: Useful automation macros and scripts to increase your productivity when working in SOLIDWORKS
image: macros-library.png
labels: [tools,solidworks,macro]
search: false
styles:
  - /assets/styles/catalogue.css
group-descriptions:
  Model: Automation of 3D Models (parts and assemblies) and 2D drawings
  Materials: Automation of SOLIDWORKS materials database and properties in parts
  Frame: Automation of SOLIDWORKS menus, toolbars, 3rd party add-ins, documents management
  Developers: Helpful utilities for developers building software utilizing SOLIDWORKS API
  Custom Properties: Automation of SOLIDWORKS generic, configuration and cut-list custom properties
  Part: 'Automation of SOLIDWORKS part documents (*.sldprt): geometry, feature tree'
  Assembly: 'Automation of SOLIDWORKS assembly documents (*.sldasm): components, mates'
  Drawing: 'Automation of SOLIDWORKS drawing documents (*.slddrw): tables, views, sheets'
  Security: Enabling additional security and protection for the models and applications using SOLIDWORKS API
  Sketch: Automation of SOLIDWORKS sketch, segments and relations
  Performance: Boosting operational performance of SOLIDWORKS documents and application
  Geometry: 'SOLIDWORKS geometry automation: custom features, topology optimization'
  Import/Export: Automation of importing and exporting SOLIDWORKS files to different formats
  Motion Study: Automation of SOLIDWORKS Motion Study module
  Options: SOLIDWORKS document and system option automation
redirect-from:
  - /p/solidworks-goodies.html
---
# Macro Library for SOLIDWORKS Automation
{% social-share %}

![SOLIDWORKS Macros Library](macros-library.svg){ width=400 }

This page contains a library of useful macros, utilities and scripts for SOLIDWORKS engineers. Macros are grouped by categories: part assembly, drawing, performance etc.

Follow the [Programming VBA and VSTA macros using SOLIDWORKS API](/solidworks-api/getting-started/macros/) section for guidelines of using and creating macros in SOLIDWORKS.

## Best practices for organizing macro library

[MyToolbar](/labs/solidworks/my-toolbar/) is a free and open-source add-in for SOLIDWORKS which allows organize the macro library in custom toolbars integrated to SOLIDWORKS environment. Add-in also allows to manage multi-user environment by storing the configuration in the centralized location.

![Custom macro buttons in the toolbar](macro-library-toolbar.png){ width=450 }

Alternatively macro buttons can be created using native SOLIDWORKS functionality. Read [Creating macro buttons in SOLIDWORKS toolbars](/solidworks-api/getting-started/macros/macro-buttons/) for more information.

Explore this section to find productivity and automation tools which suit your needs.

For additional productivity add-ins visit the [SOLIDWORKS Labs](/labs/solidworks/) page.

---
{% catalogue { type: sw-tool } %}
