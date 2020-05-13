---
layout: master
title: Library of macros and scripts to automate SOLIDWORKS
caption: "'Goodies'"
description: Useful automation macros and scripts to increase your productivity when working in SOLIDWORKS
menu: SOLIDWORKS
lang: en
image: /solidworks-tools/macros-library.png
labels: [tools,solidworks,macro]
order: 5
group_descriptions:
  - name: "Model"
    description: "Automation of 3D Models (parts and assemblies) and 2D drawings"
  - name: "Materials"
    description: "Automation of SOLIDWORKS materials database and properties in parts"
  - name: "Frame"
    description: "Automation of SOLIDWORKS menus, toolbars, 3rd party add-ins, documents management"
  - name: "Developers"
    description: "Helpful utilities for developers building software utilizing SOLIDWORKS API"
  - name: "Custom Properties"
    description: "Automation of SOLIDWORKS generic, configuration and cut-list custom properties"
  - name: "Part"
    description: "Automation of SOLIDWORKS part documents (*.sldprt): geometry, feature tree"
  - name: "Assembly"
    description: "Automation of SOLIDWORKS assembly documents (*.sldasm): components, mates"
  - name: "Drawing"
    description: "Automation of SOLIDWORKS drawing documents (*.slddrw): tables, views, sheets"
  - name: "Security"
    description: "Enabling additional security and protection for the models and applications using SOLIDWORKS API"
  - name: "Sketch"
    description: "Automation of SOLIDWORKS sketch, segments and relations"
  - name: "Performance"
    description: "Boosting operational performance of SOLIDWORKS documents and application"
  - name: "Geometry"
    description: "SOLIDWORKS geometry automation: custom features, topology optimization"
  - name: "Import/Export"
    description: "Automation of importing and exporting SOLIDWORKS files to different formats"
  - name: "Motion Study"
    description: "Automation of SOLIDWORKS Motion Study module"
  - name: "Options"
    description: "SOLIDWORKS document and system option automation"
redirect_from:
  - /p/solidworks-goodies.html
---
# Macro Library for SOLIDWORKS Automation

{% include img.html src="macros-library.svg" width=400 nofigcaption=true alt="SOLIDWORKS Macros Library" align="center" %}

{%- include social-share.html -%}

This page contains a library of useful macros, utilities and scripts for SOLIDWORKS engineers. Macros are grouped by categories: part assembly, drawing, performance etc.

Follow the [Programming VBA and VSTA macros using SOLIDWORKS API]({{ "/solidworks-api/getting-started/macros/" | relative_url }}) section for guidelines of using and creating macros in SOLIDWORKS.

## Best practices for organizing macro library

[MyToolbar]({{ "/labs/solidworks/my-toolbar/" | relative_url }}) is a free and open-source add-in for SOLIDWORKS which allows organize the macro library in custom toolbars integrated to SOLIDWORKS environment. Add-in also allows to manage multi-user environment by storing the configuration in the centralized location.

{% include img.html src="macro-library-toolbar.png" width=450 alt="Custom macro buttons in the toolbar" align="center" %}

Alternatively macro buttons can be created using native SOLIDWORKS functionality. Read [Creating macro buttons in SOLIDWORKS toolbars]({{ "/solidworks-api/getting-started/macros/macro-buttons/" | relative_url }}) for more information.

Explore this section to find productivity and automation tools which suit your needs.

For additional productivity add-ins visit the [SOLIDWORKS Labs]({{ "/labs/solidworks/" | relative_url }}) page.

{% assign pages = site.pages | where_exp: "p", "p.categories contains 'sw-tools'" %}

---
{% include catalogue.html pages=pages group_descriptions=page.group_descriptions %}
