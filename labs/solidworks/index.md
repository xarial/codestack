---
layout: default
title: Productivity tools and add-ins for SOLIDWORKS
caption: SOLIDWORKS
menu: LABS
description: Collection of tools and add-ins for SOLIDWORKS
image: /labs/solidworks/solidworks-labs.png
order: 5
---
![SOLIDWORKS Labs](solidworks-labs.svg){ height=150 }

This section lists the various productivity tools and add-ins developed for SOLIDWORKS application.

Most of the products are free and open source, driven by requests coming from the SOLIDWORKS community. Add-ins are fully integrated into SOLIDWORKS providing the native look and feel.

Add-ins are grouped by sections. Each add-in has an online documentation available.

{#% assign pages = site.pages | where_exp: "p", "p.categories contains 'sw-labs'" %}

{#% include catalogue.html pages=pages %}
