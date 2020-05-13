---
layout: article
title: Troubleshooting the registration of SOLIDWORKS PDM add-ins
caption: Add-Ins
description: Symptoms and resolutions for the most common errors with SOLIDWORKS PDM add-ins development, debugging and registering.
lang: en
image: /images/codestack-snippet.png
labels: [add-in registration]
---
This section contains the list of descriptions of symptoms, explanation of causes and suggestions for resolutions for the most common errors while registering or debugging add-ins for SOLIDWORKS PDM Professional.

{% assign pages = site.pages | where_exp: "p", "p.url contains page.url" | where_exp: "p", "p.url != page.url" %}

{% include catalogue.html pages=pages %}