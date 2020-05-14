---
layout: article
title: 'SOLIDWORKS Add-Ins Troubleshooting: Issues And Resolutions'
caption: 'Add-Ins Troubleshooting: Issues And Resolutions'
description: Overview and solutions for the most common errors of SOLIDWORKS add-ins and SDK
image: /images/codestack-snippet.png
labels: [add-in, sdk, not working, problem, solidworks api, troubleshooting]
---
This section covers the most common symptoms of the errors in the SOLIDWORKS add-ins and Software Development Kit (SDK).

Follow the corresponding link to get the detailed description of the issues, its cause and the steps to resolve the problem.

{#% assign pages = site.pages | where_exp: "p", "p.url contains page.url" | where_exp: "p", "p.url != page.url" %}

{#% include catalogue.html pages=pages %}
