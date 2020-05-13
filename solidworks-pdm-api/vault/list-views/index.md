---
layout: article
title: List all vault views using SOLIDWORKS PDM API
caption: List All Views
description: Example demonstrates how to list all available vault views and their paths using SOLIDWORKS PDM API
lang: en
image: /solidworks-pdm-api/vault/list-views/pdm-vaults-list.png
labels: [vault, view]
---
{% include img.html src="pdm-vaults-list.png" width=250 alt="Vault views info printed to Console window" align="center" %}

This example demonstrates how to list all available vault views and their paths and prints the information to the console window.

[IEdmVault8::GetVaultViews](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault8~getvaultviews.html) SOLIDWORKS PDM API is used to list the information about all available PDM vaults. Alternatively this information can be retrieved from the Registry.

{% include_relative Console.cs.codesnippet %}
