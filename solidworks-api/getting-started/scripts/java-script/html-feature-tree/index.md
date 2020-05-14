---
layout: article
title: Using SOLIDWORKS API to render feature tree in HTML page
caption: Render Feature Tree In HTML Page
description: Example demonstrates how to extract and render feature tree of SOLIDWORKS part document in HTML page using SOLIDWORKS API with JavaScript and ActiveX control in Internet Explorer
image: /solidworks-api/getting-started/scripts/java-script/html-feature-tree/html-feature-tree-rendered.png
labels: [JavaScript, feature manager]
---
This example demonstrates how to load feature tree content of the SOLIDWORKS part file using SOLIDWORKS API into the HTML page using JavaScript and ActiveX in Internet Explorer (this will not work in any other browsers as ActiveX is not supported by default - it might be required to install special plugins to enable the support).

* Create new html file
* Copy paste the following code into the file
{% include_relative page.html.codesnippet %}

* Save the file and open in in MS Internet Explorer
{% include img.html src="input-html-page.png" alt="HTML page with input fields" align="center" %}

This page is using ActiveX so the following message can be displayed:

{% include img.html src="ie-activex-run-restriction.png" alt="ActiveX restrictions warning in Internet Explorer" align="center" %}

Click *Allow blocked content* button

* Enter the full path to the SOLIDWORKS part into the text box input field

* Click the *Get Feature Tree* button

* Click *Yes* on the following popup

{% include img.html src="ie-allow-activex.png" width=350 alt="Warning message regarding the ActiveX content" align="center" %}

As the result the feature tree of the part is rendered on the page

{% include img.html src="html-feature-tree-rendered.png" width=250 alt="SOLIDWORKS part feature tree rendered in HTML" align="center" %}