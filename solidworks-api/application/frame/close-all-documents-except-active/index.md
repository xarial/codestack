---
layout: sw-tool
title: Macro to close all SOLIDWORKS documents except active
caption: Close All Documents Except Active
description: Closes all opened documents except of an active one using SOLIDWORKS API
lang: en
image: /solidworks-api/application/frame/close-all-documents-except-active/close-all-but-active.png
logo: /solidworks-api/application/frame/close-all-documents-except-active/close-all-but-active.svg
labels: [close, window]
categories: sw-tools
group: Frame
---
{% include img.html src="opened-documents.png" width=250 alt="Documents opened in SOLIDWORKS" align="center" %}

This macro utilizes SOLIDWORKS API and closes all opened documents except of an active one.

If document is dirty (i.e. has any unsaved changes) the macro will prompt user to specify the action (save, do not save or cancel) for the closing documents. Otherwise the document will be closed silently.

{% include_relative Macro.vba.codesnippet %}
