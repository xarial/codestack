---
layout: sw-tool
title: Macro to close all SOLIDWORKS documents except active
caption: Close All Documents Except Active
description: Closes all opened documents except of an active one using SOLIDWORKS API
image: close-all-but-active.svg
labels: [close, window]
group: Frame
---
![Documents opened in SOLIDWORKS](opened-documents.png){ width=250 }

This macro utilizes SOLIDWORKS API and closes all opened documents except of an active one.

If document is dirty (i.e. has any unsaved changes) the macro will prompt user to specify the action (save, do not save or cancel) for the closing documents. Otherwise the document will be closed silently.

Watch [video demonstration](https://youtu.be/9uZCecGg25I?t=166)

{% code-snippet { file-name: Macro.vba } %}
