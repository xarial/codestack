---
layout: article
title: Catch new feature creation event from SOLIDWORKS API notification
caption: Catch New Feature Creation Event
description: Example listens for feature added event of the active part document and displays the message box
image: /images/codestack-snippet.png
labels: [event, example, feature manager, new feature, solidworks api]
redirect_from:
  - /2018/03/solidworks-api-features-manager-catch-adding-feat-event.html
---
This example listens for feature added event of the active part document using SOLIDWORKS API.

Once the new feature creation notification is caught, macro displays the message box to the user.

The listener is detached as soon as active part is closed.

*Macro Module*

{% include_relative Macro.vba.codesnippet %}

*EventListener Class*

{% include_relative EventListenerClass.vba.codesnippet %}