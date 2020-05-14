---
layout: article
title: Localizing SOLIDWORKS add-ins using SwEx framework
caption: Localization
description: How to support multi language SOLIDWORKS add-ins by using of localized resources in SwEx framework
image: /labs/solidworks/swex/localization/menu-localized.png
toc_group_name: labs-solidworks-swex
order: 6
---
SwEx frameworks supports [resources in .NET applications](https://docs.microsoft.com/en-us/dotnet/framework/resources/index) to enable localization of the add-in, e.g. supporting multiple languages.

This technique allows to load localized string at runtime based on the Windows settings in the control panel.

{% include img.html src="region-format.png" height=450 alt="Region and language page in Control Panel" align="center" %}

Resources should be added to the corresponding localized .resx file (e.g. Resources.resx for default, Resources.ru.resx for Russian, Resources.fr.resx for French, etc.)

{% include img.html src="resource-files.png" alt="Resource files in the solutions" align="center" %}

In order to reference the string from the resource, use the overloads of the constructors for the [TitleAttribute](https://docs.codestack.net/swex/common/html/M_CodeStack_SwEx_Common_Attributes_TitleAttribute__ctor_1.htm) and [SummaryAttribute](https://docs.codestack.net/swex/common/html/M_CodeStack_SwEx_Common_Attributes_SummaryAttribute__ctor_1.htm) which allows to define title, tooltips, hint string for all elements across SwEx framework (i.e. [menu commands](#menu), [property page controls](#property-manager-page), [macro feature](#macro-feature), etc.)

Below is an example which demonstrates this technique. Text is localized as per resources below:

{% include img.html src="visual-studio-resources.png" alt="Localized resource files in the Visual Studio" align="center" %}

## Menu

Two commands in menu are localized for Russian and English versions of the add-in.

{% include img.html src="menu-localized.png" alt="Localized menu commands" align="center" %}

{% include code-tabs.html src="LocalizationAddIn.Commands" %}

## Property Manager Page

Property Manager page title and tooltips for the controls are localized for Russian and English versions of the add-in.

{% include img.html src="property-page-localized.png" alt="Localized Property Manager Page" align="center" %}

{% include code-tabs.html src="LocalizationAddIn.PMPage" %}

## Macro Feature

Macro feature base name is localized to Russian and English versions of the add-in.

> Note. Base name is only assigned while feature creation, feature won't be renamed after locale has changed.

{% include img.html src="macro-feature-localized.png" alt="Localized Macro Feature base name" align="center" %}

In a similar way it is possible to use strings from the resources to return another data, e.g. text of the error for macro feature.

{% include img.html src="macro-feature-error-localized.png" alt="Localized macro feature error" align="center" %}

{% include code-tabs.html src="LocalizationAddIn.MacroFeature" %}
