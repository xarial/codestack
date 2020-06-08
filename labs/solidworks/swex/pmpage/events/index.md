---
title: Handling events of SOLIDWORKS property manager page
caption: Events
description: Overview of property manager page events handling
toc-group-name: labs-solidworks-swex
order: 3
---
[PropertyManagerPageHandlerEx](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_PropertyManagerPageHandlerEx.htm) class is responsible for providing the events raised by property manager page to the client

Instance of the handler will be created by the framework and can be accessed via [PropertyManagerPageEx::Handler](https://docs.codestack.net/swex/pmpage/html/P_CodeStack_SwEx_PMPage_PropertyManagerPageEx_2_Handler.htm) property

~~~ cs
...
m_Page = new PropertyManagerPageEx<MyPMPageHandler, DataModel>(m_Data, m_App);
m_Page.Handler.Closed += r =>
{
    ...
};
...
~~~