---
title: Managing permissions in SOLIDWORKS Professional PDM API
caption: Permissions
description: Managing user, groups and folders permissions in SOLIDWORKS PDM Professional using SOLIDWORKS PDM API
labels: [rights, permissions]
order: 5
---
SOLIDWORKS PDM API provides functions to manage various permissions for users, folders and groups.

Folder permissions can be managed via [IEdmUserMgr7::GetFolderPermissions](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmusermgr7~getfolderpermissions.html) and [IEdmUserMgr7::SetFolderPermissions](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmusermgr7~setfolderpermissions.html) SOLIDWORKS PDM API methods.

State permissions can be managed via [IEdmUserMgr9::GetStatePermissions](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmUserMgr9~GetStatePermissions.html) and [IEdmUserMgr9::SetStatePermissions](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmUserMgr9~SetStatePermissions.html) methods.

Transition permissions can be managed via [IEdmUserMgr9::GetTransitionPermissions](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmUserMgr9~GetTransitionPermissions.html) and [IEdmUserMgr9::SetTransitionPermissions](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmUserMgr9~SetTransitionPermissions.html)