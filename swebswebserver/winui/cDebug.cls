VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CSEH: WinUI Custom
'***************************************************************************
'
' SWEBS/WinUI
'
' Copyright (c) 2003 Adam Caudill.
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'***************************************************************************

Option Explicit

'local variable(s) to hold property value(s)
Private mvarEnabled As Boolean 'local copy
Private mvarDebugLang As Boolean 'local copy
Private mvarDisableSplash As Boolean 'local copy
Private mvarDisableFade As Boolean 'local copy
Private mvarDisableUpdate As Boolean 'local copy
Private mvarDisableTips As Boolean 'local copy

Public Sub KillUpdate()
    '<EhHeader>
    On Error GoTo KillUpdate_Err
    '</EhHeader>
100     WinUI.Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "UpdateEnabled", "False"
104     mvarDisableUpdate = True
    '<EhFooter>
    Exit Sub

KillUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.KillUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Let DisableTips(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableTips_Err
    '</EhHeader>
100     mvarDisableTips = vData
    '<EhFooter>
    Exit Property

DisableTips_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableTips", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableTips() As Boolean
    '<EhHeader>
    On Error GoTo DisableTips_Err
    '</EhHeader>
100     DisableTips = mvarDisableTips
    '<EhFooter>
    Exit Property

DisableTips_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableTips", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableUpdate(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableUpdate_Err
    '</EhHeader>
100     mvarDisableUpdate = vData
    '<EhFooter>
    Exit Property

DisableUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableUpdate() As Boolean
    '<EhHeader>
    On Error GoTo DisableUpdate_Err
    '</EhHeader>
100     DisableUpdate = mvarDisableUpdate
    '<EhFooter>
    Exit Property

DisableUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableFade(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableFade_Err
    '</EhHeader>
100     mvarDisableFade = vData
    '<EhFooter>
    Exit Property

DisableFade_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableFade", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableFade() As Boolean
    '<EhHeader>
    On Error GoTo DisableFade_Err
    '</EhHeader>
100     DisableFade = mvarDisableFade
    '<EhFooter>
    Exit Property

DisableFade_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableFade", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableSplash(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableSplash_Err
    '</EhHeader>
100     mvarDisableSplash = vData
    '<EhFooter>
    Exit Property

DisableSplash_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableSplash", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableSplash() As Boolean
    '<EhHeader>
    On Error GoTo DisableSplash_Err
    '</EhHeader>
100     DisableSplash = mvarDisableSplash
    '<EhFooter>
    Exit Property

DisableSplash_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableSplash", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DebugLang(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DebugLang_Err
    '</EhHeader>
100     mvarDebugLang = vData
    '<EhFooter>
    Exit Property

DebugLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugLang", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DebugLang() As Boolean
    '<EhHeader>
    On Error GoTo DebugLang_Err
    '</EhHeader>
100     DebugLang = mvarDebugLang
    '<EhFooter>
    Exit Property

DebugLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugLang", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
100     mvarEnabled = vData
    '<EhFooter>
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
100     Enabled = mvarEnabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property
