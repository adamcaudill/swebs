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
    WinUI.Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "UpdateEnabled", "False"
    mvarDisableUpdate = True
End Sub

Public Property Let DisableTips(ByVal vData As Boolean)
    mvarDisableTips = vData
End Property

Public Property Get DisableTips() As Boolean
    DisableTips = mvarDisableTips
End Property

Public Property Let DisableUpdate(ByVal vData As Boolean)
    mvarDisableUpdate = vData
End Property

Public Property Get DisableUpdate() As Boolean
    DisableUpdate = mvarDisableUpdate
End Property

Public Property Let DisableFade(ByVal vData As Boolean)
    mvarDisableFade = vData
End Property

Public Property Get DisableFade() As Boolean
    DisableFade = mvarDisableFade
End Property

Public Property Let DisableSplash(ByVal vData As Boolean)
    mvarDisableSplash = vData
End Property

Public Property Get DisableSplash() As Boolean
    DisableSplash = mvarDisableSplash
End Property

Public Property Let DebugLang(ByVal vData As Boolean)
    mvarDebugLang = vData
End Property

Public Property Get DebugLang() As Boolean
    DebugLang = mvarDebugLang
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property
