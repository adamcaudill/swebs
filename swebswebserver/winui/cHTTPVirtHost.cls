VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPVirtHost"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CSEH: WinUI - Custom
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
Private mvarHostName As String 'local copy
Private mvarDomain As String 'local copy
Private mvarRoot As String 'local copy
Private mvarLog As String 'local copy
Public Property Let Log(ByVal vData As String)
    '<EhHeader>
    On Error GoTo Log_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Log")
    '</EhHeader>
100     mvarLog = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Log_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Log", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Log() As String
    '<EhHeader>
    On Error GoTo Log_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Log")
    '</EhHeader>
100     Log = mvarLog
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Log_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Log", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Root(ByVal vData As String)
    '<EhHeader>
    On Error GoTo Root_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Root")
    '</EhHeader>
100     mvarRoot = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Root_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Root", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Root() As String
    '<EhHeader>
    On Error GoTo Root_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Root")
    '</EhHeader>
100     Root = mvarRoot
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Root_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Root", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Domain(ByVal vData As String)
    '<EhHeader>
    On Error GoTo Domain_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Domain")
    '</EhHeader>
100     mvarDomain = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Domain_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Domain", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Domain() As String
    '<EhHeader>
    On Error GoTo Domain_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.Domain")
    '</EhHeader>
100     Domain = mvarDomain
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Domain_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.Domain", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let HostName(ByVal vData As String)
    '<EhHeader>
    On Error GoTo HostName_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.HostName")
    '</EhHeader>
100     mvarHostName = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

HostName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.HostName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get HostName() As String
    '<EhHeader>
    On Error GoTo HostName_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPVirtHost.HostName")
    '</EhHeader>
100     HostName = mvarHostName
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

HostName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPVirtHost.HostName", Erl, False
    Resume Next
    '</EhFooter>
End Property



