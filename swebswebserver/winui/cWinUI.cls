VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinUI"
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
Private mvarConfigFile As String 'local copy
Private mvarStatsFile As String 'local copy
Private mvarPath As String 'local copy
Private mvarVersion As String 'local copy

Private Type tvHost
    Name As String
    Domain As String
    Root As String
    Log As String
End Type

Private Type tConfig
    Index() As String
    CGI() As String 'this needs to be converted to a UDT
    vHost() As tvHost
End Type

Public Config As cConfig
Public DynDNS As cDynDNS
Public Stats As cStats
Public EventLog As cEventLog
Public Registration As cRegistration
Public Update As cUpdate
Public Net As cNet

Public Property Get Version() As String
    '<EhHeader>
    On Error GoTo Version_Err
    '</EhHeader>
100     Version = mvarVersion
    '<EhFooter>
    Exit Property

Version_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cWinUI.Version", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Path() As String
    '<EhHeader>
    On Error GoTo Path_Err
    '</EhHeader>
100     Path = mvarPath
    '<EhFooter>
    Exit Property

Path_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cWinUI.Path", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get StatsFile() As String
    '<EhHeader>
    On Error GoTo StatsFile_Err
    '</EhHeader>
100     StatsFile = mvarStatsFile
    '<EhFooter>
    Exit Property

StatsFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cWinUI.StatsFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ConfigFile() As String
    '<EhHeader>
    On Error GoTo ConfigFile_Err
    '</EhHeader>
100     ConfigFile = mvarConfigFile
    '<EhFooter>
    Exit Property

ConfigFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cWinUI.ConfigFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set EventLog = New cEventLog
104     Set Stats = New cStats
108     Set Registration = New cRegistration
112     Set Update = New cUpdate
116     Set Net = New cNet
120     Set DynDNS = New cDynDNS
124     Set Config = New cConfig
    
128     mvarConfigFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
132     mvarStatsFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
136     mvarVersion = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
140     mvarPath = GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cWinUI.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub
