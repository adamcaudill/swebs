VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private mvarPath As String
Private mvarVersion As String

Private strLang As String

Public EventLog As cEventLog
Public Registration As cRegistration
Public Update As cUpdate
Public Net As cNet
Public Util As cUtil
Public Debuger As cDebug
Public Server As cServer

Public Property Get Version() As String
    Version = mvarVersion
End Property

Public Property Get Path() As String
    Path = mvarPath
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    Set Util = New cUtil
    Set Debuger = New cDebug
    Set EventLog = New cEventLog
    Set Registration = New cRegistration
    Set Update = New cUpdate
    Set Net = New cNet
    Set Server = New cServer
End Sub

Private Sub Class_Terminate()
    If Debuger.PerfMon.Enabled = True Then
        Debuger.PerfMon.Save mvarPath & "ccperfmon.log"
    End If
    Set EventLog = Nothing
    Set Registration = Nothing
    Set Update = Nothing
    Set Net = Nothing
    Set Util = Nothing
    Set Debuger = Nothing
    Set Server = Nothing
End Sub

Public Sub Setup()
    mvarVersion = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
    mvarPath = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
    mvarPath = IIf(Right$(mvarPath, 1) <> "\", mvarPath & "\", mvarPath)
    
    Debuger.Setup
    Server.HTTP.Config.Setup
    Server.HTTP.Stats.Setup
    
    Translator.LoadLang "1033", mvarPath & "lang.xml"
    Util.LoadUser32 True
    Util.InitCommonControlsVB
End Sub
