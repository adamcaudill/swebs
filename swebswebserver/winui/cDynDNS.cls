VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDynDNS"
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
Private mvarHostName As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPassword As String 'local copy
Private mvarLastUpdate As String 'local copy
Private mvarLastResult As String 'local copy
Private mvarLastIP As String 'local copy

Public Sub Save()
    '<EhHeader>
    On Error GoTo Save_Err
    '</EhHeader>
100     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", mvarHostName
104     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", mvarLastIP
108     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", mvarLastResult
112     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", mvarLastUpdate
116     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", mvarPassword
120     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", mvarUserName
    
124     If mvarEnabled = True Then
128         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
        Else
132         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
        End If
    '<EhFooter>
    Exit Sub

Save_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Save", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub Reload()
    '<EhHeader>
    On Error GoTo Reload_Err
    '</EhHeader>
    Dim strResult As String
    
100     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSEnable")
104     If LCase$(strResult) = "true" Then
108         mvarEnabled = True
        Else
112         mvarEnabled = False
        End If
116     mvarHostName = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSHostname")
120     mvarLastIP = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastIP")
124     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastResult")
128     If strResult = "" Then
132         mvarLastResult = "(None)"
        Else
136         mvarLastResult = strResult
        End If
140     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastUpdate")
144     If strResult = "" Then
148         mvarLastUpdate = CDate(2.00001)
        Else
152         mvarLastUpdate = CDate(strResult)
        End If
156     mvarPassword = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSPassword")
160     mvarUserName = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSUsername")
    '<EhFooter>
    Exit Sub

Reload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Reload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Get LastIP() As String
    '<EhHeader>
    On Error GoTo LastIP_Err
    '</EhHeader>
100     LastIP = mvarLastIP
    '<EhFooter>
    Exit Property

LastIP_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastIP", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let LastIP(ByVal Value As String)
    '<EhHeader>
    On Error GoTo LastIP_Err
    '</EhHeader>
100     mvarLastIP = Value
    '<EhFooter>
    Exit Property

LastIP_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastIP", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get LastResult() As String
    '<EhHeader>
    On Error GoTo LastResult_Err
    '</EhHeader>
100     LastResult = mvarLastResult
    '<EhFooter>
    Exit Property

LastResult_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastResult", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let LastResult(ByVal Value As String)
    '<EhHeader>
    On Error GoTo LastResult_Err
    '</EhHeader>
100     mvarLastResult = Value
    '<EhFooter>
    Exit Property

LastResult_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastResult", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get LastUpdate() As String
    '<EhHeader>
    On Error GoTo LastUpdate_Err
    '</EhHeader>
100     LastUpdate = mvarLastUpdate
    '<EhFooter>
    Exit Property

LastUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let LastUpdate(ByVal Value As String)
    '<EhHeader>
    On Error GoTo LastUpdate_Err
    '</EhHeader>
100     mvarLastUpdate = Value
    '<EhFooter>
    Exit Property

LastUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.LastUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Password() As String
    '<EhHeader>
    On Error GoTo Password_Err
    '</EhHeader>
100     Password = mvarPassword
    '<EhFooter>
    Exit Property

Password_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Password", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Password(ByVal Value As String)
    '<EhHeader>
    On Error GoTo Password_Err
    '</EhHeader>
100     mvarPassword = Value
    '<EhFooter>
    Exit Property

Password_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Password", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get UserName() As String
    '<EhHeader>
    On Error GoTo UserName_Err
    '</EhHeader>
100     UserName = mvarUserName
    '<EhFooter>
    Exit Property

UserName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.UserName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let UserName(ByVal Value As String)
    '<EhHeader>
    On Error GoTo UserName_Err
    '</EhHeader>
100     mvarUserName = Value
    '<EhFooter>
    Exit Property

UserName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.UserName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get HostName() As String
    '<EhHeader>
    On Error GoTo HostName_Err
    '</EhHeader>
100     HostName = mvarHostName
    '<EhFooter>
    Exit Property

HostName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.HostName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let HostName(ByVal Value As String)
    '<EhHeader>
    On Error GoTo HostName_Err
    '</EhHeader>
100     mvarHostName = Value
    '<EhFooter>
    Exit Property

HostName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.HostName", Erl, False
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
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
100     mvarEnabled = Value
    '<EhFooter>
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDynDNS.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property
