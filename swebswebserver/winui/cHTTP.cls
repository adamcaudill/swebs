VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTP"
Attribute VB_GlobalNameSpace = False
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

Public Stats As cHTTPStats
Public Config As cHTTPConfig

Private Declare Function WebStart Lib "SWEBS.dll" Alias "SWEBSStart" () As Long
Private Declare Function WebStop Lib "SWEBS.dll" Alias "SWEBSStop" () As Long

Private blnRunning As Boolean

'swebs.dll error codes
Private Const SWEBS_RETURN_UNKNOWN = &H0                '// Unknown error occured
Private Const SWEBS_RETURN_SUCCESS = &H1                '// Server ran fine
Private Const SWEBS_RETURN_COULDNOTBIND = &H2           '// Could not bind() to port
Private Const SWEBS_RETURN_CONFIGNOTLOADED = &H3        '// Could not load config file
Private Const SWEBS_RETURN_COULDNOTLISTEN = &H4         '// Could not listen()
Private Const SWEBS_RETURN_COULDNOTACCEPT = &H5         '// Could not accept()

Public Function StopServer() As Boolean
    '<EhHeader>
    On Error GoTo StopServer_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTP.StopServer")
    '</EhHeader>
    Dim lngRetVal As Long

100     lngRetVal = WebStop
104     MsgBox "WebStop RetVal: " & lngRetVal
108     blnRunning = False
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

StopServer_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTP.StopServer", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function StartServer() As Boolean
    '<EhHeader>
    On Error GoTo StartServer_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTP.StartServer")
    '</EhHeader>
    Dim lngRetVal As Long

100     lngRetVal = WebStart
104     If lngRetVal = SWEBS_RETURN_SUCCESS Then
108         blnRunning = True
        Else
112         MsgBox "Server Failed to start. Error code is: 0x" & lngRetVal
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

StartServer_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTP.StartServer", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Get Status() As String
    '<EhHeader>
    On Error GoTo Status_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTP.Status")
    '</EhHeader>
100     If blnRunning = True Then
104         Status = "Running"
        Else
108         Status = "Stopped"
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Status_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTP.Status", Erl, False
    Resume Next
    '</EhFooter>
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set Stats = New cHTTPStats
104     Set Config = New cHTTPConfig
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTP.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set Stats = Nothing
    Set Config = Nothing
End Sub
