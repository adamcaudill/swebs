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

Public Stats As cHTTPStats
Public Config As cHTTPConfig

Private Declare Function WebStart Lib "sHTTP.dll" Alias "SWEBSStart" () As Long
Private Declare Function WebStop Lib "sHTTP.dll" Alias "SWEBSStop" () As Long

Private blnRunning As Boolean

'swebs.dll error codes
Private Const SWEBS_RETURN_UNKNOWN = &H0                '// Unknown error occured
Private Const SWEBS_RETURN_SUCCESS = &H1                '// Server ran fine
Private Const SWEBS_RETURN_COULDNOTBIND = &H2           '// Could not bind() to port
Private Const SWEBS_RETURN_CONFIGNOTLOADED = &H3        '// Could not load config file
Private Const SWEBS_RETURN_COULDNOTLISTEN = &H4         '// Could not listen()
Private Const SWEBS_RETURN_COULDNOTACCEPT = &H5         '// Could not accept()

Public Function StopServer() As Boolean
Dim lngRetVal As Long

    lngRetVal = WebStop
    blnRunning = False
End Function

Public Function StartServer() As Boolean
Dim lngRetVal As Long

    lngRetVal = WebStart
    If lngRetVal = SWEBS_RETURN_SUCCESS Then
        blnRunning = True
    Else
        MsgBox "Server Failed to start. Error code is: 0x" & lngRetVal
    End If
End Function

Public Property Get Status() As String
    If blnRunning = True Then
        Status = "Running"
    Else
        Status = "Stopped"
    End If
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    Set Stats = New cHTTPStats
    Set Config = New cHTTPConfig
End Sub

Private Sub Class_Terminate()
    Set Stats = Nothing
    Set Config = Nothing
End Sub
