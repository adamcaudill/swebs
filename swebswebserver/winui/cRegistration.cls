VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistration"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Property Get IsRegistered() As Boolean
    '<EhHeader>
    On Error GoTo IsRegistered_Err
    '</EhHeader>
    Dim strResult As String
100     strResult = mWinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")
104     If strResult <> "" Then
108         IsRegistered = True
        Else
112         IsRegistered = False
        End If
    
        'lets default to yes either way until somebody gets around to writing &%$#@*& script
116     IsRegistered = True
    '<EhFooter>
    Exit Property

IsRegistered_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cRegistration.IsRegistered", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub Start()
    '<EhHeader>
    On Error GoTo Start_Err
    '</EhHeader>
    Dim lngResult As Long
100     lngResult = MsgBox(mWinUI.GetTranslatedText("Would you like to register your software? It's fast and Free!\r\rProduct registration is used to provide the best possible service, products, and support for our users.\rWe will not contact you nor will we sell or give away any of your information.\r\rWould you like to register now?"), vbQuestion + vbYesNo + vbApplicationModal)
104     If lngResult = vbYes Then
108         mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cRegistration.Start", "Loading Registration Form"
112         Load frmRegistration
116         frmRegistration.Show vbModal
        End If
    '<EhFooter>
    Exit Sub

Start_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cRegistration.Start", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub Renew()
        'this is what should get called:
        'PageSource "http://swebs.sf.net/register/regupdate.php?email=" & UrlEncode(GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")) & "&ver=" & UrlEncode(WinUI.Version)
    '<EhHeader>
    On Error GoTo Renew_Err
    '</EhHeader>
100     DoEvents
    '<EhFooter>
    Exit Sub

Renew_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cRegistration.Renew", Erl, False
    Resume Next
    '</EhFooter>
End Sub

