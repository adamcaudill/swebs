VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventLog"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cCommonDialog"
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

Private blnEnabled As Boolean
Private blnChanged As Boolean
Private strEventLog As String

Public Property Get Log() As String
    blnChanged = False
    Log = strEventLog
End Property

Public Property Get Changed() As Boolean
    Changed = blnChanged
End Property

Public Sub AddEvent(ByVal strLocation As String, ByVal strEvent As String)
Attribute AddEvent.VB_UserMemId = 0
    If blnEnabled = True Then
        blnChanged = True
        strEventLog = strEventLog & "(" & Format$(Now, "hh:mm:ss") & ") " & strLocation & ": " & strEvent & vbCrLf
    Else
        strEventLog = ""
    End If
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    blnEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = blnEnabled
End Property
