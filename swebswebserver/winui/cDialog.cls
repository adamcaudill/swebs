VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialog"
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

Private mForms As Collection
Private mstrStatus As String

Public Sub Show(strForm As String, Optional blnModal As Boolean = False)
    '<EhHeader>
    On Error GoTo Show_Err
    '</EhHeader>
100     Select Case LCase(strForm)
            Case "about"
                Dim ofrmAbout As frmAbout
104             Set ofrmAbout = New frmAbout
108             mForms.Add ofrmAbout, LCase(strForm)
112             Set ofrmAbout = Nothing
116         Case "splash"
                Dim ofrmSplash As frmSplash
120             Set ofrmSplash = New frmSplash
124             mForms.Add ofrmSplash, LCase(strForm)
128             Set ofrmSplash = Nothing
132         Case "registration"
                Dim ofrmRegistration As frmRegistration
136             Set ofrmRegistration = New frmRegistration
140             mForms.Add ofrmRegistration, LCase(strForm)
144             Set ofrmRegistration = Nothing
148         Case "update"
                Dim ofrmUpdate As frmUpdate
152             Set ofrmUpdate = New frmUpdate
156             mForms.Add ofrmUpdate, LCase(strForm)
160             Set ofrmUpdate = Nothing
164             Screen.MousePointer = vbHourglass
168             mWinUI.Update.Check
172             Screen.MousePointer = vbNormal
        End Select
176     mForms.Item(LCase(strForm)).Show IIf(blnModal = True, vbModal, 0)
    '<EhFooter>
    Exit Sub

Show_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cDialog.Show", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub Destroy(strForm As String)
    '<EhHeader>
    On Error GoTo Destroy_Err
    '</EhHeader>
    Dim objForm As Form

100     Set objForm = mForms.Item(LCase(strForm))
104     objForm.Hide
108     DoEvents
112     Unload objForm
116     mForms.Remove (LCase(strForm))
120     Set objForm = Nothing
    '<EhFooter>
    Exit Sub

Destroy_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cDialog.Destroy", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub SetStatus(strStatus As String, Optional blnBusy As Boolean = False)
    '<EhHeader>
    On Error GoTo SetStatus_Err
    '</EhHeader>
    Dim lngSplashID As Long
    Dim blnUseSplash As Boolean
    Dim i As Long

100     For i = 0 To Forms.Count - 1
104         If Forms(i).Caption = "SWEBS-Splash" Then
108             lngSplashID = i
112             blnUseSplash = True
            End If
        Next
116     If blnUseSplash = True Then
120         Forms(lngSplashID).lblStatus.Caption = strStatus
        Else
124         If blnBusy = True Then
128             Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
            Else
132             Screen.MousePointer = vbDefault  '0 default
            End If
        End If
136     mstrStatus = strStatus
140     mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cDialog.SetStatus", "App Status Message: " & strStatus
144     DoEvents
    '<EhFooter>
    Exit Sub

SetStatus_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cDialog.SetStatus", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set mForms = New Collection
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cDialog.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set mForms = Nothing
End Sub

Public Property Get Status() As String
    '<EhHeader>
    On Error GoTo Status_Err
    '</EhHeader>
100     Status = mstrStatus
    '<EhFooter>
    Exit Property

Status_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cDialog.Status", Erl, False
    Resume Next
    '</EhFooter>
End Property
