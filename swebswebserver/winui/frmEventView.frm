VERSION 5.00
Begin VB.Form frmEventView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SWEBS Web Server - Control Center Event Viewer"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmEventView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEvents 
      Interval        =   750
      Left            =   2040
      Top             =   3000
   End
   Begin VB.TextBox txtEvents 
      Height          =   1695
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmEventView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     WinUI.EventLog.Clear
104     WinUI.EventLog.Enabled = True
108     WinUI.EventLog.AddEvent "WinUI.frmEventView.Form_Load", "Event Viewer Loaded"
112     Form_Resize
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmEventView.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    txtEvents.Move 0, 0, (Me.ScaleWidth), (Me.ScaleHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
100     WinUI.EventLog.Enabled = False
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmEventView.Form_Unload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub tmrEvents_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    If txtEvents.Text <> WinUI.EventLog.Log Then
        txtEvents.Text = WinUI.EventLog.Log
        txtEvents.SelStart = Len(txtEvents.Text)
    End If
End Sub
