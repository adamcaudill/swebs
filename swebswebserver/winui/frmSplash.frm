VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "SWEBS-Splash"
   ClientHeight    =   1830
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   1830
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7740
      TabIndex        =   0
      Top             =   1560
      Width           =   795
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_Click()
    '<EhHeader>
    On Error GoTo Form_Click_Err
    '</EhHeader>
100     Me.Hide
    '<EhFooter>
    Exit Sub

Form_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmSplash.Form_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo Form_KeyPress_Err
    '</EhHeader>
100     Me.Hide
    '<EhFooter>
    Exit Sub

Form_KeyPress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmSplash.Form_KeyPress", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     Me.MousePointer = vbHourglass
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmSplash.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
100     Me.MousePointer = vbDefault
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmSplash.Form_Unload", Erl, False
    Resume Next
    '</EhFooter>
End Sub
