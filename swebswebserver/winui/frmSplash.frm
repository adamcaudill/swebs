VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
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
      ForeColor       =   &H00FEFEFE&
      Height          =   210
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Width           =   4995
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Core - Custom
'***************************************************************************
'
' SWEBS/Core
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

Dim lngOriginalRgn As Long

Private Sub Form_Click()
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
    lngOriginalRgn = FormRegion(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormRemoveRegion Me.hWnd, lngOriginalRgn
    Me.MousePointer = vbDefault
End Sub
