VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewISAPI 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New ISAPI Plugin"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   Icon            =   "frmNewISAPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   5520
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNewISAPIInterp 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox txtNewISAPIExt 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblBrowse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   5400
      MouseIcon       =   "frmNewISAPI.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3195
      MouseIcon       =   "frmNewISAPI.frx":0E1C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2640
      MouseIcon       =   "frmNewISAPI.frx":0F6E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblNewISAPITitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add a new ISAPI interpreter:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2010
   End
   Begin VB.Label lblNewISAPIInterp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Where is the executable that will interpret this script type?"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   4050
   End
   Begin VB.Label lblNewISAPIIExt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What is the file extension for this file type?"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   2955
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New ISAPI Plugin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2850
   End
   Begin VB.Line Line1 
      X1              =   6330
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H00804008&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6330
   End
End
Attribute VB_Name = "frmNewISAPI"
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

Private Sub Form_Load()
    lblNewISAPITitle.Caption = Translator.GetText("Add a new CGI interpreter:")
    lblNewISAPIInterp.Caption = Translator.GetText("Where is the executable that will interpret this script type?")
    lblNewISAPIIExt.Caption = Translator.GetText("What is the file extension for this file type?")
    lblBrowse.Caption = Translator.GetText("&Browse")
    lblOK.Caption = Translator.GetText("&OK")
    lblCancel.Caption = Translator.GetText("&Cancel")
End Sub

Private Sub lblBrowse_Click()
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("ISAPI Plgin Files (*.dll)|*.dll|All Files (*.*)|*.*")
    dlgMain.ShowSave
    If dlgMain.FileName <> "" Then
        txtNewISAPIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblOK_Click()
    If txtNewISAPIInterp.Text <> "" And txtNewISAPIExt.Text <> "" Then
        Core.Server.HTTP.Config.ISAPI.Add txtNewISAPIInterp.Text, txtNewISAPIExt.Text, txtNewISAPIExt.Text
        Unload Me
    Else
        MsgBox Translator.GetText("Please fill all fields.")
    End If
End Sub
