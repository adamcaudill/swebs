VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SWEBS Web Server"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfCredits 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      FileName        =   "C:\Documents and Settings\Adam\My Documents\Projects\swebs\swebswebserver\winui\credits.rtf"
      TextRTF         =   $"frmAbout.frx":0CCA
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      Caption         =   "swebs.sourceforge.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmAbout.frx":13AC
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Go To URL: http://swebs.sourceforge.net/"
      Top             =   5640
      Width           =   1605
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   600
      Picture         =   "frmAbout.frx":16B6
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblUIBuild 
      Alignment       =   2  'Center
      Caption         =   "Control Center Build: XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblSrvVersion 
      Alignment       =   2  'Center
      Caption         =   "Server Version: X.XX.XX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Line lneUI 
      Index           =   1
      X1              =   600
      X2              =   4200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "SWEBS Web Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1118
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdClose_Click()
    '<EhHeader>
    On Error GoTo cmdClose_Click_Err
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    Exit Sub

cmdClose_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmAbout.cmdClose_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     cmdClose.Caption = WinUI.GetTranslatedText("&Close")
104     lblSrvVersion.Caption = WinUI.GetTranslatedText("Server Version") & ": " & WinUI.Version
108     lblUIBuild.Caption = WinUI.GetTranslatedText("Control Center Build") & ": " & App.Revision
112     rtfCredits.TextRTF = Replace(rtfCredits.TextRTF, "Lang-Maintainer", WinUI.GetTranslatedText("Lang-Maintainer"))
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmAbout.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHomePage_Click()
    '<EhHeader>
    On Error GoTo lblHomePage_Click_Err
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
    '<EhFooter>
    Exit Sub

lblHomePage_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmAbout.lblHomePage_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub
