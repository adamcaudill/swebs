VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About SWEBS Web Server"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   ForeColor       =   &H00000000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0CCA
   ScaleHeight     =   6795
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCreditsScroll 
      Interval        =   50
      Left            =   5160
      Top             =   0
   End
   Begin VB.PictureBox picCreditsScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   5640
      ScaleHeight     =   5775
      ScaleWidth      =   2175
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin RichTextLib.RichTextBox rtfCredits 
         Height          =   7695
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   13573
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MousePointer    =   1
         Appearance      =   0
         FileName        =   "D:\MyDocs\Projects\swebs\swebswebserver\winui\credits.rtf"
         TextRTF         =   $"frmAbout.frx":6960
      End
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Close"
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
      Left            =   7225
      MouseIcon       =   "frmAbout.frx":6DA4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   5140
      MouseIcon       =   "frmAbout.frx":6EF6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Go To URL: http://swebs.sourceforge.net/"
      Top             =   6120
      Width           =   1605
   End
   Begin VB.Label lblUIBuild 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label lblSrvVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub Form_Unload(Cancel As Integer)
    FormRemoveRegion Me.hWnd, lngOriginalRgn
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
    lngOriginalRgn = FormRegion(Me)
    lblClose.Caption = Translator.GetText("&Close")
    lblSrvVersion.Caption = Translator.GetText("Server Version") & ": " & Core.Version
    lblUIBuild.Caption = Translator.GetText("Control Center Build") & ": " & App.Revision
    rtfCredits.TextRTF = Replace(rtfCredits.TextRTF, "Lang-Maintainer", Translator.GetText("Lang-Maintainer"))
    rtfCredits.Top = rtfCredits.Height * -1
End Sub

Private Sub lblHomePage_Click()
    Core.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub tmrCreditsScroll_Timer()
    rtfCredits.Top = rtfCredits.Top - 12
    If rtfCredits.Top < rtfCredits.Height * -1 Then
        rtfCredits.Top = picCreditsScroll.Height + 10
    End If
    picCreditsScroll.Refresh
End Sub
