VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SWEBS Web Server"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame fraCredits 
      Caption         =   "Credits:"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
      Begin VB.Line lneUI 
         Index           =   3
         X1              =   240
         X2              =   4200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line lneUI 
         Index           =   2
         X1              =   240
         X2              =   4200
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line lneUI 
         Index           =   0
         X1              =   240
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblNames 
         Caption         =   "Adam Caudill"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblNames 
         Caption         =   "Windows UI Developer / Windows Packager"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label lblNames 
         Caption         =   "UNIX/Linux Maintainer"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label lblNames 
         Caption         =   "Thomas Fletcher"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblNames 
         Caption         =   "Project Manager / Lead Server Devloper"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblNames 
         Caption         =   "Paul Stovell"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblUIBuild 
      Caption         =   "Control Center Build: XXXX"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lblSrvVersion 
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Line lneUI 
      Index           =   1
      X1              =   120
      X2              =   4560
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
      Left            =   840
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
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblSrvVersion.Caption = "Server Version: " & strInstalledVer
    lblUIBuild.Caption = "Control Center Build: " & App.Revision
End Sub
