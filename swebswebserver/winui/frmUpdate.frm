VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  SWEBS Web Server - Control Center Update"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdMoreInfo 
      Caption         =   "More Information..."
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download Upgrade..."
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame fraDetails 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8775
      Begin VB.TextBox txtDesc 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   840
         Width           =   8535
      End
      Begin VB.Label lblFileSize 
         Caption         =   "File Size: 0,000,000"
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblUpdateLevel 
         Caption         =   "Update Level: 0000"
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version: 00.00.0000"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Caption         =   "Date: 00/00/0000"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   $"frmUpdate.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDownload_Click()
    OpenURL Update.DownloadURL
End Sub

Private Sub cmdMoreInfo_Click()
    OpenURL Update.InfoURL
End Sub

Private Sub Form_Load()
    lblDate.Caption = "Date: " & Update.Date
    lblVersion.Caption = "Version: " & Update.Version
    lblUpdateLevel.Caption = "Update Level: " & Update.UpdateLevel
    lblFileSize.Caption = "File Size: " & Format$(Update.FileSize, "###,###,###,###,###")
    txtDesc.Text = Update.Description
End Sub
