VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  SWEBS Web Server - Control Center Update"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
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
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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

Private Sub cmdCancel_Click()
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.frmUpdate.cmdCancel_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdDownload_Click()
    '<EhHeader>
    On Error GoTo cmdDownload_Click_Err
    '</EhHeader>
100     mWinUI.Network.LaunchURL mWinUI.Update.DownloadURL
    '<EhFooter>
    Exit Sub

cmdDownload_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.frmUpdate.cmdDownload_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdMoreInfo_Click()
    '<EhHeader>
    On Error GoTo cmdMoreInfo_Click_Err
    '</EhHeader>
100     mWinUI.Network.LaunchURL mWinUI.Update.InfoURL
    '<EhFooter>
    Exit Sub

cmdMoreInfo_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.frmUpdate.cmdMoreInfo_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     lblTitle.Caption = mWinUI.GetTranslatedText("There is an update available for this software, it may have additional features, bug fixes and security updates. To maintain security and performance we recommend you always use the latest version available.")
104     lblDesc.Caption = mWinUI.GetTranslatedText("Description:")
108     cmdDownload.Caption = mWinUI.GetTranslatedText("Download Upgrade...")
112     cmdMoreInfo.Caption = mWinUI.GetTranslatedText("More Information...")
116     cmdCancel.Caption = mWinUI.GetTranslatedText("&Cancel")
120     lblDate.Caption = mWinUI.GetTranslatedText("Date") & ": " & mWinUI.Update.ReleaseDate
124     lblVersion.Caption = mWinUI.GetTranslatedText("Version") & ": " & mWinUI.Update.Version
128     lblUpdateLevel.Caption = mWinUI.GetTranslatedText("Update Level") & ": " & mWinUI.Update.UpdateLevel
132     lblFileSize.Caption = mWinUI.GetTranslatedText("File Size") & ": " & Format$(mWinUI.Update.FileSize, "###,###,###,###,###")
136     txtDesc.Text = mWinUI.Update.Description
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.frmUpdate.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub
