VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SWEBS Web Server - Control Center"
   ClientHeight    =   5025
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6945
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStatus 
      Interval        =   750
      Left            =   2160
      Top             =   4560
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   2400
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet netMain 
      Left            =   2640
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   51
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   48
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Server Status"
      TabPicture(0)   =   "frmMain.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgLogo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLogo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lneLogo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraSrvStatus"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraUpdate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Configuration"
      TabPicture(1)   =   "frmMain.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sstConfig"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logs "
      TabPicture(2)   =   "frmMain.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtViewLogFiles"
      Tab(2).Control(1)=   "cmbViewLogFiles"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraUpdate 
         Caption         =   "Update Status:"
         Height          =   1095
         Left            =   3360
         TabIndex        =   52
         Top             =   480
         Width           =   3255
         Begin VB.Label lblUpdateStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "New Version Available"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   660
            MouseIcon       =   "frmMain.frx":0D1E
            MousePointer    =   99  'Custom
            TabIndex        =   55
            ToolTipText     =   "Click here for details."
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblUpdateVersion 
            Caption         =   "Update Version: 0.00.0000"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblCurVersion 
            Caption         =   "Current Version: 0.00.0000"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.TextBox txtViewLogFiles 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   36
         Text            =   "frmMain.frx":1028
         Top             =   840
         Width           =   6495
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":104E
         Left            =   -74880
         List            =   "frmMain.frx":1050
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Width           =   6495
      End
      Begin VB.Frame fraSrvStatus 
         Caption         =   "Current Service Status:"
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3135
         Begin VB.PictureBox picSrvButtons 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   2895
            TabIndex        =   56
            Top             =   600
            Width           =   2895
            Begin VB.CommandButton cmdSrvStart 
               Caption         =   "Start"
               Height          =   375
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdSrvStop 
               Caption         =   "Stop"
               Height          =   375
               Left            =   960
               TabIndex        =   58
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdSrvRestart 
               Caption         =   "Restart"
               Height          =   375
               Left            =   1920
               TabIndex        =   57
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Label lblSrvStatusCur 
            Caption         =   "<current-status>"
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblSrvStatus 
            Caption         =   "Status: "
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
      End
      Begin TabDlg.SSTab sstConfig 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Basic"
         TabPicture(0)   =   "frmMain.frx":1052
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblServerName"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblPort"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblWebroot"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblLogFile"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtServerName"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtPort"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtWebroot"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdBrowseRoot"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdBrowseLogFile"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtLogFile"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Advanced"
         TabPicture(1)   =   "frmMain.frx":106E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblMaxConnect"
         Tab(1).Control(1)=   "lblIndexFiles"
         Tab(1).Control(2)=   "lblAllowIndex"
         Tab(1).Control(3)=   "lblErrorPages"
         Tab(1).Control(4)=   "txtMaxConnect"
         Tab(1).Control(5)=   "txtIndexFiles"
         Tab(1).Control(6)=   "txtAllowIndex"
         Tab(1).Control(7)=   "txtErrorPages"
         Tab(1).Control(8)=   "cmdBrowseErrorPages"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "vHosts"
         TabPicture(2)   =   "frmMain.frx":108A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblvHostName"
         Tab(2).Control(1)=   "lblvHostDomain"
         Tab(2).Control(2)=   "lblvHostRoot"
         Tab(2).Control(3)=   "lblvHostLog"
         Tab(2).Control(4)=   "lstvHosts"
         Tab(2).Control(5)=   "txtvHostName"
         Tab(2).Control(6)=   "txtvHostDomain"
         Tab(2).Control(7)=   "txtvHostRoot"
         Tab(2).Control(8)=   "txtvHostLog"
         Tab(2).Control(9)=   "cmdBrowsevHostRoot"
         Tab(2).Control(10)=   "cmdBrowsevHostLog"
         Tab(2).Control(11)=   "cmdvHostNew"
         Tab(2).Control(12)=   "cmdvHostRemove"
         Tab(2).ControlCount=   13
         TabCaption(3)   =   "CGI Handlers"
         TabPicture(3)   =   "frmMain.frx":10A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lblCGIInterp"
         Tab(3).Control(1)=   "lblCGIExt"
         Tab(3).Control(2)=   "lstCGI"
         Tab(3).Control(3)=   "txtCGIInterp"
         Tab(3).Control(4)=   "txtCGIExt"
         Tab(3).Control(5)=   "cmdBrowseCGIInterp"
         Tab(3).Control(6)=   "cmdCGINew"
         Tab(3).Control(7)=   "cmdCGIRemove"
         Tab(3).ControlCount=   8
         Begin VB.CommandButton cmdvHostRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71880
            TabIndex        =   50
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdCGIRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71880
            TabIndex        =   49
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdBrowseErrorPages 
            Caption         =   "..."
            Height          =   255
            Left            =   -69000
            TabIndex        =   47
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtErrorPages 
            Height          =   285
            Left            =   -74760
            TabIndex        =   46
            Top             =   3480
            Width           =   5655
         End
         Begin VB.CommandButton cmdvHostNew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   -72960
            TabIndex        =   44
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdCGINew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   -72960
            TabIndex        =   43
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox txtLogFile 
            Height          =   285
            Left            =   240
            TabIndex        =   40
            Top             =   3360
            Width           =   5655
         End
         Begin VB.CommandButton cmdBrowseLogFile 
            Caption         =   "..."
            Height          =   255
            Left            =   6000
            TabIndex        =   39
            Top             =   3360
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseCGIInterp 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69120
            TabIndex        =   38
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostLog 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69120
            TabIndex        =   37
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostRoot 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69120
            TabIndex        =   34
            Top             =   1920
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseRoot 
            Caption         =   "..."
            Height          =   255
            Left            =   6000
            TabIndex        =   33
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox txtCGIExt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   28
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtCGIInterp 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   26
            Top             =   840
            Width           =   3615
         End
         Begin VB.ListBox lstCGI 
            Height          =   2985
            ItemData        =   "frmMain.frx":10C2
            Left            =   -74880
            List            =   "frmMain.frx":10C4
            TabIndex        =   24
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtvHostLog 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   23
            Top             =   2520
            Width           =   3615
         End
         Begin VB.TextBox txtvHostRoot 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   21
            Top             =   1920
            Width           =   3615
         End
         Begin VB.TextBox txtvHostDomain 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   19
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtvHostName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            TabIndex        =   17
            Top             =   720
            Width           =   2415
         End
         Begin VB.ListBox lstvHosts 
            Height          =   2985
            ItemData        =   "frmMain.frx":10C6
            Left            =   -74880
            List            =   "frmMain.frx":10C8
            TabIndex        =   15
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtAllowIndex 
            Height          =   285
            Left            =   -74760
            TabIndex        =   11
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtIndexFiles 
            Height          =   285
            Left            =   -74760
            TabIndex        =   10
            Top             =   2640
            Width           =   5655
         End
         Begin VB.TextBox txtMaxConnect 
            Height          =   285
            Left            =   -74760
            TabIndex        =   9
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtWebroot 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   2400
            Width           =   5655
         End
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtServerName 
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   5655
         End
         Begin VB.Label lblErrorPages 
            Caption         =   "Where is the location of the folder which stores pages to be used when the server receives an error."
            Height          =   495
            Left            =   -74880
            TabIndex        =   45
            Top             =   3000
            Width           =   5895
         End
         Begin VB.Label lblLogFile 
            Caption         =   "This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here."
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   2880
            Width           =   6135
         End
         Begin VB.Label lblCGIExt 
            Caption         =   "What is the extension that is mapped to this interpreter."
            Height          =   255
            Left            =   -72960
            TabIndex        =   27
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label lblCGIInterp 
            Caption         =   "Where is the executable that will interpret these CGI scripts?"
            Height          =   255
            Left            =   -72960
            TabIndex        =   25
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label lblvHostLog 
            Caption         =   "Where do you want to keep the log file for this vHost?"
            Height          =   255
            Left            =   -72960
            TabIndex        =   22
            Top             =   2280
            Width           =   4095
         End
         Begin VB.Label lblvHostRoot 
            Caption         =   "This is the root directory where files are kept for this vHost."
            Height          =   255
            Left            =   -72960
            TabIndex        =   20
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Label lblvHostDomain 
            Caption         =   "What is it's domain name?"
            Height          =   255
            Left            =   -72960
            TabIndex        =   18
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lblvHostName 
            Caption         =   "What is the name of this vHost?"
            Height          =   255
            Left            =   -72960
            TabIndex        =   16
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblAllowIndex 
            Caption         =   "This allows the server print out a list of all the files in the folder if no index file can be found."
            Height          =   495
            Left            =   -74880
            TabIndex        =   14
            Top             =   1320
            Width           =   6135
         End
         Begin VB.Label lblIndexFiles 
            Caption         =   $"frmMain.frx":10CA
            Height          =   495
            Left            =   -74880
            TabIndex        =   13
            Top             =   2160
            Width           =   6135
         End
         Begin VB.Label lblMaxConnect 
            Caption         =   "What is the maximum number of connections that your server can handle at any one time."
            Height          =   495
            Left            =   -74880
            TabIndex        =   12
            Top             =   480
            Width           =   6255
         End
         Begin VB.Label lblWebroot 
            Caption         =   $"frmMain.frx":1178
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   5895
         End
         Begin VB.Label lblPort 
            Caption         =   "What port do you want to use? (Default is 80)"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   6135
         End
         Begin VB.Label lblServerName 
            Caption         =   "What is the name of your server?"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Line lneLogo 
         X1              =   3120
         X2              =   6480
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label lblLogo 
         Caption         =   "SWEBS Web Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   42
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Left            =   3120
         Picture         =   "frmMain.frx":121C
         Top             =   3720
         Width           =   480
      End
   End
   Begin VB.Label lblAppStatus 
      Caption         =   "Current App Status..."
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Data..."
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export Settings..."
      End
      Begin VB.Menu mnuFileReload 
         Caption         =   "&Reload Data..."
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHomePage 
         Caption         =   "SWEBS Home Page..."
      End
      Begin VB.Menu mnuHelpForum 
         Caption         =   "SWEBS Forum..."
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUpdate 
         Caption         =   "Check for Update..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim blnDirty As Boolean 'if true then assume that some bit of data has changed

Private Sub cmbViewLogFiles_Click()
Dim strLog As String
Dim strTemp As String
    AppStatus True, "Loading Log File..."
    If Dir$(cmbViewLogFiles.Text) <> "" Then
        Open cmbViewLogFiles.Text For Random As 1 Len = FileLen(cmbViewLogFiles.Text)
            Get #1, , strLog
        Close 1
    Else
        DoEvents
        MsgBox "File not found, it may not have been created yet."
    End If
    AppStatus False
End Sub

Private Sub cmdApply_Click()
    If SaveConfigData(strConfigFile) = False Then
        MsgBox "Data was not saved, no idea why..."
    Else
        MsgBox "You data has been saved." & vbCrLf & vbCrLf & "You will need to restart the SWEBS Service before these setting will take effect.", vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdBrowseCGIInterp_Click()
Dim strDefaultFile As String
    blnDirty = True
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Executable Files (*.exe)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid$(Config.CGI((lstCGI.ListIndex + 1), 1), (InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Config.CGI((lstCGI.ListIndex + 1), 1), 1, (Len(Config.CGI((lstCGI.ListIndex + 1), 1)) - InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtCGIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdBrowseErrorPages_Click()
Dim strPath As String
    blnDirty = True
    strPath = BrowseForFolder(Me, , True, Config.ErrorPages)
    If strPath <> "" Then
        txtErrorPages.Text = strPath
    End If
End Sub

Private Sub cmdBrowseRoot_Click()
Dim strPath As String
    blnDirty = True
    strPath = BrowseForFolder(Me, , True, Config.WebRoot)
    If strPath <> "" Then
        txtWebroot.Text = strPath
    End If
End Sub

Private Sub cmdBrowsevHostLog_Click()
Dim strDefaultFile As String
    blnDirty = True
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid$(Config.vHost((lstvHosts.ListIndex + 1), 4), (InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Config.vHost((lstvHosts.ListIndex + 1), 4), 1, (Len(Config.vHost((lstvHosts.ListIndex + 1), 4)) - InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtvHostLog.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdBrowsevHostRoot_Click()
Dim strPath As String
    strPath = BrowseForFolder(Me, , True, Config.vHost((lstvHosts.ListIndex + 1), 3))
    If strPath <> "" Then
        txtvHostRoot.Text = strPath
    End If
End Sub

Private Sub cmdBrowseLogFile_Click()
Dim strDefaultFile As String
    blnDirty = True
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid$(Config.LogFile, (InStrRev(Config.LogFile, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Config.LogFile, 1, (Len(Config.LogFile) - InStrRev(Config.LogFile, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtLogFile.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCGINew_Click()
'this is a kludgey way to do this, but it works for now...
'i'll replace this with a form someday soon
Dim strNewExt As String
Dim strNewInterp As String
Dim i As Long

    blnDirty = True
    strNewInterp = InputBox("Where is the executable that will interpret this script type?")
    If strNewInterp = "" Then Exit Sub
    strNewExt = InputBox("What is the file extension for this file type?")
    If strNewExt = "" Then Exit Sub
    AddNewCGI strNewExt, strNewInterp
    If Config.CGI(1, 2) <> "" Then
        lstCGI.Clear
        For i = 1 To UBound(Config.CGI)
            lstCGI.AddItem Config.CGI(i, 2)
        Next
    Else
        lstCGI.Enabled = False
    End If
End Sub

Private Sub cmdCGIRemove_Click()
Dim lngRetVal As Long
Dim i As Long

    If lstCGI.ListIndex >= 0 Then
        lngRetVal = MsgBox("Are you sure you want to delete this item?" & vbCrLf & vbCrLf & "This can not be undone.", vbQuestion + vbYesNo)
        If lngRetVal = vbYes Then
            blnDirty = True
            RemoveCGI (lstCGI.ListIndex + 1)
            lstCGI.Clear
            If Config.CGI(1, 2) <> "" Then
                For i = 1 To UBound(Config.CGI)
                    lstCGI.AddItem Config.CGI(i, 2)
                Next
            Else
                lstCGI.Enabled = False
                cmdBrowseCGIInterp.Enabled = False
                cmdCGIRemove.Enabled = False
                txtCGIInterp.Enabled = False
                txtCGIExt.Enabled = False
                txtCGIInterp.Text = ""
                txtCGIExt.Text = ""
            End If
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSrvRestart_Click()
    AppStatus True, "Restarting Service..."
    ServiceStop "", "SWS Web Server"
    Do Until lblSrvStatusCur.Caption = "Stopped"
        DoEvents
    Loop
    ServiceStart "", "SWS Web Server"
    AppStatus False
End Sub

Private Sub cmdSrvStart_Click()
    AppStatus True, "Starting Service..."
    ServiceStart "", "SWS Web Server"
    AppStatus False
End Sub

Private Sub cmdSrvStop_Click()
    AppStatus True, "Stopping Service..."
    ServiceStop "", "SWS Web Server"
    AppStatus False
End Sub

Private Sub cmdvHostNew_Click()
'this is a kludgey way to do this, but it works for now...
'i'll replace this with a form someday soon
Dim strNewName As String
Dim strNewDomain As String
Dim strNewRoot As String
Dim strNewLog As String
Dim i As Long

    blnDirty = True
    strNewName = InputBox("What is the name of this vHost?")
    If strNewName = "" Then Exit Sub
    strNewDomain = InputBox("What is the domain for this vHost?")
    If strNewDomain = "" Then Exit Sub
    strNewRoot = InputBox("Where is the root folder for this vHost?")
    If strNewRoot = "" Then Exit Sub
    strNewLog = InputBox("Where do you want to keep the log for this vHost?")
    If strNewLog = "" Then Exit Sub
    AddNewvHost strNewName, strNewDomain, strNewRoot, strNewLog
    lstvHosts.Clear
    If Config.vHost(1, 1) <> "" Then
        For i = 1 To UBound(Config.vHost)
            lstvHosts.AddItem Config.vHost(i, 1)
        Next
        lstvHosts.Enabled = True
    Else
        lstvHosts.Enabled = False
    End If
End Sub

Private Sub cmdvHostRemove_Click()
Dim lngRetVal As Long
Dim i As Long

    If lstvHosts.ListIndex >= 0 Then
        lngRetVal = MsgBox("Are you sure you want to delete this item?" & vbCrLf & vbCrLf & "This can not be undone.", vbQuestion + vbYesNo)
        If lngRetVal = vbYes Then
            blnDirty = True
            RemovevHost (lstvHosts.ListIndex + 1)
            lstvHosts.Clear
            If Config.vHost(1, 1) <> "" Then
                For i = 1 To UBound(Config.vHost)
                    lstvHosts.AddItem Config.vHost(i, 1)
                Next
            Else
                cmdBrowsevHostRoot.Enabled = False
                cmdBrowsevHostLog.Enabled = False
                cmdvHostRemove.Enabled = False
                txtvHostName.Enabled = False
                txtvHostDomain.Enabled = False
                txtvHostRoot.Enabled = False
                txtvHostLog.Enabled = False
                txtvHostName.Text = ""
                txtvHostDomain.Text = ""
                txtvHostRoot.Text = ""
                txtvHostLog.Text = ""
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long
    If LoadConfigData = False Then
        RetVal = MsgBox("There was an error while loading your configuration data." & vbCrLf & vbCrLf & "Press 'Abort' to give up and exit, 'Retry' to try to load th data again," & vbCrLf & "or 'Ignore' to continue.", vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                End
            Case vbRetry
                If LoadConfigData = False Then
                    MsgBox "A second attempt to load your configuration data failed. Aborting." & vbCrLf & vbCrLf & "This application will now close.", vbApplicationModal + vbCritical
                    End
                End If
            Case vbIgnore
                MsgBox "NOTICE: You have chosen to proceed after a data error," & vbCrLf & "this application may not function properly or you may loose data."
        End Select
    End If
    GetUpdateInfo
    lblCurVersion.Caption = "Current Version: " & strInstalledVer
    lblUpdateVersion.Caption = "Update Version: " & IIf(Update.Version <> "", Update.Version, strInstalledVer)
    If Update.Available = True Then
        lblUpdateStatus.Caption = "New Version Available"
    Else
        lblUpdateStatus.Caption = "No Updates Available"
        lblUpdateStatus.Font.Underline = False
        lblUpdateStatus.ForeColor = vbButtonText
        lblUpdateStatus.MousePointer = vbDefault
    End If
    tmrStatus_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lngRetVal As Long
    If blnDirty = True Then
        lngRetVal = MsgBox("Do you want to save your settings before closing?", vbYesNo + vbQuestion + vbApplicationModal)
        If lngRetVal = vbYes Then
            If SaveConfigData(strConfigFile) = False Then
                MsgBox "Data was not saved, no idea why..."
            End If
        End If
    End If
    Me.Visible = False
    DoEvents
End Sub

Private Sub lblUpdateStatus_Click()
    If Update.Available = True Then
        Load frmUpdate
        frmUpdate.Show
    End If
End Sub

Private Sub lstCGI_Click()
    cmdBrowseCGIInterp.Enabled = True
    cmdCGIRemove.Enabled = True
    txtCGIInterp.Enabled = True
    txtCGIExt.Enabled = True
    txtCGIInterp.Text = Config.CGI((lstCGI.ListIndex + 1), 1)
    txtCGIExt.Text = Config.CGI((lstCGI.ListIndex + 1), 2)
End Sub

Private Sub lstvHosts_Click()
    cmdBrowsevHostRoot.Enabled = True
    cmdBrowsevHostLog.Enabled = True
    cmdvHostRemove.Enabled = True
    txtvHostName.Enabled = True
    txtvHostDomain.Enabled = True
    txtvHostRoot.Enabled = True
    txtvHostLog.Enabled = True
    txtvHostName.Text = Config.vHost((lstvHosts.ListIndex + 1), 1)
    txtvHostDomain.Text = Config.vHost((lstvHosts.ListIndex + 1), 2)
    txtvHostRoot.Text = Config.vHost((lstvHosts.ListIndex + 1), 3)
    txtvHostLog.Text = Config.vHost((lstvHosts.ListIndex + 1), 4)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExport_Click()
    'this needs some kind of error control, file checks, etc..
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgMain.ShowSave
    If dlgMain.FileName <> "" Then
        Open dlgMain.FileName For Append As 1
            Print #1, GetConfigReport
        Close 1
    End If
End Sub

Private Sub mnuFileReload_Click()
Dim RetVal As Long
    RetVal = MsgBox("This will reset any changes you make." & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion)
    If RetVal = vbYes Then
        If LoadConfigData = False Then
            RetVal = MsgBox("There was an error while loading your configuration data." & vbCrLf & vbCrLf & "Press 'Abort' to give up and exit, 'Retry' to try to load th data again," & vbCrLf & "or 'Ignore' to continue.", vbCritical + vbAbortRetryIgnore + vbApplicationModal)
            Select Case RetVal
                Case vbAbort
                    Unload Me
                Case vbRetry
                    If LoadConfigData = False Then
                        MsgBox "A second attempt to load your configuration data failed. Aborting.", vbApplicationModal + vbCritical
                    End If
                Case vbIgnore
                    MsgBox "NOTICE: You have chosen to proceed after a data error, this application may" & vbCrLf & "not function properly or you may loose data."
            End Select
        End If
    End If
End Sub

Private Sub mnuFileSave_Click()
    If SaveConfigData(strConfigFile) = False Then
        MsgBox "Data was not saved, no idea why..."
    Else
        MsgBox "You data has been saved." & vbCrLf & vbCrLf & "You will need to restart the SWEBS Service before these setting will take effect.", vbOKOnly + vbInformation
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpForum_Click()
    OpenURL "http://swebs.sourceforge.net/community/index.php"
End Sub

Private Sub mnuHelpHomePage_Click()
    OpenURL "http://swebs.sourceforge.net/"
End Sub

Private Sub mnuHelpUpdate_Click()
    AppStatus True, "Retrieving Update Information..."
    GetUpdateInfo
    If Update.Available = True Then
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox "You have the most current version available.", vbOKOnly + vbInformation
    End If
    AppStatus False
End Sub

Private Sub tmrStatus_Timer()
Dim strSrvStatusCur As String
    strSrvStatusCur = ServiceStatus("", "SWS Web Server")
    lblSrvStatusCur.Font.Bold = False
    Select Case strSrvStatusCur
        Case "Stopped"
            lblSrvStatusCur.Caption = "Stopped"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Start Pending"
            lblSrvStatusCur.Caption = "Start Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Stop Pending"
            lblSrvStatusCur.Caption = "Stop Pending"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = "Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
        Case "Continue Pending"
            lblSrvStatusCur.Caption = "Continue Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Pause Pending"
            lblSrvStatusCur.Caption = "Pause Pending"
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Paused"
            lblSrvStatusCur.Caption = "Paused"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
    End Select
End Sub

Private Sub AppStatus(blnBusy As Boolean, Optional strMessage As String = "Ready...")
    If blnBusy = True Then
        Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
    Else
        Screen.MousePointer = vbDefault  '0 default
    End If
    lblAppStatus.Caption = strMessage
    DoEvents 'i'm not sure if this will stay, causes the lbl to flash for fast operations...
End Sub

Private Function LoadConfigData() As Boolean
Dim i As Long
Dim strTemp As String
    
    AppStatus True, "Loading Configuration Data..."
    LoadConfigData = GetConfigData(strConfigFile)
    
    'Setup the form...
    txtServerName.Text = Config.ServerName
    txtPort.Text = Config.Port
    txtWebroot.Text = Config.WebRoot
    txtMaxConnect.Text = Config.MaxConnections
    txtLogFile.Text = Config.LogFile
    txtAllowIndex.Text = Config.AllowIndex
    txtErrorPages.Text = Config.ErrorPages
    For i = 1 To UBound(Config.Index)
        strTemp = strTemp & Config.Index(i) & " "
    Next
    txtIndexFiles.Text = Trim$(strTemp)
    If Config.CGI(1, 2) <> "" Then
        lstCGI.Clear
        For i = 1 To UBound(Config.CGI)
            lstCGI.AddItem Config.CGI(i, 2)
        Next
    Else
        lstCGI.Enabled = False
    End If
    If Config.vHost(1, 1) <> "" Then
        lstvHosts.Clear
        For i = 1 To UBound(Config.vHost)
            lstvHosts.AddItem Config.vHost(i, 1)
        Next
    Else
        lstvHosts.Enabled = False
    End If
    cmbViewLogFiles.Clear
    cmbViewLogFiles.AddItem Config.LogFile
    For i = 1 To UBound(Config.vHost)
        cmbViewLogFiles.AddItem Config.vHost(i, 4)
    Next
    AppStatus False
End Function

Private Sub txtAllowIndex_Change()
    Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
End Sub

Private Sub txtAllowIndex_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtCGIExt_Change()
    If lstCGI.ListIndex <> -1 Then
        Config.CGI((lstCGI.ListIndex + 1), 2) = txtCGIExt.Text
    End If
End Sub

Private Sub txtCGIExt_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtCGIInterp_Change()
    If lstCGI.ListIndex <> -1 Then
        Config.CGI((lstCGI.ListIndex + 1), 1) = txtCGIInterp.Text
    End If
End Sub

Private Sub txtCGIInterp_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtErrorPages_Change()
    Config.ErrorPages = txtErrorPages.Text
End Sub

Private Sub txtErrorPages_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtIndexFiles_Change()
Dim strTmpArray() As String
Dim lngRecCount As Long
Dim i As Long
    strTmpArray = Split(Trim$(txtIndexFiles.Text), " ")
    If UBound(strTmpArray) >= 1 Then
        ReDim Config.Index(1 To (UBound(strTmpArray) + 1))
        lngRecCount = UBound(strTmpArray)
        For i = 0 To lngRecCount
            Config.Index(i + 1) = strTmpArray(i)
        Next
    End If
End Sub

Private Sub txtIndexFiles_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtLogFile_Change()
    Config.LogFile = Trim$(txtLogFile.Text)
End Sub

Private Sub txtLogFile_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtMaxConnect_Change()
    Config.MaxConnections = Int(Val(txtMaxConnect.Text))
End Sub

Private Sub txtMaxConnect_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtPort_Change()
    Config.Port = Int(Val(txtPort.Text))
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtServerName_Change()
    Config.ServerName = Trim$(txtServerName.Text)
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtvHostDomain_Change()
    If lstvHosts.ListIndex <> -1 Then
        Config.vHost((lstvHosts.ListIndex + 1), 2) = txtvHostDomain.Text
    End If
End Sub

Private Sub txtvHostDomain_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtvHostLog_Change()
    If lstvHosts.ListIndex <> -1 Then
        Config.vHost((lstvHosts.ListIndex + 1), 4) = txtvHostLog.Text
    End If
End Sub

Private Sub txtvHostLog_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtvHostName_Change()
    If lstvHosts.ListIndex <> -1 Then
        Config.vHost((lstvHosts.ListIndex + 1), 1) = txtvHostName.Text
    End If
End Sub

Private Sub txtvHostName_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtvHostRoot_Change()
    If lstvHosts.ListIndex <> -1 Then
        Config.vHost((lstvHosts.ListIndex + 1), 3) = txtvHostRoot.Text
    End If
End Sub

Private Sub txtvHostRoot_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtWebroot_Change()
    Config.WebRoot = Trim$(txtWebroot.Text)
End Sub

Private Sub txtWebroot_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub GetUpdateInfo()
Dim strData As String

    'get data, this pulls it from a local file, for testing only.
    'Open strUIPath & "upgrade.xml" For Input As 1
    '    Do Until EOF(1)
    '        Line Input #1, strTemp
    '        strData = strData & strTemp & vbCrLf
    '    Loop
    'Close 1

    'get data from server
    If GetNetStatus = True Then
        strData = Replace(netMain.OpenURL("http://swebs.sf.net/upgrade.xml", icString), vbLf, vbCrLf)
    End If
    
    Call GetUpdateStatus(strData)
End Sub
