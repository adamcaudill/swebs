VERSION 5.00
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.0#0"; "vbalExpBar6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SWEBS Web Server - Control Center"
   ClientHeight    =   4290
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   5040
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraLogs 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   40
      Top             =   0
      Width           =   6975
      Begin RichTextLib.RichTextBox rtfViewLogFiles 
         Height          =   3255
         Left            =   120
         TabIndex        =   108
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5741
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":0CCA
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Height          =   315
         ItemData        =   "frmMain.frx":0D4C
         Left            =   120
         List            =   "frmMain.frx":0D4E
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Frame fraStatus 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   10
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   6615
         TabIndex        =   85
         Top             =   120
         Width           =   6615
         Begin VB.Frame fraBasicStats 
            Caption         =   "Basic Stats:"
            Height          =   1095
            Left            =   0
            TabIndex        =   97
            Top             =   1200
            Width           =   3135
            Begin VB.Label lblStatsBytesSent 
               Caption         =   "Total Bytes Sent: 000,000,000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label lblStatsRequestCount 
               Caption         =   "Request Count: 000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label lblStatsLastRestart 
               Caption         =   "Last Restart: 00/00/0000 00:00:00PM"
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraUpdate 
            Caption         =   "Update Status:"
            Height          =   1095
            Left            =   3240
            TabIndex        =   86
            Top             =   0
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
               MouseIcon       =   "frmMain.frx":0D50
               MousePointer    =   99  'Custom
               TabIndex        =   89
               ToolTipText     =   "Click here for details."
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label lblUpdateVersion 
               Caption         =   "Update Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label lblCurVersion 
               Caption         =   "Current Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraSrvStatus 
            Caption         =   "Current Service Status:"
            Height          =   1095
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   3135
            Begin VB.PictureBox picSrvButtons 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   2895
               TabIndex        =   91
               Top             =   600
               Width           =   2895
               Begin VB.CommandButton cmdSrvStart 
                  Caption         =   "Start"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   94
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdSrvStop 
                  Caption         =   "Stop"
                  Height          =   375
                  Left            =   960
                  TabIndex        =   93
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdSrvRestart 
                  Caption         =   "Restart"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   92
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.Label lblSrvStatusCur 
               Caption         =   "<current-status>"
               Height          =   255
               Left            =   720
               TabIndex        =   96
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label lblSrvStatus 
               Caption         =   "Status: "
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   615
            End
         End
      End
      Begin VB.Line lneLogo 
         X1              =   3360
         X2              =   6840
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Left            =   3360
         Picture         =   "frmMain.frx":105A
         Top             =   3120
         Width           =   480
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
         Left            =   3960
         TabIndex        =   5
         Top             =   3240
         Width           =   2895
      End
   End
   Begin VB.Frame fraConfigvHost 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   24
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   3
         Left            =   6480
         ScaleHeight     =   855
         ScaleWidth      =   255
         TabIndex        =   50
         Top             =   1680
         Width           =   255
         Begin VB.CommandButton cmdBrowsevHostLog 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   52
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostRoot 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   2040
         ScaleHeight     =   375
         ScaleWidth      =   2055
         TabIndex        =   47
         Top             =   3240
         Width           =   2055
         Begin VB.CommandButton cmdvHostRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   49
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdvHostNew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.ListBox lstvHosts 
         Height          =   3375
         ItemData        =   "frmMain.frx":1D24
         Left            =   120
         List            =   "frmMain.frx":1D26
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtvHostName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtvHostDomain 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtvHostRoot 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtvHostLog 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label lblvHostName 
         Caption         =   "What is the name of this Virtual Host?"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblvHostDomain 
         Caption         =   "What is it's domain name?"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblvHostRoot 
         Caption         =   "This is the root directory where files are kept for this Virtual Host."
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblvHostLog 
         Caption         =   "Where do you want to keep the log file for this Virtual Host?"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   2040
         Width           =   4335
      End
   End
   Begin VB.Frame fraConfigBasic 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   6360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   106
         Top             =   1440
         Width           =   255
         Begin VB.CommandButton cmdBrowseErrorLog 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtConfigBasicErrorLog 
         Height          =   285
         Left            =   240
         TabIndex        =   105
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   5
         Left            =   6360
         ScaleHeight     =   1215
         ScaleWidth      =   255
         TabIndex        =   55
         Top             =   2400
         Width           =   255
         Begin VB.CommandButton cmdBrowseLogFile 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseRoot 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtWebroot 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   6015
      End
      Begin VB.TextBox txtLogFile 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   6015
      End
      Begin VB.Label lblConfigBasicErrorLog 
         Caption         =   "Where do you want to store the server error log?"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label lblLogFile 
         Caption         =   "This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here."
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   6135
      End
      Begin VB.Label lblServerName 
         Caption         =   "What is the name of your server?"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblPort 
         Caption         =   "What port do you want to use? (Default is 80)"
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblWebroot 
         Caption         =   $"frmMain.frx":1D28
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   6135
      End
   End
   Begin VB.Frame fraConfigAdv 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   15
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtConfigAdvIPBind 
         Height          =   285
         Left            =   240
         TabIndex        =   103
         Top             =   1560
         Width           =   2295
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   53
         Top             =   3240
         Width           =   255
         Begin VB.CommandButton cmdBrowseErrorPages 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtMaxConnect 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtIndexFiles 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   5655
      End
      Begin VB.TextBox txtAllowIndex 
         Height          =   285
         Left            =   4320
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtErrorPages 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label lblConfigAdvIPBind 
         Caption         =   "What IP should the server listen to? (Default: Leave blank for all available)"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblMaxConnect 
         Caption         =   "What is the maximum number of connections that your server can handle at any one time."
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblIndexFiles 
         Caption         =   $"frmMain.frx":1DCC
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label lblAllowIndex 
         Caption         =   "Display file list if no index is found?"
         Height          =   495
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblErrorPages 
         Caption         =   "Where is the location of the folder which stores pages to be used when the server receives an error."
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   5895
      End
   End
   Begin VB.Timer tmrStats 
      Interval        =   60000
      Left            =   5520
      Top             =   3840
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbaSideBar 
      Height          =   4215
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7435
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin VB.Timer tmrStatus 
      Interval        =   5000
      Left            =   5280
      Top             =   3840
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame fraNewvHost 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   58
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2280
         ScaleHeight     =   375
         ScaleWidth      =   2175
         TabIndex        =   82
         Top             =   3240
         Width           =   2175
         Begin VB.CommandButton cmdNewvHostOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   0
            TabIndex        =   84
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdNewvHostCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1200
            TabIndex        =   83
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdBrowseNewvHostRoot 
         Caption         =   "..."
         Height          =   255
         Left            =   5880
         TabIndex        =   69
         Top             =   2160
         Width           =   255
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   6
         Left            =   5880
         ScaleHeight     =   855
         ScaleWidth      =   255
         TabIndex        =   68
         Top             =   2160
         Width           =   255
         Begin VB.CommandButton cmdBrowseNewvHostLogs 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewvHostLogs 
         Height          =   285
         Left            =   600
         TabIndex        =   67
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostRoot 
         Height          =   285
         Left            =   600
         TabIndex        =   65
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostDomain 
         Height          =   285
         Left            =   600
         TabIndex        =   64
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtNewvHostName 
         Height          =   285
         Left            =   600
         TabIndex        =   61
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblNewvHostLogs 
         Caption         =   "Where do you want to keep the log for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   66
         Top             =   2520
         Width           =   5295
      End
      Begin VB.Label lblNewvHostDomain 
         Caption         =   "What is the domain for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   63
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblNewvHostRoot 
         Caption         =   "Where is the root folder for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label lblNewvHostName 
         Caption         =   "What is the name of this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   60
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label lblNewvHostTitle 
         Caption         =   "Add a new Virtual Host:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraNewCGI 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   71
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2520
         ScaleHeight     =   375
         ScaleWidth      =   2055
         TabIndex        =   79
         Top             =   3120
         Width           =   2055
         Begin VB.CommandButton cmdNewCGICancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1080
            TabIndex        =   81
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdNewCGIOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   5880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   77
         Top             =   960
         Width           =   255
         Begin VB.CommandButton cmdBrowseNewCGIInterp 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewCGIExt 
         Height          =   285
         Left            =   1080
         TabIndex        =   76
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNewCGIInterp 
         Height          =   285
         Left            =   1080
         TabIndex        =   74
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label lblNewCGIExt 
         Caption         =   "What is the file extension for this file type?"
         Height          =   255
         Left            =   840
         TabIndex        =   75
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lblNewCGIInterp 
         Caption         =   "Where is the executable that will interpret this script type?"
         Height          =   255
         Left            =   840
         TabIndex        =   73
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label lblNewCGITitle 
         Caption         =   "Add a new CGI interpreter:"
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame fraConfigCGI 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   34
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2040
         ScaleHeight     =   375
         ScaleWidth      =   2055
         TabIndex        =   44
         Top             =   3240
         Width           =   2055
         Begin VB.CommandButton cmdCGIRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   46
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdCGINew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   5880
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   42
         Top             =   600
         Width           =   375
         Begin VB.CommandButton cmdBrowseCGIInterp 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.ListBox lstCGI 
         Height          =   3375
         ItemData        =   "frmMain.frx":1E7A
         Left            =   120
         List            =   "frmMain.frx":1E7C
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCGIInterp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtCGIExt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblCGIInterp 
         Caption         =   "Where is the executable that will interpret these CGI scripts?"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblCGIExt 
         Caption         =   "What is the extension that is mapped to this interpreter."
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   1080
         Width           =   4815
      End
   End
   Begin VB.Label lblAppStatus 
      Caption         =   "Ready..."
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   3960
      Width           =   3135
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
      Begin VB.Menu mnuHelpRegister 
         Caption         =   "&Register..."
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpEventViewer 
         Caption         =   "Open Event &Viewer..."
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuSysTrayPopup 
      Caption         =   "mnuSysTrayPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayPopupOpenCC 
         Caption         =   "&Open Control Center..."
      End
      Begin VB.Menu mnuSpacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTrayPopupHomePage 
         Caption         =   "SWEBS Home Page..."
      End
      Begin VB.Menu mnuSysTrayPopupForum 
         Caption         =   "SWEBS Forum..."
      End
      Begin VB.Menu mnuSpacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTrayPopupUpdate 
         Caption         =   "Check for Update..."
      End
      Begin VB.Menu mnuSpacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysTrayPopupAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuSysTrayPopupExit 
         Caption         =   "E&xit..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private WithEvents SysTray As cSysTray
Attribute SysTray.VB_VarHelpID = -1

Dim blnDirty As Boolean 'if true then assume that some bit of data has changed

Private Sub cmbViewLogFiles_Click()
Dim strLog As String
    
    SetStatus WinUI.GetTranslatedText("Loading Log File") & "...", True
    If Dir$(cmbViewLogFiles.Text) <> "" Then
        DoEvents
        strLog = Space$(FileLen(cmbViewLogFiles.Text))
        Open cmbViewLogFiles.Text For Binary As 1
            Get #1, 1, strLog
        Close 1
        rtfViewLogFiles.Text = Replace(strLog, vbCr, "")
        rtfViewLogFiles.RightMargin = 214748364
        rtfViewLogFiles.SetFocus
    Else
        DoEvents
        MsgBox WinUI.GetTranslatedText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
    End If
    SetStatus "Ready..."
End Sub

Private Sub cmdApply_Click()
    If WinUI.Server.HTTP.Config.Save(WinUI.ConfigFile) = False Then
        MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
    Else
        blnDirty = False
        MsgBox WinUI.GetTranslatedText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdBrowseCGIInterp_Click()
Dim strDefaultFile As String
    blnDirty = True
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
    strDefaultFile = Mid$(WinUI.Server.HTTP.Config.CGI(lstCGI.ListIndex + 1).Interpreter, (InStrRev(WinUI.Server.HTTP.Config.CGI(lstCGI.ListIndex + 1).Interpreter, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(WinUI.Server.HTTP.Config.CGI(lstCGI.ListIndex + 1).Interpreter, 1, (Len(WinUI.Server.HTTP.Config.CGI(lstCGI.ListIndex + 1).Interpreter) - InStrRev(WinUI.Server.HTTP.Config.CGI(lstCGI.ListIndex + 1).Interpreter, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtCGIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdBrowseErrorLog_Click()
'Dim strDefaultFile As String
'    blnDirty = True
'    dlgMain.DialogTitle = "Please select a file..."
'    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
'    strDefaultFile = Mid$(Config.vHost((WinUI.Config.ErrorLog + 1), 4), (InStrRev(Config.vHost((WinUI.Config.ErrorLog + 1), 4), "\") + 1))
'    dlgMain.FileName = strDefaultFile
'    dlgMain.InitDir = WinUI.Path
'    dlgMain.ShowSave
'    If dlgMain.FileName <> strDefaultFile Then
'        txtvHostLog.Text = dlgMain.FileName
'    End If
End Sub

Private Sub cmdBrowseErrorPages_Click()
Dim strPath As String
    blnDirty = True
    strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.ErrorPages)
    If strPath <> "" Then
        txtErrorPages.Text = strPath
    End If
End Sub

Private Sub cmdBrowseNewCGIInterp_Click()
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Executable Files (*.exe)|*.log|All Files (*.*)|*.*"
    dlgMain.ShowSave
    If dlgMain.FileName <> "" Then
        txtNewCGIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdBrowseNewvHostLogs_Click()
    blnDirty = True
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    dlgMain.InitDir = WinUI.Path
    dlgMain.ShowSave
    txtvHostLog.Text = dlgMain.FileName
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
Dim strPath As String
    strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.WebRoot)
    If strPath <> "" Then
        txtNewvHostRoot.Text = strPath
    End If
End Sub

Private Sub cmdBrowseRoot_Click()
Dim strPath As String
    blnDirty = True
    strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.WebRoot)
    If strPath <> "" Then
        txtWebroot.Text = strPath
    End If
End Sub

Private Sub cmdBrowsevHostLog_Click()
'Dim strDefaultFile As String
'    blnDirty = True
'    dlgMain.DialogTitle = "Please select a file..."
'    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
'    strDefaultFile = Mid$(WinUI.Config.vHost((lstvHosts.ListIndex + 1), 4), (InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\") + 1))
'    dlgMain.FileName = strDefaultFile
'    dlgMain.InitDir = Mid$(Config.vHost((lstvHosts.ListIndex + 1), 4), 1, (Len(Config.vHost((lstvHosts.ListIndex + 1), 4)) - InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\")))
'    dlgMain.ShowSave
'    If dlgMain.FileName <> strDefaultFile Then
'        txtvHostLog.Text = dlgMain.FileName
'    End If
End Sub

Private Sub cmdBrowsevHostRoot_Click()
Dim strPath As String
    strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.VirtHost((lstvHosts.ListIndex + 1)).Root)
    If strPath <> "" Then
        txtvHostRoot.Text = strPath
    End If
End Sub

Private Sub cmdBrowseLogFile_Click()
'Dim strDefaultFile As String
'    blnDirty = True
'    dlgMain.DialogTitle = "Please select a file..."
'    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
'    strDefaultFile = Mid$(Config.LogFile, (InStrRev(Config.LogFile, "\") + 1))
'    dlgMain.FileName = strDefaultFile
'    dlgMain.InitDir = Mid$(Config.LogFile, 1, (Len(Config.LogFile) - InStrRev(Config.LogFile, "\")))
'    dlgMain.ShowSave
'    If dlgMain.FileName <> strDefaultFile Then
'        txtLogFile.Text = dlgMain.FileName
'    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCGINew_Click()
    fraNewCGI.ZOrder 0
    vbaSideBar.ZOrder 0
End Sub

Private Sub cmdCGIRemove_Click()
'***this needs replaced
'Dim lngRetVal As Long
'Dim i As Long
'
'    If lstCGI.ListIndex >= 0 Then
'        lngRetVal = MsgBox(WinUI.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
'        If lngRetVal = vbYes Then
'            blnDirty = True
'            RemoveCGI (lstCGI.ListIndex + 1)
'            lstCGI.Clear
'            If WinUI.Config.CGI(1, 2) <> "" Then
'                For i = 1 To UBound(WinUI.Config.CGI)
'                    lstCGI.AddItem WinUI.Config.CGI(i, 2)
'                Next
'            Else
'                lstCGI.Enabled = False
'                cmdBrowseCGIInterp.Enabled = False
'                cmdCGIRemove.Enabled = False
'                txtCGIInterp.Enabled = False
'                txtCGIExt.Enabled = False
'                txtCGIInterp.Text = ""
'                txtCGIExt.Text = ""
'            End If
'        End If
'    End If
End Sub

Private Sub cmdNewCGICancel_Click()
    fraNewCGI.ZOrder 1
    txtNewCGIInterp.Text = ""
    txtNewCGIExt.Text = ""
End Sub

Private Sub cmdNewCGIOK_Click()
'***this needs replaced
'Dim i As Long
'
'    If txtNewCGIInterp.Text <> "" And txtNewCGIExt.Text <> "" Then
'        blnDirty = True
'        AddNewCGI txtNewCGIExt.Text, txtNewCGIInterp.Text
'        If WinUI.Config.CGI(1, 2) <> "" Then
'            lstCGI.Clear
'            For i = 1 To UBound(WinUI.Config.CGI)
'                lstCGI.AddItem WinUI.Config.CGI(i, 2)
'            Next
'        Else
'            lstCGI.Enabled = False
'        End If
'        fraNewCGI.ZOrder 1
'        txtNewCGIInterp.Text = ""
'        txtNewCGIExt.Text = ""
'    Else
'        MsgBox WinUi.GetTranslatedText("Please fill all fields.")
'    End If
End Sub

Private Sub cmdNewvHostCancel_Click()
    fraNewvHost.ZOrder 1
    txtNewvHostName.Text = ""
    txtNewvHostDomain.Text = ""
    txtNewvHostRoot.Text = ""
    txtNewvHostLogs.Text = ""
End Sub

Private Sub cmdNewvHostOK_Click()
'***this needs replaced
'Dim i As Long
'
'    If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
'        blnDirty = True
'        AddNewvHost txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text
'        lstvHosts.Clear
'        If WinUI.Config.vHost(1).Name <> "" Then
'            For i = 1 To UBound(WinUI.Config.vHost)
'                lstvHosts.AddItem WinUI.Config.vHost(i).Name
'            Next
'            lstvHosts.Enabled = True
'        Else
'            lstvHosts.Enabled = False
'        End If
'        fraNewvHost.ZOrder 1
'        txtNewvHostName.Text = ""
'        txtNewvHostDomain.Text = ""
'        txtNewvHostRoot.Text = ""
'        txtNewvHostLogs.Text = ""
'    Else
'        MsgBox WinUi.GetTranslatedText("Please fill all fields.")
'    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSrvRestart_Click()
'    WinUI.Dialog.SetStatus WinUI.GetTranslatedText("Restarting Service") & "...", True
'    ServiceStop "", "SWEBS Web Server"
'    Do Until ServiceStatus("", "SWEBS Web Server") = "Stopped"
'        DoEvents
'    Loop
'    ServiceStart "", "SWEBS Web Server"
'    UpdateStats
'    WinUI.Dialog.SetStatus "Ready..."
End Sub

Private Sub cmdSrvStart_Click()
'    WinUI.Dialog.SetStatus WinUI.GetTranslatedText("Starting Service") & "...", True
'    ServiceStart "", "SWEBS Web Server"
'    UpdateStats
'    WinUI.Dialog.SetStatus "Ready..."
End Sub

Private Sub cmdSrvStop_Click()
'    WinUI.Dialog.SetStatus WinUI.GetTranslatedText("Stopping Service") & "...", True
'    ServiceStop "", "SWEBS Web Server"
'    WinUI.Dialog.SetStatus "Ready..."
End Sub

Private Sub cmdvHostNew_Click()
    fraNewvHost.ZOrder 0
    vbaSideBar.ZOrder 0
End Sub

Private Sub cmdvHostRemove_Click()
Dim lngRetVal As Long
Dim blnMore As Boolean
Dim vItem As Variant

    If lstvHosts.ListIndex >= 0 Then
        lngRetVal = MsgBox(WinUI.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
        If lngRetVal = vbYes Then
            blnDirty = True
            WinUI.Server.HTTP.Config.VirtHost.Remove lstvHosts.Text
            txtvHostName.Text = ""
            txtvHostDomain.Text = ""
            txtvHostRoot.Text = ""
            txtvHostLog.Text = ""
            lstvHosts.Clear
            For Each vItem In WinUI.Server.HTTP.Config.VirtHost
                lstvHosts.AddItem vItem.HostName
                blnMore = True
            Next
            If blnMore = False Then
                cmdBrowsevHostRoot.Enabled = False
                cmdBrowsevHostLog.Enabled = False
                cmdvHostRemove.Enabled = False
                txtvHostName.Enabled = False
                txtvHostDomain.Enabled = False
                txtvHostRoot.Enabled = False
                txtvHostLog.Enabled = False
                lstvHosts.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long
Dim cBar As cExplorerBar
Dim cItem As cExplorerBarItem

    'setup the translated strings...
    SetStatus "Loading Translated Strings..."
    
    mnuFile.Caption = WinUI.GetTranslatedText("&File")
    mnuFileSave.Caption = WinUI.GetTranslatedText("Save Data") & "..."
    mnuFileExport.Caption = WinUI.GetTranslatedText("Export Setings") & "..."
    mnuFileExit.Caption = WinUI.GetTranslatedText("E&xit")
    mnuHelp.Caption = WinUI.GetTranslatedText("&Help")
    mnuHelpHomePage.Caption = WinUI.GetTranslatedText("SWEBS Home Page") & "..."
    mnuHelpForum.Caption = WinUI.GetTranslatedText("SWEBS Forum") & "..."
    mnuHelpUpdate.Caption = WinUI.GetTranslatedText("Check For Update") & "..."
    mnuHelpRegister.Caption = WinUI.GetTranslatedText("Register") & "..."
    mnuHelpAbout.Caption = WinUI.GetTranslatedText("&About") & "..."
    cmdOK.Caption = WinUI.GetTranslatedText("&OK")
    cmdApply.Caption = WinUI.GetTranslatedText("&Apply")
    cmdCancel.Caption = WinUI.GetTranslatedText("&Cancel")
    fraSrvStatus.Caption = WinUI.GetTranslatedText("Current Service Status:")
    lblSrvStatus.Caption = WinUI.GetTranslatedText("Status:")
    cmdSrvStart.Caption = WinUI.GetTranslatedText("S&tart")
    cmdSrvStop.Caption = WinUI.GetTranslatedText("St&op")
    cmdSrvRestart.Caption = WinUI.GetTranslatedText("R&estart")
    fraUpdate.Caption = WinUI.GetTranslatedText("Update Status:")
    fraBasicStats.Caption = WinUI.GetTranslatedText("Basic Stats:")
    lblMaxConnect.Caption = WinUI.GetTranslatedText("What is the maximum number of connections that your server can handle at any one time.")
    lblAllowIndex.Caption = WinUI.GetTranslatedText("Display file list if no index is found?")
    lblIndexFiles.Caption = WinUI.GetTranslatedText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
    lblErrorPages.Caption = WinUI.GetTranslatedText("Where is the location of the folder which stores pages to be used when the server receives an error.")
    lblServerName.Caption = WinUI.GetTranslatedText("What is the name of your server?")
    lblPort.Caption = WinUI.GetTranslatedText("What port do you want to use? (Default is 80)")
    lblWebroot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
    lblLogFile.Caption = WinUI.GetTranslatedText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
    lblCGIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret these CGI scripts?")
    lblCGIExt.Caption = WinUI.GetTranslatedText("What is the extension that is mapped to this interpreter.")
    cmdCGINew.Caption = WinUI.GetTranslatedText("Add New...")
    cmdCGIRemove.Caption = WinUI.GetTranslatedText("Remove...")
    cmdvHostNew.Caption = WinUI.GetTranslatedText("Add New...")
    cmdvHostRemove.Caption = WinUI.GetTranslatedText("Remove...")
    lblvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
    lblvHostDomain.Caption = WinUI.GetTranslatedText("What is it's domain name?")
    lblvHostRoot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept for this Virtual Host.")
    lblvHostLog.Caption = WinUI.GetTranslatedText("Where do you want to keep the log file for this Virtual Host?")
    lblNewCGITitle.Caption = WinUI.GetTranslatedText("Add a new CGI interpreter:")
    lblNewCGIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret this script type?")
    lblNewCGIExt.Caption = WinUI.GetTranslatedText("What is the file extension for this file type?")
    cmdNewCGIOK.Caption = WinUI.GetTranslatedText("&OK")
    cmdNewCGICancel.Caption = WinUI.GetTranslatedText("&Cancel")
    lblNewvHostTitle.Caption = WinUI.GetTranslatedText("Add a new Virtual Host:")
    lblNewvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
    lblNewvHostDomain.Caption = WinUI.GetTranslatedText("What is the domain for this Virtual Host?")
    lblNewvHostRoot.Caption = WinUI.GetTranslatedText("Where is the root folder for this Virtual Host?")
    lblNewvHostLogs.Caption = WinUI.GetTranslatedText("Where do you want to keep the log for this Virtual Host?")
    cmdNewvHostOK.Caption = WinUI.GetTranslatedText("&OK")
    cmdNewvHostCancel.Caption = WinUI.GetTranslatedText("&Cancel")
    lblConfigAdvIPBind.Caption = WinUI.GetTranslatedText("What IP should the server listen to? (Default: Leave blank for all available)")
    lblConfigBasicErrorLog.Caption = WinUI.GetTranslatedText("Where do you want to store the server error log?")
    
    If LoadConfigData = False Then
        RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                End
            Case vbRetry
                If LoadConfigData = False Then
                    MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                    End
                End If
            Case vbIgnore
                MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
        End Select
    End If
    
    With vbaSideBar
        .Redraw = False
        Set cBar = .Bars.Add(, "status", WinUI.GetTranslatedText("System Status"))
        Set cItem = cBar.Items.Add(, "status", WinUI.GetTranslatedText("Current Status"), 0)
        
        Set cBar = .Bars.Add(, "config", WinUI.GetTranslatedText("Configuration"))
        Set cItem = cBar.Items.Add(, "basic", WinUI.GetTranslatedText("Basic"), 0)
        Set cItem = cBar.Items.Add(, "advanced", WinUI.GetTranslatedText("Advanced"), 0)
        Set cItem = cBar.Items.Add(, "vhost", WinUI.GetTranslatedText("Virtual Host"), 0)
        Set cItem = cBar.Items.Add(, "cgi", WinUI.GetTranslatedText("CGI"), 0)
        
        Set cBar = .Bars.Add(, "logs", WinUI.GetTranslatedText("System Logs"))
        Set cItem = cBar.Items.Add(, "logs", WinUI.GetTranslatedText("View Logs"), 0)
        .Height = Me.Height
        .Redraw = True
    End With
    
    Set SysTray = New cSysTray
    Set SysTray.SourceWindow = Me
    SysTray.IconInSysTray
    SysTray.ToolTip = "SWEBS Web Server " & WinUI.Version
    SysTray.Icon = Me.Icon

    fraStatus.ZOrder 0
    vbaSideBar.ZOrder 0
    tmrStatus_Timer
    SetStatus "Ready..."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lngRetVal As Long
Dim i As Long

    If blnDirty = True Then
        lngRetVal = MsgBox(WinUI.GetTranslatedText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
        If lngRetVal = vbYes Then
            If WinUI.Server.HTTP.Config.Save(WinUI.ConfigFile) = False Then
                MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
            End If
        End If
    End If
    
    SysTray.RemoveFromSysTray
    Set SysTray = Nothing
    DoEvents
    Me.Hide
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next
    WinUI.Util.LoadUser32 False
    Set WinUI = Nothing
    'SetExceptionFilter False
    End
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub lblUpdateStatus_Click()
    If WinUI.Update.IsAvailable = True Then
        Load frmUpdate
        frmUpdate.Show
    End If
End Sub

Private Sub lstCGI_Click()
    cmdBrowseCGIInterp.Enabled = True
    cmdCGIRemove.Enabled = True
    txtCGIInterp.Enabled = True
    txtCGIExt.Enabled = True
    txtCGIInterp.Text = WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Interpreter
    txtCGIExt.Text = WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Extention
End Sub

Private Sub lstvHosts_Click()
    cmdBrowsevHostRoot.Enabled = True
    cmdBrowsevHostLog.Enabled = True
    cmdvHostRemove.Enabled = True
    txtvHostName.Enabled = True
    txtvHostDomain.Enabled = True
    txtvHostRoot.Enabled = True
    txtvHostLog.Enabled = True
    txtvHostName.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName
    txtvHostDomain.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain
    txtvHostRoot.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root
    txtvHostLog.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log
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
            Print #1, WinUI.Server.HTTP.Config.Report
        Close 1
    End If
End Sub

Private Sub mnuFileReload_Click()
Dim RetVal As Long
    RetVal = MsgBox(WinUI.GetTranslatedText("This will reset any changes you make.\r\rDo you want to continue?"), vbYesNo + vbQuestion)
    If RetVal = vbYes Then
        If LoadConfigData = False Then
            RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
            Select Case RetVal
                Case vbAbort
                    Unload Me
                Case vbRetry
                    If LoadConfigData = False Then
                        MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                    End If
                Case vbIgnore
                    MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    End If
End Sub

Private Sub mnuFileSave_Click()
    If WinUI.Server.HTTP.Config.Save(WinUI.ConfigFile) = False Then
        MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
    Else
        blnDirty = False
        MsgBox WinUI.GetTranslatedText("You data has been saved./r/rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuHelpEventViewer_Click()
    Load frmEventView
    frmEventView.Show
End Sub

Private Sub mnuHelpForum_Click()
    WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
End Sub

Private Sub mnuHelpHomePage_Click()
    WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub mnuHelpRegister_Click()
    WinUI.Registration.Start
End Sub

Private Sub mnuHelpUpdate_Click()
    SetStatus WinUI.GetTranslatedText("Retrieving Update Information") & "...", True
    WinUI.Update.Check
    If WinUI.Update.IsAvailable = True Then
        lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
        lblUpdateStatus.Font.Underline = True
        lblUpdateStatus.ForeColor = vbBlue
        lblUpdateStatus.MousePointer = vbCustom
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox WinUI.GetTranslatedText("You have the most current version available."), vbOKOnly + vbInformation
    End If
    SetStatus "Ready..."
End Sub

Private Sub mnuSysTrayPopupAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuSysTrayPopupExit_Click()
    Unload Me
End Sub

Private Sub mnuSysTrayPopupForum_Click()
    WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
End Sub

Private Sub mnuSysTrayPopupHomePage_Click()
    WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub mnuSysTrayPopupOpenCC_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub mnuSysTrayPopupUpdate_Click()
    SetStatus WinUI.GetTranslatedText("Retrieving Update Information") & "...", True
    WinUI.Update.Check
    If WinUI.Update.IsAvailable = True Then
        lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
        lblUpdateStatus.Font.Underline = True
        lblUpdateStatus.ForeColor = vbBlue
        lblUpdateStatus.MousePointer = vbCustom
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox WinUI.GetTranslatedText("You have the most current version available."), vbOKOnly + vbInformation
    End If
    SetStatus "Ready..."
End Sub

Private Sub SysTray_LButtonDblClk()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub SysTray_RButtonUp()
    SetForegroundWindow Me.hwnd
    PopupMenu mnuSysTrayPopup, , , , mnuSysTrayPopupOpenCC
    PostMessage Me.hwnd, 0&, 0&, 0&
End Sub

Private Sub tmrStats_Timer()
    UpdateStats
End Sub

Private Sub tmrStatus_Timer()
'Dim strSrvStatusCur As String
'    strSrvStatusCur = ServiceStatus("", "SWEBS Web Server")
'    lblSrvStatusCur.Font.Bold = False
'    Select Case strSrvStatusCur
'        Case "Stopped"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Stopped")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Stopped"
'            lblSrvStatusCur.Font.Bold = True
'            lblSrvStatusCur.ForeColor = vbRed
'            cmdSrvStart.Enabled = True
'            cmdSrvStop.Enabled = False
'            cmdSrvRestart.Enabled = False
'        Case "Start Pending"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Start Pending")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Start Pending"
'            lblSrvStatusCur.ForeColor = vbYellow
'            cmdSrvStart.Enabled = False
'            cmdSrvStop.Enabled = True
'            cmdSrvRestart.Enabled = False
'        Case "Stop Pending"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Stop Pending")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Stop Pending"
'            lblSrvStatusCur.Font.Bold = True
'            lblSrvStatusCur.ForeColor = vbRed
'            cmdSrvStart.Enabled = True
'            cmdSrvStop.Enabled = False
'            cmdSrvRestart.Enabled = False
'        Case "Running"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Running")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Running"
'            lblSrvStatusCur.Font.Bold = True
'            lblSrvStatusCur.ForeColor = vbGreen
'            cmdSrvStart.Enabled = False
'            cmdSrvStop.Enabled = True
'            cmdSrvRestart.Enabled = True
'        Case "Continue Pending"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Continue Pending")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Continue Pending"
'            lblSrvStatusCur.ForeColor = vbYellow
'            cmdSrvStart.Enabled = False
'            cmdSrvStop.Enabled = True
'            cmdSrvRestart.Enabled = False
'        Case "Pause Pending"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Pause Pending")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status:  Pending"
'            lblSrvStatusCur.ForeColor = vbRed
'            cmdSrvStart.Enabled = False
'            cmdSrvStop.Enabled = True
'            cmdSrvRestart.Enabled = False
'        Case "Paused"
'            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Paused")
'            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Paused"
'            lblSrvStatusCur.Font.Bold = True
'            lblSrvStatusCur.ForeColor = vbRed
'            cmdSrvStart.Enabled = True
'            cmdSrvStop.Enabled = True
'            cmdSrvRestart.Enabled = True
'    End Select
End Sub


Private Function LoadConfigData() As Boolean
Dim strTemp As String
Dim strResult As String
Dim vItem As Variant
    
    WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.LoadConfigData", "Loading Config Data"
    SetStatus WinUI.GetTranslatedText("Loading Configuration Data") & "...", True
    LoadConfigData = WinUI.Server.HTTP.Config.LoadData
    
    'Setup the form...
    txtServerName.Text = WinUI.Server.HTTP.Config.ServerName
    txtPort.Text = WinUI.Server.HTTP.Config.Port
    txtWebroot.Text = WinUI.Server.HTTP.Config.WebRoot
    txtMaxConnect.Text = WinUI.Server.HTTP.Config.MaxConnections
    txtLogFile.Text = WinUI.Server.HTTP.Config.LogFile
    txtConfigAdvIPBind.Text = WinUI.Server.HTTP.Config.ListeningAddress
    txtAllowIndex.Text = WinUI.Server.HTTP.Config.AllowIndex
    txtErrorPages.Text = WinUI.Server.HTTP.Config.ErrorPages
    txtConfigBasicErrorLog.Text = WinUI.Server.HTTP.Config.ErrorLog
    
    For Each vItem In WinUI.Server.HTTP.Config.Index
        strTemp = strTemp & vItem.FileName & " "
    Next
    txtIndexFiles.Text = Trim$(strTemp)
    
    lstCGI.Enabled = False
    lstCGI.Clear
    For Each vItem In WinUI.Server.HTTP.Config.CGI
        lstCGI.AddItem vItem.Extention
        lstCGI.Enabled = True
    Next
    
    lstvHosts.Enabled = False
    lstvHosts.Clear
    For Each vItem In WinUI.Server.HTTP.Config.VirtHost
        lstvHosts.AddItem vItem.HostName
        lstvHosts.Enabled = True
    Next
    
    cmbViewLogFiles.Clear
    If Dir$(WinUI.Server.HTTP.Config.LogFile) <> "" Then
        cmbViewLogFiles.AddItem WinUI.Server.HTTP.Config.LogFile
    End If
    If Dir$(WinUI.Server.HTTP.Config.ErrorLog) <> "" Then
        cmbViewLogFiles.AddItem WinUI.Server.HTTP.Config.ErrorLog
    End If
    For Each vItem In WinUI.Server.HTTP.Config.VirtHost
        If Dir$(vItem.Log) <> "" Then
            cmbViewLogFiles.AddItem vItem.Log
        End If
    Next
    
    'we now only check for updates every 24 hours, this could confuse some people.
    'but this should make loading faster.
    SetStatus "Checking For Updates...", True
    strResult = WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "LastUpdateCheck")
    If strResult = "" Then
        strResult = CDate(1.1)
    End If
    If DateDiff("h", CDate(strResult), Now) >= 24 Then
        WinUI.Update.Check
        If WinUI.Update.IsAvailable = True Then
            lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
        Else
            lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
            lblUpdateStatus.Font.Underline = False
            lblUpdateStatus.ForeColor = vbButtonText
            lblUpdateStatus.MousePointer = vbDefault
            WinUI.Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "LastUpdateCheck", Now
        End If
    Else
        lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
        lblUpdateStatus.Font.Underline = False
        lblUpdateStatus.ForeColor = vbButtonText
        lblUpdateStatus.MousePointer = vbDefault
    End If
    
    UpdateStats
        
    If WinUI.Registration.IsRegistered = True Then
        SetStatus "Updating Registration..."
        mnuHelpRegister.Enabled = False
        WinUI.Registration.Renew
    End If
    
    SetStatus "Ready..."
End Function

Private Sub txtAllowIndex_Change()
    If WinUI.Server.HTTP.Config.AllowIndex <> IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false") Then
        WinUI.Server.HTTP.Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
        blnDirty = True
    End If
End Sub

Private Sub txtCGIExt_Change()
    If lstCGI.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Extention <> txtCGIExt.Text Then
            WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Extention = txtCGIExt.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtCGIInterp_Change()
    If lstCGI.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Interpreter <> txtCGIInterp.Text Then
            WinUI.Server.HTTP.Config.CGI.Item(lstCGI.Text).Interpreter = txtCGIInterp.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtConfigAdvIPBind_Change()
    If WinUI.Server.HTTP.Config.ListeningAddress = txtConfigAdvIPBind.Text Then
        WinUI.Server.HTTP.Config.ListeningAddress = txtConfigAdvIPBind.Text
        blnDirty = True
    End If
End Sub

Private Sub txtConfigBasicErrorLog_Change()
    If WinUI.Server.HTTP.Config.ErrorLog <> txtConfigBasicErrorLog.Text Then
        WinUI.Server.HTTP.Config.ErrorLog = txtConfigBasicErrorLog.Text
        blnDirty = True
    End If
End Sub

Private Sub txtErrorPages_Change()
    If WinUI.Server.HTTP.Config.ErrorPages <> txtErrorPages.Text Then
        WinUI.Server.HTTP.Config.ErrorPages = txtErrorPages.Text
        blnDirty = True
    End If
End Sub

Private Sub txtIndexFiles_Change()
'***this needs replaced
'Dim strTmpArray() As String
'Dim lngRecCount As Long
'Dim i As Long
'    strTmpArray = Split(Trim$(txtIndexFiles.Text), " ")
'    If UBound(strTmpArray) >= 1 Then
'        ReDim WinUI.Config.Index(1 To (UBound(strTmpArray) + 1))
'        lngRecCount = UBound(strTmpArray)
'        For i = 0 To lngRecCount
'            WinUI.Config.Index(i + 1) = strTmpArray(i)
'        Next
'    End If
End Sub

Private Sub txtIndexFiles_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtIndexFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDirty = True
End Sub

Private Sub txtLogFile_Change()
    If WinUI.Server.HTTP.Config.LogFile <> Trim$(txtLogFile.Text) Then
        WinUI.Server.HTTP.Config.LogFile = Trim$(txtLogFile.Text)
        blnDirty = True
    End If
End Sub

Private Sub txtMaxConnect_Change()
    If WinUI.Server.HTTP.Config.MaxConnections <> Int(Val(txtMaxConnect.Text)) Then
        WinUI.Server.HTTP.Config.MaxConnections = Int(Val(txtMaxConnect.Text))
        blnDirty = True
    End If
End Sub

Private Sub txtPort_Change()
    If WinUI.Server.HTTP.Config.Port <> Int(Val(txtPort.Text)) Then
        WinUI.Server.HTTP.Config.Port = Int(Val(txtPort.Text))
        blnDirty = True
    End If
End Sub

Private Sub txtServerName_Change()
    If WinUI.Server.HTTP.Config.ServerName <> Trim$(txtServerName.Text) Then
        WinUI.Server.HTTP.Config.ServerName = Trim$(txtServerName.Text)
        blnDirty = True
    End If
End Sub

Private Sub txtvHostDomain_Change()
    If lstvHosts.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain <> txtvHostDomain.Text Then
            WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain = txtvHostDomain.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtvHostLog_Change()
    If lstvHosts.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log <> txtvHostLog.Text Then
            WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log = txtvHostLog.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtvHostName_Change()
    If lstvHosts.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName <> txtvHostName.Text Then
            blnDirty = True
            WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName = txtvHostName.Text
        End If
    End If
End Sub

Private Sub txtvHostRoot_Change()
    If lstvHosts.ListIndex <> -1 Then
        If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root <> txtvHostRoot.Text Then
            WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root = txtvHostRoot.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtWebroot_Change()
    If WinUI.Server.HTTP.Config.WebRoot <> Trim$(txtWebroot.Text) Then
        WinUI.Server.HTTP.Config.WebRoot = Trim$(txtWebroot.Text)
        blnDirty = True
    End If
End Sub

Private Sub vbaSideBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    WinUI.Util.StopWinUpdate Me.hwnd
    Select Case itm.Key
        Case "status"
            fraStatus.ZOrder 0
        Case "basic"
            fraConfigBasic.ZOrder 0
        Case "advanced"
            fraConfigAdv.ZOrder 0
        Case "vhost"
            fraConfigvHost.ZOrder 0
        Case "cgi"
            fraConfigCGI.ZOrder 0
        Case "logs"
            fraLogs.ZOrder 0
    End Select
    vbaSideBar.ZOrder 0
    WinUI.Util.StopWinUpdate
End Sub

Private Sub UpdateStats()
    WinUI.Server.HTTP.Stats.Reload
    lblStatsLastRestart.Caption = WinUI.GetTranslatedText("Last Restart") & ": " & WinUI.Server.HTTP.Stats.LastRestart
    lblStatsRequestCount.Caption = WinUI.GetTranslatedText("Request Count") & ": " & WinUI.Server.HTTP.Stats.RequestCount
    lblStatsBytesSent.Caption = WinUI.GetTranslatedText("Total Bytes Sent") & ": " & Format$(WinUI.Server.HTTP.Stats.TotalBytesSent, "###,###,###,###,##0")
    lblCurVersion.Caption = WinUI.GetTranslatedText("Current Version") & ": " & WinUI.Version
    lblUpdateVersion.Caption = WinUI.GetTranslatedText("Update Version") & ": " & IIf(WinUI.Update.Version <> "", WinUI.Update.Version, WinUI.Version)
End Sub
