VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.0#0"; "vbalExpBar6.ocx"
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
         TabIndex        =   51
         Top             =   1680
         Width           =   255
         Begin VB.CommandButton cmdBrowsevHostLog 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostRoot 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   52
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
         TabIndex        =   48
         Top             =   3240
         Width           =   2055
         Begin VB.CommandButton cmdvHostRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   50
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdvHostNew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.ListBox lstvHosts 
         Height          =   3375
         ItemData        =   "frmMain.frx":0CCA
         Left            =   120
         List            =   "frmMain.frx":0CCC
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
   Begin VB.Frame fraConfigDynDns 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   103
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   3135
         TabIndex        =   119
         Top             =   960
         Width           =   3135
         Begin VB.CheckBox chkDynDNSEnable 
            Caption         =   "Enable DynDNS Updates?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   120
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   975
         TabIndex        =   117
         Top             =   3240
         Width           =   975
         Begin VB.CommandButton cmdDynDNSUpdate 
            Caption         =   "&Update"
            Height          =   375
            Left            =   0
            TabIndex        =   118
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox txtDynDNSPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "l"
         TabIndex        =   116
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtDynDNSUsername 
         Height          =   285
         Left            =   960
         TabIndex        =   114
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtDynDNSHostname 
         Height          =   285
         Left            =   960
         TabIndex        =   112
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtDynDNSLastResult 
         Height          =   285
         Left            =   4080
         TabIndex        =   110
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtDynDNSLastUpdate 
         Height          =   285
         Left            =   4080
         TabIndex        =   108
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtDynDNSCurrentIP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   106
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblDynDNSPassword 
         Caption         =   "DynDNS.org Password:"
         Height          =   255
         Left            =   3960
         TabIndex        =   115
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label lblDynDNSUsername 
         Caption         =   "DynDNS.org Username:"
         Height          =   255
         Left            =   840
         TabIndex        =   113
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblDynDNSHostname 
         Caption         =   "DynDNS.org Hostname:"
         Height          =   255
         Left            =   840
         TabIndex        =   111
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblDynDNSLastResult 
         Caption         =   "Last Update Result:"
         Height          =   255
         Left            =   3960
         TabIndex        =   109
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblDynDNSLastUpdate 
         Caption         =   "Last Update:"
         Height          =   255
         Left            =   3960
         TabIndex        =   107
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblDynDNSCurrentIP 
         Caption         =   "Current IP:"
         Height          =   255
         Left            =   840
         TabIndex        =   105
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblDynDNSTitle 
         Caption         =   $"frmMain.frx":0CCE
         Height          =   735
         Left            =   240
         TabIndex        =   104
         Top             =   240
         Width           =   6495
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
         TabIndex        =   125
         Top             =   1440
         Width           =   255
         Begin VB.CommandButton cmdBrowseErrorLog 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   126
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtConfigBasicErrorLog 
         Height          =   285
         Left            =   240
         TabIndex        =   124
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
         TabIndex        =   56
         Top             =   2400
         Width           =   255
         Begin VB.CommandButton cmdBrowseLogFile 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseRoot 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   57
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
         TabIndex        =   123
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
         Caption         =   $"frmMain.frx":0D8E
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
         TabIndex        =   122
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
         TabIndex        =   54
         Top             =   3240
         Width           =   255
         Begin VB.CommandButton cmdBrowseErrorPages 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   55
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
         TabIndex        =   121
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
         Caption         =   $"frmMain.frx":0E32
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
   Begin VB.Frame fraLogs 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   40
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtViewLogFiles 
         Height          =   3135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Top             =   480
         Width           =   6735
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Height          =   315
         ItemData        =   "frmMain.frx":0EE0
         Left            =   120
         List            =   "frmMain.frx":0EE2
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Timer tmrStats 
      Interval        =   60000
      Left            =   5520
      Top             =   3960
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
         TabIndex        =   86
         Top             =   120
         Width           =   6615
         Begin VB.Frame fraBasicStats 
            Caption         =   "Basic Stats:"
            Height          =   1095
            Left            =   0
            TabIndex        =   98
            Top             =   1200
            Width           =   3135
            Begin VB.Label lblStatsBytesSent 
               Caption         =   "Total Bytes Sent: 000,000,000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label lblStatsRequestCount 
               Caption         =   "Request Count: 000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label lblStatsLastRestart 
               Caption         =   "Last Restart: 00/00/0000 00:00:00PM"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraUpdate 
            Caption         =   "Update Status:"
            Height          =   1095
            Left            =   3240
            TabIndex        =   87
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
               MouseIcon       =   "frmMain.frx":0EE4
               MousePointer    =   99  'Custom
               TabIndex        =   90
               ToolTipText     =   "Click here for details."
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label lblUpdateVersion 
               Caption         =   "Update Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label lblCurVersion 
               Caption         =   "Current Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraSrvStatus 
            Caption         =   "Current Service Status:"
            Height          =   1095
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Width           =   3135
            Begin VB.PictureBox picSrvButtons 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   2895
               TabIndex        =   92
               Top             =   600
               Width           =   2895
               Begin VB.CommandButton cmdSrvStart 
                  Caption         =   "Start"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   95
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdSrvStop 
                  Caption         =   "Stop"
                  Height          =   375
                  Left            =   960
                  TabIndex        =   94
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdSrvRestart 
                  Caption         =   "Restart"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   93
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.Label lblSrvStatusCur 
               Caption         =   "<current-status>"
               Height          =   255
               Left            =   720
               TabIndex        =   97
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label lblSrvStatus 
               Caption         =   "Status: "
               Height          =   255
               Left            =   120
               TabIndex        =   96
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
         Picture         =   "frmMain.frx":11EE
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
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbaSideBar 
      Height          =   4215
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7435
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin VB.Timer tmrStatus 
      Interval        =   5000
      Left            =   4920
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
      TabIndex        =   59
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2280
         ScaleHeight     =   375
         ScaleWidth      =   2175
         TabIndex        =   83
         Top             =   3240
         Width           =   2175
         Begin VB.CommandButton cmdNewvHostOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdNewvHostCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1200
            TabIndex        =   84
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdBrowseNewvHostRoot 
         Caption         =   "..."
         Height          =   255
         Left            =   5880
         TabIndex        =   70
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
         TabIndex        =   69
         Top             =   2160
         Width           =   255
         Begin VB.CommandButton cmdBrowseNewvHostLogs 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewvHostLogs 
         Height          =   285
         Left            =   600
         TabIndex        =   68
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostRoot 
         Height          =   285
         Left            =   600
         TabIndex        =   66
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostDomain 
         Height          =   285
         Left            =   600
         TabIndex        =   65
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtNewvHostName 
         Height          =   285
         Left            =   600
         TabIndex        =   62
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblNewvHostLogs 
         Caption         =   "Where do you want to keep the log for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   67
         Top             =   2520
         Width           =   5295
      End
      Begin VB.Label lblNewvHostDomain 
         Caption         =   "What is the domain for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   64
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblNewvHostRoot 
         Caption         =   "Where is the root folder for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   63
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label lblNewvHostName 
         Caption         =   "What is the name of this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   61
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label lblNewvHostTitle 
         Caption         =   "Add a new Virtual Host:"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraNewCGI 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   72
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2520
         ScaleHeight     =   375
         ScaleWidth      =   2055
         TabIndex        =   80
         Top             =   3120
         Width           =   2055
         Begin VB.CommandButton cmdNewCGICancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1080
            TabIndex        =   82
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdNewCGIOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   0
            TabIndex        =   81
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
         TabIndex        =   78
         Top             =   960
         Width           =   255
         Begin VB.CommandButton cmdBrowseNewCGIInterp 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewCGIExt 
         Height          =   285
         Left            =   1080
         TabIndex        =   77
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNewCGIInterp 
         Height          =   285
         Left            =   1080
         TabIndex        =   75
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label lblNewCGIExt 
         Caption         =   "What is the file extension for this file type?"
         Height          =   255
         Left            =   840
         TabIndex        =   76
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lblNewCGIInterp 
         Caption         =   "Where is the executable that will interpret this script type?"
         Height          =   255
         Left            =   840
         TabIndex        =   74
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label lblNewCGITitle 
         Caption         =   "Add a new CGI interpreter:"
         Height          =   255
         Left            =   480
         TabIndex        =   73
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
         TabIndex        =   45
         Top             =   3240
         Width           =   2055
         Begin VB.CommandButton cmdCGIRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   47
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdCGINew 
            Caption         =   "Add New..."
            Height          =   375
            Left            =   0
            TabIndex        =   46
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
         TabIndex        =   43
         Top             =   600
         Width           =   375
         Begin VB.CommandButton cmdBrowseCGIInterp 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.ListBox lstCGI 
         Height          =   3375
         ItemData        =   "frmMain.frx":1EB8
         Left            =   120
         List            =   "frmMain.frx":1EBA
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
   Begin InetCtlsObjects.Inet netDynDNS 
      Left            =   5280
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
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

Dim blnDirty As Boolean 'if true then assume that some bit of data has changed

Private Sub chkDynDNSEnable_Click()
    '<EhHeader>
    On Error GoTo chkDynDNSEnable_Click_Err
    '</EhHeader>
100     blnDirty = True
104     If chkDynDNSEnable.Value = vbChecked Then
108         WinUI.DynDNS.Enabled = True
        Else
112         WinUI.DynDNS.Enabled = False
        End If
    '<EhFooter>
    Exit Sub

chkDynDNSEnable_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.chkDynDNSEnable_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmbViewLogFiles_Click()
    '<EhHeader>
    On Error GoTo cmbViewLogFiles_Click_Err
    '</EhHeader>
    Dim strLog As String
    
100     WinUI.Status WinUI.GetTranslatedText("Loading Log File") & "...", True
104     If Dir$(cmbViewLogFiles.Text) <> "" Then
108         strLog = Space$(FileLen(cmbViewLogFiles.Text))
112         Open cmbViewLogFiles.Text For Binary As 1
116             Get #1, 1, strLog
120         Close 1
124         txtViewLogFiles.Text = strLog
128         txtViewLogFiles.SetFocus
        Else
132         DoEvents
136         MsgBox WinUI.GetTranslatedText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
        End If
140     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Sub

cmbViewLogFiles_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmbViewLogFiles_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdApply_Click()
    '<EhHeader>
    On Error GoTo cmdApply_Click_Err
    '</EhHeader>
100     If WinUI.Config.Save(WinUI.ConfigFile) = False Then
104         MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox WinUI.GetTranslatedText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
    '<EhFooter>
    Exit Sub

cmdApply_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdApply_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseCGIInterp_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseCGIInterp_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String
    Dim strStartDir As String

100     Set cDlg = New cCommonDialog
104     strStartDir = Mid$(WinUI.Config.CGI.Item(lstCGI.Text).Interpreter, 1, (Len(WinUI.Config.CGI.Item(lstCGI.Text).Interpreter)) - InStrRev(WinUI.Config.CGI.Item(lstCGI.Text).Interpreter, "\"))
108     If cDlg.VBGetOpenFileName(strFile, , True, , , , "Executable Files (*.exe)|*.exe", , strStartDir, , "exe") Then
112         txtCGIInterp.Text = strFile
        End If
116     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowseCGIInterp_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseCGIInterp_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseErrorLog_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseErrorLog_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String

100     Set cDlg = New cCommonDialog
104     If cDlg.VBGetOpenFileName(strFile, , True, , , , "Log Files (*.log)|*.log", , , , "log") Then
108         txtConfigBasicErrorLog.Text = strFile
        End If
112     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowseErrorLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseErrorLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseErrorPages_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseErrorPages_Click_Err
    '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = BrowseForFolder(Me, , True, WinUI.Config.ErrorPages)
108     If strPath <> "" Then
112         txtErrorPages.Text = strPath
        End If
    '<EhFooter>
    Exit Sub

cmdBrowseErrorPages_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseErrorPages_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewCGIInterp_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewCGIInterp_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String

100     Set cDlg = New cCommonDialog
104     If cDlg.VBGetOpenFileName(strFile, , True, , , , "Executable Files (*.exe)|*.exe", , , , "exe") Then
108         txtNewCGIInterp.Text = strFile
        End If
112     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowseNewCGIInterp_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewCGIInterp_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewvHostLogs_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewvHostLogs_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String

100     Set cDlg = New cCommonDialog
104     If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
108         txtNewvHostLogs.Text = strFile
        End If
112     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowseNewvHostLogs_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewvHostLogs_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewvHostRoot_Click_Err
    '</EhHeader>
    Dim strPath As String
100     strPath = BrowseForFolder(Me, , True, WinUI.Config.WebRoot)
104     If strPath <> "" Then
108         txtNewvHostRoot.Text = strPath
        End If
    '<EhFooter>
    Exit Sub

cmdBrowseNewvHostRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewvHostRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseRoot_Click_Err
    '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = BrowseForFolder(Me, , True, WinUI.Config.WebRoot)
108     If strPath <> "" Then
112         txtWebroot.Text = strPath
        End If
    '<EhFooter>
    Exit Sub

cmdBrowseRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowsevHostLog_Click()
    '<EhHeader>
    On Error GoTo cmdBrowsevHostLog_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String
    Dim strStartDir As String

100     Set cDlg = New cCommonDialog
104     blnDirty = True
108     strStartDir = Mid$(WinUI.Config.vHost((lstvHosts.ListIndex + 1)).Log, (InStrRev(WinUI.Config.vHost((lstvHosts.ListIndex + 1)).Log, "\") + 1))
112     If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
116         txtvHostLog.Text = strFile
        End If
120     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowsevHostLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowsevHostLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowsevHostRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowsevHostRoot_Click_Err
    '</EhHeader>
    Dim strPath As String
100     strPath = BrowseForFolder(Me, , True, WinUI.Config.vHost((lstvHosts.ListIndex + 1)).Root)
104     If strPath <> "" Then
108         txtvHostRoot.Text = strPath
        End If
    '<EhFooter>
    Exit Sub

cmdBrowsevHostRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowsevHostRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseLogFile_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseLogFile_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String
    Dim strStartDir As String

100     Set cDlg = New cCommonDialog
104     blnDirty = True
108     strStartDir = Mid$(WinUI.Config.LogFile, (InStrRev(WinUI.Config.LogFile, "\") + 1))
112     If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
116         txtLogFile.Text = strFile
        End If
120     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

cmdBrowseLogFile_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseLogFile_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdCancel_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCGINew_Click()
    '<EhHeader>
    On Error GoTo cmdCGINew_Click_Err
    '</EhHeader>
100     fraNewCGI.ZOrder 0
104     vbaSideBar.ZOrder 0
    '<EhFooter>
    Exit Sub

cmdCGINew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdCGINew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCGIRemove_Click()
'***this needs replaced
'Dim lngRetVal As Long
'Dim i As Long
'
'    If lstCGI.ListIndex >= 0 Then
'        lngRetVal = MsgBox(WinUi.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
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
'<EhHeader>
On Error GoTo cmdCGIRemove_Click_Err
'</EhHeader>
'<EhFooter>
Exit Sub

cmdCGIRemove_Click_Err:
DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdCGIRemove_Click", Erl, False
Resume Next
'</EhFooter>
End Sub

Private Sub cmdDynDNSUpdate_Click()
    '<EhHeader>
    On Error GoTo cmdDynDNSUpdate_Click_Err
    '</EhHeader>
100     WinUI.Status "Updating DNS Information...", True
104     netDynDNS.URL = "http://members.dyndns.org"
108     netDynDNS.Document = "/nic/update?system=dyndns&hostname=" & WinUI.DynDNS.HostName & "&myip=" & WinUI.Net.CurrentIP & "&wildcard=NOCHG"
112     netDynDNS.UserName = WinUI.DynDNS.UserName
116     netDynDNS.Password = WinUI.DynDNS.Password
120     netDynDNS.Execute , "GET", , "User-Agent: SWEBS WinUI " & WinUI.Version & " <adam@imspire.com>"
    '<EhFooter>
    Exit Sub

cmdDynDNSUpdate_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdDynDNSUpdate_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNewCGICancel_Click()
    '<EhHeader>
    On Error GoTo cmdNewCGICancel_Click_Err
    '</EhHeader>
100     fraNewCGI.ZOrder 1
104     txtNewCGIInterp.Text = ""
108     txtNewCGIExt.Text = ""
    '<EhFooter>
    Exit Sub

cmdNewCGICancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewCGICancel_Click", Erl, False
    Resume Next
    '</EhFooter>
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
'<EhHeader>
On Error GoTo cmdNewCGIOK_Click_Err
'</EhHeader>
'<EhFooter>
Exit Sub

cmdNewCGIOK_Click_Err:
DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewCGIOK_Click", Erl, False
Resume Next
'</EhFooter>
End Sub

Private Sub cmdNewvHostCancel_Click()
    '<EhHeader>
    On Error GoTo cmdNewvHostCancel_Click_Err
    '</EhHeader>
100     fraNewvHost.ZOrder 1
104     txtNewvHostName.Text = ""
108     txtNewvHostDomain.Text = ""
112     txtNewvHostRoot.Text = ""
116     txtNewvHostLogs.Text = ""
    '<EhFooter>
    Exit Sub

cmdNewvHostCancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewvHostCancel_Click", Erl, False
    Resume Next
    '</EhFooter>
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
'<EhHeader>
On Error GoTo cmdNewvHostOK_Click_Err
'</EhHeader>
'<EhFooter>
Exit Sub

cmdNewvHostOK_Click_Err:
DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewvHostOK_Click", Erl, False
Resume Next
'</EhFooter>
End Sub

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvRestart_Click()
    '<EhHeader>
    On Error GoTo cmdSrvRestart_Click_Err
    '</EhHeader>
100     WinUI.Status WinUI.GetTranslatedText("Restarting Service") & "...", True
104     ServiceStop "", "SWEBS Web Server"
108     Do Until ServiceStatus("", "SWEBS Web Server") = "Stopped"
112         DoEvents
        Loop
116     ServiceStart "", "SWEBS Web Server"
120     UpdateStats
124     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Sub

cmdSrvRestart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvRestart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvStart_Click()
    '<EhHeader>
    On Error GoTo cmdSrvStart_Click_Err
    '</EhHeader>
100     WinUI.Status WinUI.GetTranslatedText("Starting Service") & "...", True
104     ServiceStart "", "SWEBS Web Server"
108     UpdateStats
112     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Sub

cmdSrvStart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvStart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvStop_Click()
    '<EhHeader>
    On Error GoTo cmdSrvStop_Click_Err
    '</EhHeader>
100     WinUI.Status WinUI.GetTranslatedText("Stopping Service") & "...", True
104     ServiceStop "", "SWEBS Web Server"
108     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Sub

cmdSrvStop_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvStop_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdvHostNew_Click()
    '<EhHeader>
    On Error GoTo cmdvHostNew_Click_Err
    '</EhHeader>
100     fraNewvHost.ZOrder 0
104     vbaSideBar.ZOrder 0
    '<EhFooter>
    Exit Sub

cmdvHostNew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdvHostNew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdvHostRemove_Click()
    '<EhHeader>
    On Error GoTo cmdvHostRemove_Click_Err
    '</EhHeader>
    Dim lngRetVal As Long
    Dim blnMore As Boolean
    Dim vItem As Variant
    Dim i As Long

100     If lstvHosts.ListIndex >= 0 Then
104         lngRetVal = MsgBox(WinUI.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
108         If lngRetVal = vbYes Then
112             blnDirty = True
116             WinUI.Config.vHost.Remove lstvHosts.Text
120             txtvHostName.Text = ""
124             txtvHostDomain.Text = ""
128             txtvHostRoot.Text = ""
132             txtvHostLog.Text = ""
136             lstvHosts.Clear
140             For Each vItem In WinUI.Config.vHost
144                 lstvHosts.AddItem vItem.HostName
148                 blnMore = True
                Next
152             If blnMore = False Then
156                 cmdBrowsevHostRoot.Enabled = False
160                 cmdBrowsevHostLog.Enabled = False
164                 cmdvHostRemove.Enabled = False
168                 txtvHostName.Enabled = False
172                 txtvHostDomain.Enabled = False
176                 txtvHostRoot.Enabled = False
180                 txtvHostLog.Enabled = False
184                 lstvHosts.Enabled = False
                End If
            End If
        End If
    '<EhFooter>
    Exit Sub

cmdvHostRemove_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdvHostRemove_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
    Dim RetVal As Long
    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

        'setup the translated strings...
100     WinUI.Status "Loading Translated Strings..."
    
104     mnuFile.Caption = WinUI.GetTranslatedText("&File")
108     mnuFileSave.Caption = WinUI.GetTranslatedText("Save Data") & "..."
112     mnuFileExport.Caption = WinUI.GetTranslatedText("Export Setings") & "..."
116     mnuFileExit.Caption = WinUI.GetTranslatedText("E&xit")
120     mnuHelp.Caption = WinUI.GetTranslatedText("&Help")
124     mnuHelpHomePage.Caption = WinUI.GetTranslatedText("SWEBS Home Page") & "..."
128     mnuHelpForum.Caption = WinUI.GetTranslatedText("SWEBS Forum") & "..."
132     mnuHelpUpdate.Caption = WinUI.GetTranslatedText("Check For Update") & "..."
136     mnuHelpRegister.Caption = WinUI.GetTranslatedText("Register") & "..."
140     mnuHelpAbout.Caption = WinUI.GetTranslatedText("&About") & "..."
144     cmdOK.Caption = WinUI.GetTranslatedText("&OK")
148     cmdApply.Caption = WinUI.GetTranslatedText("&Apply")
152     cmdCancel.Caption = WinUI.GetTranslatedText("&Cancel")
156     fraSrvStatus.Caption = WinUI.GetTranslatedText("Current Service Status:")
160     lblSrvStatus.Caption = WinUI.GetTranslatedText("Status:")
164     cmdSrvStart.Caption = WinUI.GetTranslatedText("S&tart")
168     cmdSrvStop.Caption = WinUI.GetTranslatedText("St&op")
172     cmdSrvRestart.Caption = WinUI.GetTranslatedText("R&estart")
176     fraUpdate.Caption = WinUI.GetTranslatedText("Update Status:")
180     fraBasicStats.Caption = WinUI.GetTranslatedText("Basic Stats:")
184     lblMaxConnect.Caption = WinUI.GetTranslatedText("What is the maximum number of connections that your server can handle at any one time.")
188     lblAllowIndex.Caption = WinUI.GetTranslatedText("Display file list if no index is found?")
192     lblIndexFiles.Caption = WinUI.GetTranslatedText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
196     lblErrorPages.Caption = WinUI.GetTranslatedText("Where is the location of the folder which stores pages to be used when the server receives an error.")
200     lblServerName.Caption = WinUI.GetTranslatedText("What is the name of your server?")
204     lblPort.Caption = WinUI.GetTranslatedText("What port do you want to use? (Default is 80)")
208     lblWebroot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
212     lblLogFile.Caption = WinUI.GetTranslatedText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
216     lblCGIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret these CGI scripts?")
220     lblCGIExt.Caption = WinUI.GetTranslatedText("What is the extension that is mapped to this interpreter.")
224     cmdCGINew.Caption = WinUI.GetTranslatedText("Add New...")
228     cmdCGIRemove.Caption = WinUI.GetTranslatedText("Remove...")
232     cmdvHostNew.Caption = WinUI.GetTranslatedText("Add New...")
236     cmdvHostRemove.Caption = WinUI.GetTranslatedText("Remove...")
240     lblvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
244     lblvHostDomain.Caption = WinUI.GetTranslatedText("What is it's domain name?")
248     lblvHostRoot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept for this Virtual Host.")
252     lblvHostLog.Caption = WinUI.GetTranslatedText("Where do you want to keep the log file for this Virtual Host?")
256     lblNewCGITitle.Caption = WinUI.GetTranslatedText("Add a new CGI interpreter:")
260     lblNewCGIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret this script type?")
264     lblNewCGIExt.Caption = WinUI.GetTranslatedText("What is the file extension for this file type?")
268     cmdNewCGIOK.Caption = WinUI.GetTranslatedText("&OK")
272     cmdNewCGICancel.Caption = WinUI.GetTranslatedText("&Cancel")
276     lblNewvHostTitle.Caption = WinUI.GetTranslatedText("Add a new Virtual Host:")
280     lblNewvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
284     lblNewvHostDomain.Caption = WinUI.GetTranslatedText("What is the domain for this Virtual Host?")
288     lblNewvHostRoot.Caption = WinUI.GetTranslatedText("Where is the root folder for this Virtual Host?")
292     lblNewvHostLogs.Caption = WinUI.GetTranslatedText("Where do you want to keep the log for this Virtual Host?")
296     cmdNewvHostOK.Caption = WinUI.GetTranslatedText("&OK")
300     cmdNewvHostCancel.Caption = WinUI.GetTranslatedText("&Cancel")
304     lblDynDNSTitle.Caption = WinUI.GetTranslatedText("From here you can enable updates && maintance of you DynDNS.org account. To use this feature you must have a acount and setup a Dynamic DNS host. You can not add a new host via the system.")
308     lblDynDNSCurrentIP.Caption = WinUI.GetTranslatedText("Current IP:")
312     lblDynDNSLastUpdate.Caption = WinUI.GetTranslatedText("Last Update:")
316     lblDynDNSLastResult.Caption = WinUI.GetTranslatedText("Last Update Result:")
320     lblDynDNSHostname.Caption = WinUI.GetTranslatedText("DynDNS.org Hostname:")
324     lblDynDNSUsername.Caption = WinUI.GetTranslatedText("DynDNS.org Username:")
328     lblDynDNSPassword.Caption = WinUI.GetTranslatedText("DynDNS.org Password:")
332     cmdDynDNSUpdate.Caption = WinUI.GetTranslatedText("&Update")
336     chkDynDNSEnable.Caption = WinUI.GetTranslatedText("Enable DynDNS Updates?")
340     lblConfigAdvIPBind.Caption = WinUI.GetTranslatedText("What IP should the server listen to? (Default: Leave blank for all available)")
344     lblConfigBasicErrorLog.Caption = WinUI.GetTranslatedText("Where do you want to store the server error log?")
    
348     WinUI.Status "Loading Configuration Data..."
352     If LoadConfigData = False Then
356         RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
360         Select Case RetVal
                Case vbAbort
364                 End
368             Case vbRetry
372                 If LoadConfigData = False Then
376                     MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
380                     End
                    End If
384             Case vbIgnore
388                 MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    
392     WinUI.Status "Finalizing..."
396     With vbaSideBar
400         .Redraw = False
404         Set cBar = .Bars.Add(, "status", WinUI.GetTranslatedText("System Status"))
408         Set cItem = cBar.Items.Add(, "status", WinUI.GetTranslatedText("Current Status"), 0)
        
412         Set cBar = .Bars.Add(, "config", WinUI.GetTranslatedText("Configuration"))
416         Set cItem = cBar.Items.Add(, "basic", WinUI.GetTranslatedText("Basic"), 0)
420         Set cItem = cBar.Items.Add(, "advanced", WinUI.GetTranslatedText("Advanced"), 0)
424         Set cItem = cBar.Items.Add(, "vhost", WinUI.GetTranslatedText("Virtual Host"), 0)
428         Set cItem = cBar.Items.Add(, "cgi", WinUI.GetTranslatedText("CGI"), 0)
432         Set cItem = cBar.Items.Add(, "dyndns", WinUI.GetTranslatedText("Dynamic DNS"), 0)
        
436         Set cBar = .Bars.Add(, "logs", WinUI.GetTranslatedText("System Logs"))
440         Set cItem = cBar.Items.Add(, "logs", WinUI.GetTranslatedText("View Logs"), 0)
444         .Height = Me.Height
448         .Redraw = True
        End With

452     fraStatus.ZOrder 0
456     vbaSideBar.ZOrder 0
460     tmrStatus_Timer
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '<EhHeader>
    On Error GoTo Form_QueryUnload_Err
    '</EhHeader>
    Dim lngRetVal As Long
100     If blnDirty = True Then
104         lngRetVal = MsgBox(WinUI.GetTranslatedText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
108         If lngRetVal = vbYes Then
112             If WinUI.Config.Save(WinUI.ConfigFile) = False Then
116                 MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
                End If
            End If
        End If
120     Me.Visible = False
124     DoEvents
128     WinUI.UnloadApp
    '<EhFooter>
    Exit Sub

Form_QueryUnload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_QueryUnload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblUpdateStatus_Click()
    '<EhHeader>
    On Error GoTo lblUpdateStatus_Click_Err
    '</EhHeader>
100     If WinUI.Update.IsAvailable = True Then
104         Load frmUpdate
108         frmUpdate.Show
        End If
    '<EhFooter>
    Exit Sub

lblUpdateStatus_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblUpdateStatus_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstCGI_Click()
    '<EhHeader>
    On Error GoTo lstCGI_Click_Err
    '</EhHeader>
100     cmdBrowseCGIInterp.Enabled = True
104     cmdCGIRemove.Enabled = True
108     txtCGIInterp.Enabled = True
112     txtCGIExt.Enabled = True
116     txtCGIInterp.Text = WinUI.Config.CGI.Item(lstCGI.Text).Interpreter
120     txtCGIExt.Text = WinUI.Config.CGI.Item(lstCGI.Text).Extention
    '<EhFooter>
    Exit Sub

lstCGI_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lstCGI_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstvHosts_Click()
    '<EhHeader>
    On Error GoTo lstvHosts_Click_Err
    '</EhHeader>
100     cmdBrowsevHostRoot.Enabled = True
104     cmdBrowsevHostLog.Enabled = True
108     cmdvHostRemove.Enabled = True
112     txtvHostName.Enabled = True
116     txtvHostDomain.Enabled = True
120     txtvHostRoot.Enabled = True
124     txtvHostLog.Enabled = True
128     txtvHostName.Text = WinUI.Config.vHost.Item(lstvHosts.Text).HostName
132     txtvHostDomain.Text = WinUI.Config.vHost.Item(lstvHosts.Text).Domain
136     txtvHostRoot.Text = WinUI.Config.vHost.Item(lstvHosts.Text).Root
140     txtvHostLog.Text = WinUI.Config.vHost.Item(lstvHosts.Text).Log
    '<EhFooter>
    Exit Sub

lstvHosts_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lstvHosts_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileExit_Click()
    '<EhHeader>
    On Error GoTo mnuFileExit_Click_Err
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    Exit Sub

mnuFileExit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileExit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileExport_Click()
    '<EhHeader>
    On Error GoTo mnuFileExport_Click_Err
    '</EhHeader>
    Dim cDlg As cCommonDialog
    Dim strFile As String

100     Set cDlg = New cCommonDialog
104     If cDlg.VBGetSaveFileName(strFile, , , "Text Files (*.txt)|*.txt|All Files (*.*)|*.*") Then
108         Open strFile For Append As 1
112             Print #1, WinUI.Config.Report
116         Close 1
        End If
120     Set cDlg = Nothing
    '<EhFooter>
    Exit Sub

mnuFileExport_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileExport_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileReload_Click()
    '<EhHeader>
    On Error GoTo mnuFileReload_Click_Err
    '</EhHeader>
    Dim RetVal As Long
100     RetVal = MsgBox(WinUI.GetTranslatedText("This will reset any changes you make.\r\rDo you want to continue?"), vbYesNo + vbQuestion)
104     If RetVal = vbYes Then
108         If LoadConfigData = False Then
112             RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
116             Select Case RetVal
                    Case vbAbort
120                     Unload Me
124                 Case vbRetry
128                     If LoadConfigData = False Then
132                         MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                        End If
136                 Case vbIgnore
140                     MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
                End Select
            End If
        End If
    '<EhFooter>
    Exit Sub

mnuFileReload_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileReload_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileSave_Click()
    '<EhHeader>
    On Error GoTo mnuFileSave_Click_Err
    '</EhHeader>
100     If WinUI.Config.Save(WinUI.ConfigFile) = False Then
104         MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox WinUI.GetTranslatedText("You data has been saved./r/rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
    '<EhFooter>
    Exit Sub

mnuFileSave_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileSave_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpAbout_Click()
    '<EhHeader>
    On Error GoTo mnuHelpAbout_Click_Err
    '</EhHeader>
100     Load frmAbout
104     frmAbout.Show vbModal
    '<EhFooter>
    Exit Sub

mnuHelpAbout_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpAbout_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpEventViewer_Click()
    '<EhHeader>
    On Error GoTo mnuHelpEventViewer_Click_Err
    '</EhHeader>
100     Load frmEventView
104     frmEventView.Show
    '<EhFooter>
    Exit Sub

mnuHelpEventViewer_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpEventViewer_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpForum_Click()
    '<EhHeader>
    On Error GoTo mnuHelpForum_Click_Err
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
    '<EhFooter>
    Exit Sub

mnuHelpForum_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpForum_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpHomePage_Click()
    '<EhHeader>
    On Error GoTo mnuHelpHomePage_Click_Err
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
    '<EhFooter>
    Exit Sub

mnuHelpHomePage_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpHomePage_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpRegister_Click()
    '<EhHeader>
    On Error GoTo mnuHelpRegister_Click_Err
    '</EhHeader>
100     WinUI.Registration.Start
    '<EhFooter>
    Exit Sub

mnuHelpRegister_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpRegister_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpUpdate_Click()
    '<EhHeader>
    On Error GoTo mnuHelpUpdate_Click_Err
    '</EhHeader>
100     WinUI.Status WinUI.GetTranslatedText("Retrieving Update Information") & "...", True
104     WinUI.Update.Check
108     If WinUI.Update.IsAvailable = True Then
112         lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
116         lblUpdateStatus.Font.Underline = True
120         lblUpdateStatus.ForeColor = vbBlue
124         lblUpdateStatus.MousePointer = vbCustom
128         Load frmUpdate
132         frmUpdate.Show
        Else
136         MsgBox WinUI.GetTranslatedText("You have the most current version available."), vbOKOnly + vbInformation
        End If
140     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Sub

mnuHelpUpdate_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpUpdate_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub netDynDNS_StateChanged(ByVal State As Integer)
    '<EhHeader>
    On Error GoTo netDynDNS_StateChanged_Err
    '</EhHeader>
    Dim strResult As String

100     Select Case State
            Case icHostResolved
104             DoEvents
108             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icHostResolved"
112         Case icConnecting
116             DoEvents
120             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icConnecting"
124         Case icConnected
128             DoEvents
132             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icConnected"
136         Case icRequesting
140             DoEvents
144             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icRequesting"
148         Case icRequestSent
152             DoEvents
156             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icRequestSent"
160         Case icReceivingResponse
164             DoEvents
168             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icReceivingResponse"
172         Case icResponseReceived
176             DoEvents
180             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icResponseReceived"
184         Case icDisconnecting
188             DoEvents
192             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icDisconnecting"
196         Case icDisconnected
200             DoEvents
204             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icDisconnected"
208         Case icError
212             DoEvents
216             WinUI.EventLog.AddEvent "WinUI.frmMain.netDynDNS_StateChanged", "icError: Code: " & netDynDNS.ResponseCode & " Info: " & netDynDNS.ResponseInfo
220         Case icResponseCompleted
224             strResult = netDynDNS.GetChunk(1024, icString)
228             WinUI.DynDNS.LastIP = WinUI.Net.CurrentIP
232             WinUI.DynDNS.LastUpdate = Now
236             WinUI.DynDNS.LastResult = strResult
240             txtDynDNSLastUpdate.Text = WinUI.DynDNS.LastUpdate
244             txtDynDNSLastResult.Text = WinUI.DynDNS.LastResult
            
248             WinUI.DynDNS.Save

252             cmdDynDNSUpdate.Enabled = False
256             WinUI.Status "Ready..."
260             MsgBox "Update completed. DynDNS.org returned:" & vbCrLf & vbCrLf & Chr$(9) & strResult, vbInformation 'this line will go away soon, thus no GT
        End Select
    '<EhFooter>
    Exit Sub

netDynDNS_StateChanged_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.netDynDNS_StateChanged", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub tmrStats_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UpdateStats
End Sub

Private Sub tmrStatus_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
Dim strSrvStatusCur As String
    strSrvStatusCur = ServiceStatus("", "SWEBS Web Server")
    lblSrvStatusCur.Font.Bold = False
    Select Case strSrvStatusCur
        Case "Stopped"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Stopped")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Stopped"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Start Pending"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Start Pending")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Start Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Stop Pending"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Stop Pending")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Stop Pending"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Running")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
        Case "Continue Pending"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Continue Pending")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Continue Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Pause Pending"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Pause Pending")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status:  Pending"
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Paused"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Paused")
            WinUI.EventLog.AddEvent "WinUI.frmMain.tmrStatus_Timer", "Service Status: Paused"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
    End Select
End Sub


Private Function LoadConfigData() As Boolean
    '<EhHeader>
    On Error GoTo LoadConfigData_Err
    '</EhHeader>
    Dim i As Long
    Dim strTemp As String
    Dim strResult As String
    Dim vItem As Variant
    
100     WinUI.EventLog.AddEvent "WinUI.frmMain.LoadConfigData", "Loading Config Data"
104     WinUI.Status WinUI.GetTranslatedText("Loading Configuration Data") & "...", True
112     LoadConfigData = WinUI.Config.LoadData
    
        'Setup the form...
116     txtServerName.Text = WinUI.Config.ServerName
120     txtPort.Text = WinUI.Config.Port
124     txtWebroot.Text = WinUI.Config.WebRoot
128     txtMaxConnect.Text = WinUI.Config.MaxConnections
132     txtLogFile.Text = WinUI.Config.LogFile
136     txtConfigAdvIPBind.Text = WinUI.Config.ListeningAddress
140     txtAllowIndex.Text = WinUI.Config.AllowIndex
144     txtErrorPages.Text = WinUI.Config.ErrorPages
148     txtConfigBasicErrorLog.Text = WinUI.Config.ErrorLog
    
152     For Each vItem In WinUI.Config.Index
156         strTemp = strTemp & vItem.FileName & " "
        Next
160     txtIndexFiles.Text = Trim$(strTemp)
    
164     lstCGI.Enabled = False
168     lstCGI.Clear
172     For Each vItem In WinUI.Config.CGI
176         lstCGI.AddItem vItem.Extention
180         lstCGI.Enabled = True
        Next
    
184     lstvHosts.Enabled = False
188     lstvHosts.Clear
192     For Each vItem In WinUI.Config.vHost
196         lstvHosts.AddItem vItem.HostName
200         lstvHosts.Enabled = True
        Next
    
204     cmbViewLogFiles.Clear
208     If Dir$(WinUI.Config.LogFile) <> "" Then
212         cmbViewLogFiles.AddItem WinUI.Config.LogFile
        End If
216     If Dir$(WinUI.Config.ErrorLog) <> "" Then
220         cmbViewLogFiles.AddItem WinUI.Config.ErrorLog
        End If
224     For Each vItem In WinUI.Config.vHost
228         If Dir$(vItem.Log) <> "" Then
232             cmbViewLogFiles.AddItem vItem.Log
            End If
        Next
    
        'we now only check for updates every 24 hours, this could confuse some people.
        'but this should make loading faster.
236     WinUI.Status "Checking For Updates...", True
240     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "LastUpdateCheck")
244     If strResult = "" Then
248         strResult = CDate(1.1)
        End If
252     If DateDiff("h", CDate(strResult), Now) >= 24 Then
256         WinUI.Update.Check
260         If WinUI.Update.IsAvailable = True Then
264             lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
            Else
268             lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
272             lblUpdateStatus.Font.Underline = False
276             lblUpdateStatus.ForeColor = vbButtonText
280             lblUpdateStatus.MousePointer = vbDefault
284             SaveRegistryString &H80000002, "SOFTWARE\SWS", "LastUpdateCheck", Now
            End If
        Else
288         lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
292         lblUpdateStatus.Font.Underline = False
296         lblUpdateStatus.ForeColor = vbButtonText
300         lblUpdateStatus.MousePointer = vbDefault
        End If
    
304     UpdateStats
    
308     WinUI.Status "Getting DNS Data...", True
312     txtDynDNSCurrentIP.Text = WinUI.Net.CurrentIP
316     txtDynDNSHostname.Text = WinUI.DynDNS.HostName
320     txtDynDNSUsername.Text = WinUI.DynDNS.UserName
324     txtDynDNSLastUpdate.Text = WinUI.DynDNS.LastUpdate
328     txtDynDNSLastUpdate.Enabled = False
332     txtDynDNSLastResult.Text = WinUI.DynDNS.LastResult
336     txtDynDNSLastResult.Enabled = False
340     txtDynDNSPassword.Text = WinUI.DynDNS.Password
344     If WinUI.DynDNS.Enabled = True Then
348         chkDynDNSEnable.Value = vbChecked
        End If
352     If WinUI.Net.CurrentIP <> WinUI.DynDNS.LastIP Or DateDiff("d", CDate(WinUI.DynDNS.LastUpdate), Now) >= 28 Then
356         cmdDynDNSUpdate.Enabled = True
        Else
360         cmdDynDNSUpdate.Enabled = False
        End If
    
364     If WinUI.Registration.IsRegistered = True Then
368         WinUI.Status "Updating Registration..."
372         mnuHelpRegister.Enabled = False
376         WinUI.Registration.Renew
        End If
    
380     WinUI.Status "Ready..."
    '<EhFooter>
    Exit Function

LoadConfigData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.LoadConfigData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Private Sub txtAllowIndex_Change()
    '<EhHeader>
    On Error GoTo txtAllowIndex_Change_Err
    '</EhHeader>
100     If WinUI.Config.AllowIndex <> IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false") Then
104         WinUI.Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtAllowIndex_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtAllowIndex_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtCGIExt_Change()
    '<EhHeader>
    On Error GoTo txtCGIExt_Change_Err
    '</EhHeader>
100     If lstCGI.ListIndex <> -1 Then
104         If WinUI.Config.CGI.Item(lstCGI.Text).Extention <> txtCGIExt.Text Then
108             WinUI.Config.CGI.Item(lstCGI.Text).Extention = txtCGIExt.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    Exit Sub

txtCGIExt_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtCGIExt_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtCGIInterp_Change()
    '<EhHeader>
    On Error GoTo txtCGIInterp_Change_Err
    '</EhHeader>
100     If lstCGI.ListIndex <> -1 Then
104         If WinUI.Config.CGI.Item(lstCGI.Text).Interpreter <> txtCGIInterp.Text Then
108             WinUI.Config.CGI.Item(lstCGI.Text).Interpreter = txtCGIInterp.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    Exit Sub

txtCGIInterp_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtCGIInterp_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtConfigAdvIPBind_Change()
    '<EhHeader>
    On Error GoTo txtConfigAdvIPBind_Change_Err
    '</EhHeader>
100     If WinUI.Config.ListeningAddress = txtConfigAdvIPBind.Text Then
104         WinUI.Config.ListeningAddress = txtConfigAdvIPBind.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtConfigAdvIPBind_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtConfigAdvIPBind_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtConfigBasicErrorLog_Change()
    '<EhHeader>
    On Error GoTo txtConfigBasicErrorLog_Change_Err
    '</EhHeader>
100     If WinUI.Config.ErrorLog <> txtConfigBasicErrorLog.Text Then
104         WinUI.Config.ErrorLog = txtConfigBasicErrorLog.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtConfigBasicErrorLog_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtConfigBasicErrorLog_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSCurrentIP_Change()
    '<EhHeader>
    On Error GoTo txtDynDNSCurrentIP_Change_Err
    '</EhHeader>
100     If WinUI.Net.CurrentIP <> WinUI.DynDNS.LastIP Or DateDiff("d", CDate(WinUI.DynDNS.LastUpdate), Now) >= 28 Then
104         cmdDynDNSUpdate.Enabled = True
        Else
108         cmdDynDNSUpdate.Enabled = False
        End If
    '<EhFooter>
    Exit Sub

txtDynDNSCurrentIP_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSCurrentIP_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSCurrentIP_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo txtDynDNSCurrentIP_KeyPress_Err
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    Exit Sub

txtDynDNSCurrentIP_KeyPress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSCurrentIP_KeyPress", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSCurrentIP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '<EhHeader>
    On Error GoTo txtDynDNSCurrentIP_MouseUp_Err
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    Exit Sub

txtDynDNSCurrentIP_MouseUp_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSCurrentIP_MouseUp", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSHostname_Change()
    '<EhHeader>
    On Error GoTo txtDynDNSHostname_Change_Err
    '</EhHeader>
100     If WinUI.DynDNS.HostName <> txtDynDNSHostname.Text Then
104         WinUI.DynDNS.HostName = txtDynDNSHostname.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtDynDNSHostname_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSHostname_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSPassword_Change()
    '<EhHeader>
    On Error GoTo txtDynDNSPassword_Change_Err
    '</EhHeader>
100     If WinUI.DynDNS.Password <> txtDynDNSPassword.Text Then
104         WinUI.DynDNS.Password = txtDynDNSPassword.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtDynDNSPassword_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSPassword_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtDynDNSUsername_Change()
    '<EhHeader>
    On Error GoTo txtDynDNSUsername_Change_Err
    '</EhHeader>
100     If WinUI.DynDNS.UserName <> txtDynDNSUsername.Text Then
104         WinUI.DynDNS.UserName = txtDynDNSUsername.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtDynDNSUsername_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtDynDNSUsername_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtErrorPages_Change()
    '<EhHeader>
    On Error GoTo txtErrorPages_Change_Err
    '</EhHeader>
100     If WinUI.Config.ErrorPages <> txtErrorPages.Text Then
104         WinUI.Config.ErrorPages = txtErrorPages.Text
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtErrorPages_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtErrorPages_Change", Erl, False
    Resume Next
    '</EhFooter>
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
'<EhHeader>
On Error GoTo txtIndexFiles_Change_Err
'</EhHeader>
'<EhFooter>
Exit Sub

txtIndexFiles_Change_Err:
DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_Change", Erl, False
Resume Next
'</EhFooter>
End Sub

Private Sub txtIndexFiles_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo txtIndexFiles_KeyPress_Err
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    Exit Sub

txtIndexFiles_KeyPress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_KeyPress", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtIndexFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '<EhHeader>
    On Error GoTo txtIndexFiles_MouseUp_Err
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    Exit Sub

txtIndexFiles_MouseUp_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_MouseUp", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtLogFile_Change()
    '<EhHeader>
    On Error GoTo txtLogFile_Change_Err
    '</EhHeader>
100     If WinUI.Config.LogFile <> Trim$(txtLogFile.Text) Then
104         WinUI.Config.LogFile = Trim$(txtLogFile.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtLogFile_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtLogFile_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtMaxConnect_Change()
    '<EhHeader>
    On Error GoTo txtMaxConnect_Change_Err
    '</EhHeader>
100     If WinUI.Config.MaxConnections <> Int(Val(txtMaxConnect.Text)) Then
104         WinUI.Config.MaxConnections = Int(Val(txtMaxConnect.Text))
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtMaxConnect_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtMaxConnect_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtPort_Change()
    '<EhHeader>
    On Error GoTo txtPort_Change_Err
    '</EhHeader>
100     If WinUI.Config.Port <> Int(Val(txtPort.Text)) Then
104         WinUI.Config.Port = Int(Val(txtPort.Text))
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtPort_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtPort_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtServerName_Change()
    '<EhHeader>
    On Error GoTo txtServerName_Change_Err
    '</EhHeader>
100     If WinUI.Config.ServerName <> Trim$(txtServerName.Text) Then
104         WinUI.Config.ServerName = Trim$(txtServerName.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtServerName_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtServerName_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostDomain_Change()
    '<EhHeader>
    On Error GoTo txtvHostDomain_Change_Err
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Config.vHost.Item(lstvHosts.Text).Domain <> txtvHostDomain.Text Then
108             WinUI.Config.vHost.Item(lstvHosts.Text).Domain = txtvHostDomain.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    Exit Sub

txtvHostDomain_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostDomain_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostLog_Change()
    '<EhHeader>
    On Error GoTo txtvHostLog_Change_Err
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Config.vHost.Item(lstvHosts.Text).Log <> txtvHostLog.Text Then
108             WinUI.Config.vHost.Item(lstvHosts.Text).Log = txtvHostLog.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    Exit Sub

txtvHostLog_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostLog_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostName_Change()
    '<EhHeader>
    On Error GoTo txtvHostName_Change_Err
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Config.vHost.Item(lstvHosts.Text).HostName <> txtvHostName.Text Then
108             blnDirty = True
112             WinUI.Config.vHost.Item(lstvHosts.Text).HostName = txtvHostName.Text
            End If
        End If
    '<EhFooter>
    Exit Sub

txtvHostName_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostName_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostRoot_Change()
    '<EhHeader>
    On Error GoTo txtvHostRoot_Change_Err
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Config.vHost.Item(lstvHosts.Text).Root <> txtvHostRoot.Text Then
108             WinUI.Config.vHost.Item(lstvHosts.Text).Root = txtvHostRoot.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    Exit Sub

txtvHostRoot_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostRoot_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtWebroot_Change()
    '<EhHeader>
    On Error GoTo txtWebroot_Change_Err
    '</EhHeader>
100     If WinUI.Config.WebRoot <> Trim$(txtWebroot.Text) Then
104         WinUI.Config.WebRoot = Trim$(txtWebroot.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    Exit Sub

txtWebroot_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtWebroot_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub vbaSideBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    '<EhHeader>
    On Error GoTo vbaSideBar_ItemClick_Err
    '</EhHeader>
100     StopWinUpdate Me.hWnd
104     Select Case itm.Key
            Case "status"
108             fraStatus.ZOrder 0
112         Case "basic"
116             fraConfigBasic.ZOrder 0
120         Case "advanced"
124             fraConfigAdv.ZOrder 0
128         Case "vhost"
132             fraConfigvHost.ZOrder 0
136         Case "cgi"
140             fraConfigCGI.ZOrder 0
144         Case "dyndns"
148             fraConfigDynDns.ZOrder 0
152         Case "logs"
156             fraLogs.ZOrder 0
        End Select
160     vbaSideBar.ZOrder 0
164     StopWinUpdate
    '<EhFooter>
    Exit Sub

vbaSideBar_ItemClick_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.vbaSideBar_ItemClick", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub UpdateStats()
    '<EhHeader>
    On Error GoTo UpdateStats_Err
    '</EhHeader>
100     WinUI.Stats.Reload
104     lblStatsLastRestart.Caption = WinUI.GetTranslatedText("Last Restart") & ": " & WinUI.Stats.LastRestart
108     lblStatsRequestCount.Caption = WinUI.GetTranslatedText("Request Count") & ": " & WinUI.Stats.RequestCount
112     lblStatsBytesSent.Caption = WinUI.GetTranslatedText("Total Bytes Sent") & ": " & Format$(WinUI.Stats.TotalBytesSent, "###,###,###,###,##0")
116     lblCurVersion.Caption = WinUI.GetTranslatedText("Current Version") & ": " & WinUI.Version
120     lblUpdateVersion.Caption = WinUI.GetTranslatedText("Update Version") & ": " & IIf(WinUI.Update.Version <> "", WinUI.Update.Version, WinUI.Version)
    '<EhFooter>
    Exit Sub

UpdateStats_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.UpdateStats", Erl, False
    Resume Next
    '</EhFooter>
End Sub
