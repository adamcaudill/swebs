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
         Caption         =   $"frmMain.frx":0CCE
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
         Caption         =   $"frmMain.frx":0D72
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
         ItemData        =   "frmMain.frx":0E20
         Left            =   120
         List            =   "frmMain.frx":0E22
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Timer tmrStats 
      Interval        =   120
      Left            =   5520
      Top             =   3960
   End
   Begin InetCtlsObjects.Inet netDynDNS 
      Left            =   5040
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
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
         Caption         =   $"frmMain.frx":0E24
         Height          =   735
         Left            =   240
         TabIndex        =   104
         Top             =   240
         Width           =   6495
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
      Interval        =   750
      Left            =   4920
      Top             =   3840
   End
   Begin InetCtlsObjects.Inet netMain 
      Left            =   5280
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
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
   Begin VB.Label lblAppStatus 
      Caption         =   "Current App Status..."
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
108         DynDNS.Enabled = True
        Else
112         DynDNS.Enabled = False
        End If
        '<EhFooter>
        Exit Sub

chkDynDNSEnable_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.chkDynDNSEnable_Click", Erl, False
120     Resume Next
        '</EhFooter>
End Sub

Private Sub cmbViewLogFiles_Click()
        '<EhHeader>
        On Error GoTo cmbViewLogFiles_Click_Err
        '</EhHeader>
    Dim strLog As String
    Dim lngLen As Long
    
100     AppStatus True, GetText("Loading Log File") & "..."
104     If Dir$(cmbViewLogFiles.Text) <> "" Then
108         lngLen = FileLen(cmbViewLogFiles.Text)
112         strLog = Space$(lngLen)
116         Open cmbViewLogFiles.Text For Binary As 1 Len = lngLen
120             Get #1, 1, strLog
124         Close 1
128         txtViewLogFiles.Text = strLog
132         txtViewLogFiles.SetFocus
        Else
136         DoEvents
140         MsgBox GetText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
        End If
144     AppStatus False
        '<EhFooter>
        Exit Sub

cmbViewLogFiles_Click_Err:
148     DisplayErrMsg Err.Description, "WinUI.frmMain.cmbViewLogFiles_Click", Erl, False
152     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdApply_Click()
        '<EhHeader>
        On Error GoTo cmdApply_Click_Err
        '</EhHeader>
100     If SaveConfigData(strConfigFile) = False Then
104         MsgBox GetText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox GetText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
        '<EhFooter>
        Exit Sub

cmdApply_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdApply_Click", Erl, False
120     Resume Next
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
104     strStartDir = Mid$(Config.CGI((lstCGI.ListIndex + 1), 1), 1, (Len(Config.CGI((lstCGI.ListIndex + 1), 1)) - InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\")))
108     If cDlg.VBGetOpenFileName(strFile, , True, , , , "Executable Files (*.exe)|*.exe", , strStartDir, , "exe") Then
112         txtCGIInterp.Text = strFile
        End If
116     Set cDlg = Nothing
        '<EhFooter>
        Exit Sub

cmdBrowseCGIInterp_Click_Err:
120     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseCGIInterp_Click", Erl, False
124     Resume Next
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
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseErrorLog_Click", Erl, False
120     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBrowseErrorPages_Click()
        '<EhHeader>
        On Error GoTo cmdBrowseErrorPages_Click_Err
        '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = BrowseForFolder(Me, , True, Config.ErrorPages)
108     If strPath <> "" Then
112         txtErrorPages.Text = strPath
        End If
        '<EhFooter>
        Exit Sub

cmdBrowseErrorPages_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseErrorPages_Click", Erl, False
120     Resume Next
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
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseNewCGIInterp_Click", Erl, False
120     Resume Next
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
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseNewvHostLogs_Click", Erl, False
120     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
        '<EhHeader>
        On Error GoTo cmdBrowseNewvHostRoot_Click_Err
        '</EhHeader>
    Dim strPath As String
100     strPath = BrowseForFolder(Me, , True, Config.WebRoot)
104     If strPath <> "" Then
108         txtNewvHostRoot.Text = strPath
        End If
        '<EhFooter>
        Exit Sub

cmdBrowseNewvHostRoot_Click_Err:
112     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseNewvHostRoot_Click", Erl, False
116     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBrowseRoot_Click()
        '<EhHeader>
        On Error GoTo cmdBrowseRoot_Click_Err
        '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = BrowseForFolder(Me, , True, Config.WebRoot)
108     If strPath <> "" Then
112         txtWebroot.Text = strPath
        End If
        '<EhFooter>
        Exit Sub

cmdBrowseRoot_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseRoot_Click", Erl, False
120     Resume Next
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
108     strStartDir = Mid$(Config.vHost((lstvHosts.ListIndex + 1)).Log, (InStrRev(Config.vHost((lstvHosts.ListIndex + 1)).Log, "\") + 1))
112     If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
116         txtvHostLog.Text = strFile
        End If
120     Set cDlg = Nothing
        '<EhFooter>
        Exit Sub

cmdBrowsevHostLog_Click_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowsevHostLog_Click", Erl, False
128     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBrowsevHostRoot_Click()
        '<EhHeader>
        On Error GoTo cmdBrowsevHostRoot_Click_Err
        '</EhHeader>
    Dim strPath As String
100     strPath = BrowseForFolder(Me, , True, Config.vHost((lstvHosts.ListIndex + 1)).Root)
104     If strPath <> "" Then
108         txtvHostRoot.Text = strPath
        End If
        '<EhFooter>
        Exit Sub

cmdBrowsevHostRoot_Click_Err:
112     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowsevHostRoot_Click", Erl, False
116     Resume Next
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
108     strStartDir = Mid$(Config.LogFile, (InStrRev(Config.LogFile, "\") + 1))
112     If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
116         txtLogFile.Text = strFile
        End If
120     Set cDlg = Nothing
        '<EhFooter>
        Exit Sub

cmdBrowseLogFile_Click_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdBrowseLogFile_Click", Erl, False
128     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdCancel_Click", Erl, False
108     Resume Next
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
108     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdCGINew_Click", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCGIRemove_Click()
        '<EhHeader>
        On Error GoTo cmdCGIRemove_Click_Err
        '</EhHeader>
    Dim lngRetVal As Long
    Dim i As Long

100     If lstCGI.ListIndex >= 0 Then
104         lngRetVal = MsgBox(GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
108         If lngRetVal = vbYes Then
112             blnDirty = True
116             RemoveCGI (lstCGI.ListIndex + 1)
120             lstCGI.Clear
124             If Config.CGI(1, 2) <> "" Then
128                 For i = 1 To UBound(Config.CGI)
132                     lstCGI.AddItem Config.CGI(i, 2)
                    Next
                Else
136                 lstCGI.Enabled = False
140                 cmdBrowseCGIInterp.Enabled = False
144                 cmdCGIRemove.Enabled = False
148                 txtCGIInterp.Enabled = False
152                 txtCGIExt.Enabled = False
156                 txtCGIInterp.Text = ""
160                 txtCGIExt.Text = ""
                End If
            End If
        End If
        '<EhFooter>
        Exit Sub

cmdCGIRemove_Click_Err:
164     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdCGIRemove_Click", Erl, False
168     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDynDNSUpdate_Click()
        '<EhHeader>
        On Error GoTo cmdDynDNSUpdate_Click_Err
        '</EhHeader>
100     AppStatus True, "Updating DNS Information..."
104     netDynDNS.URL = "http://members.dyndns.org"
108     netDynDNS.Document = "/nic/update?system=dyndns&hostname=" & DynDNS.Hostname & "&myip=" & DynDNS.CurrentIP & "&wildcard=NOCHG"
112     netDynDNS.UserName = DynDNS.UserName
116     netDynDNS.Password = DynDNS.Password
120     netDynDNS.Execute , "GET", , "User-Agent: SWEBS WinUI " & strInstalledVer & " <plenojure@users.sf.net>"
        '<EhFooter>
        Exit Sub

cmdDynDNSUpdate_Click_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdDynDNSUpdate_Click", Erl, False
128     Resume Next
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
112     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdNewCGICancel_Click", Erl, False
116     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNewCGIOK_Click()
        '<EhHeader>
        On Error GoTo cmdNewCGIOK_Click_Err
        '</EhHeader>
    Dim i As Long

100     If txtNewCGIInterp.Text <> "" And txtNewCGIExt.Text <> "" Then
104         blnDirty = True
108         AddNewCGI txtNewCGIExt.Text, txtNewCGIInterp.Text
112         If Config.CGI(1, 2) <> "" Then
116             lstCGI.Clear
120             For i = 1 To UBound(Config.CGI)
124                 lstCGI.AddItem Config.CGI(i, 2)
                Next
            Else
128             lstCGI.Enabled = False
            End If
132         fraNewCGI.ZOrder 1
136         txtNewCGIInterp.Text = ""
140         txtNewCGIExt.Text = ""
        Else
144         MsgBox GetText("Please fill all fields.")
        End If
        '<EhFooter>
        Exit Sub

cmdNewCGIOK_Click_Err:
148     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdNewCGIOK_Click", Erl, False
152     Resume Next
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
120     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdNewvHostCancel_Click", Erl, False
124     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNewvHostOK_Click()
        '<EhHeader>
        On Error GoTo cmdNewvHostOK_Click_Err
        '</EhHeader>
    Dim i As Long

100     If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
104         blnDirty = True
108         AddNewvHost txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text
112         lstvHosts.Clear
116         If Config.vHost(1).Name <> "" Then
120             For i = 1 To UBound(Config.vHost)
124                 lstvHosts.AddItem Config.vHost(i).Name
                Next
128             lstvHosts.Enabled = True
            Else
132             lstvHosts.Enabled = False
            End If
136         fraNewvHost.ZOrder 1
140         txtNewvHostName.Text = ""
144         txtNewvHostDomain.Text = ""
148         txtNewvHostRoot.Text = ""
152         txtNewvHostLogs.Text = ""
        Else
156         MsgBox GetText("Please fill all fields.")
        End If
        '<EhFooter>
        Exit Sub

cmdNewvHostOK_Click_Err:
160     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdNewvHostOK_Click", Erl, False
164     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdOK_Click", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSrvRestart_Click()
        '<EhHeader>
        On Error GoTo cmdSrvRestart_Click_Err
        '</EhHeader>
100     AppStatus True, GetText("Restarting Service") & "..."
104     ServiceStop "", "SWEBS Web Server"
108     Do Until ServiceStatus("", "SWEBS Web Server") = "Stopped"
112         DoEvents
        Loop
116     ServiceStart "", "SWEBS Web Server"
120     UpdateStats
124     AppStatus False
        '<EhFooter>
        Exit Sub

cmdSrvRestart_Click_Err:
128     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdSrvRestart_Click", Erl, False
132     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSrvStart_Click()
        '<EhHeader>
        On Error GoTo cmdSrvStart_Click_Err
        '</EhHeader>
100     AppStatus True, GetText("Starting Service") & "..."
104     ServiceStart "", "SWEBS Web Server"
108     UpdateStats
112     AppStatus False
        '<EhFooter>
        Exit Sub

cmdSrvStart_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdSrvStart_Click", Erl, False
120     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSrvStop_Click()
        '<EhHeader>
        On Error GoTo cmdSrvStop_Click_Err
        '</EhHeader>
100     AppStatus True, GetText("Stopping Service") & "..."
104     ServiceStop "", "SWEBS Web Server"
108     AppStatus False
        '<EhFooter>
        Exit Sub

cmdSrvStop_Click_Err:
112     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdSrvStop_Click", Erl, False
116     Resume Next
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
108     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdvHostNew_Click", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub cmdvHostRemove_Click()
        '<EhHeader>
        On Error GoTo cmdvHostRemove_Click_Err
        '</EhHeader>
    Dim lngRetVal As Long
    Dim i As Long

100     If lstvHosts.ListIndex >= 0 Then
104         lngRetVal = MsgBox(GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
108         If lngRetVal = vbYes Then
112             blnDirty = True
116             RemovevHost (lstvHosts.ListIndex + 1)
120             lstvHosts.Clear
124             If Config.vHost(1).Name <> "" Then
128                 For i = 1 To UBound(Config.vHost)
132                     lstvHosts.AddItem Config.vHost(i).Name
                    Next
                Else
136                 cmdBrowsevHostRoot.Enabled = False
140                 cmdBrowsevHostLog.Enabled = False
144                 cmdvHostRemove.Enabled = False
148                 txtvHostName.Enabled = False
152                 txtvHostDomain.Enabled = False
156                 txtvHostRoot.Enabled = False
160                 txtvHostLog.Enabled = False
164                 txtvHostName.Text = ""
168                 txtvHostDomain.Text = ""
172                 txtvHostRoot.Text = ""
176                 txtvHostLog.Text = ""
                End If
            End If
        End If
        '<EhFooter>
        Exit Sub

cmdvHostRemove_Click_Err:
180     DisplayErrMsg Err.Description, "WinUI.frmMain.cmdvHostRemove_Click", Erl, False
184     Resume Next
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
100     mnuFile.Caption = GetText("&File")
104     mnuFileSave.Caption = GetText("Save Data") & "..."
108     mnuFileExport.Caption = GetText("Export Setings") & "..."
112     mnuFileExit.Caption = GetText("E&xit")
116     mnuHelp.Caption = GetText("&Help")
120     mnuHelpHomePage.Caption = GetText("SWEBS Home Page") & "..."
124     mnuHelpForum.Caption = GetText("SWEBS Forum") & "..."
128     mnuHelpUpdate.Caption = GetText("Check For Update") & "..."
132     mnuHelpRegister.Caption = GetText("Register") & "..."
136     mnuHelpAbout.Caption = GetText("&About") & "..."
140     cmdOK.Caption = GetText("&OK")
144     cmdApply.Caption = GetText("&Apply")
148     cmdCancel.Caption = GetText("&Cancel")
152     fraSrvStatus.Caption = GetText("Current Service Status:")
156     lblSrvStatus.Caption = GetText("Status:")
160     cmdSrvStart.Caption = GetText("S&tart")
164     cmdSrvStop.Caption = GetText("St&op")
168     cmdSrvRestart.Caption = GetText("R&estart")
172     fraUpdate.Caption = GetText("Update Status:")
176     fraBasicStats.Caption = GetText("Basic Stats:")
180     lblMaxConnect.Caption = GetText("What is the maximum number of connections that your server can handle at any one time.")
184     lblAllowIndex.Caption = GetText("Display file list if no index is found?")
188     lblIndexFiles.Caption = GetText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
192     lblErrorPages.Caption = GetText("Where is the location of the folder which stores pages to be used when the server receives an error.")
196     lblServerName.Caption = GetText("What is the name of your server?")
200     lblPort.Caption = GetText("What port do you want to use? (Default is 80)")
204     lblWebroot.Caption = GetText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
208     lblLogFile.Caption = GetText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
212     lblCGIInterp.Caption = GetText("Where is the executable that will interpret these CGI scripts?")
216     lblCGIExt.Caption = GetText("What is the extension that is mapped to this interpreter.")
220     cmdCGINew.Caption = GetText("Add New...")
224     cmdCGIRemove.Caption = GetText("Remove...")
228     cmdvHostNew.Caption = GetText("Add New...")
232     cmdvHostRemove.Caption = GetText("Remove...")
236     lblvHostName.Caption = GetText("What is the name of this Virtual Host?")
240     lblvHostDomain.Caption = GetText("What is it's domain name?")
244     lblvHostRoot.Caption = GetText("This is the root directory where files are kept for this Virtual Host.")
248     lblvHostLog.Caption = GetText("Where do you want to keep the log file for this Virtual Host?")
252     lblNewCGITitle.Caption = GetText("Add a new CGI interpreter:")
256     lblNewCGIInterp.Caption = GetText("Where is the executable that will interpret this script type?")
260     lblNewCGIExt.Caption = GetText("What is the file extension for this file type?")
264     cmdNewCGIOK.Caption = GetText("&OK")
268     cmdNewCGICancel.Caption = GetText("&Cancel")
272     lblNewvHostTitle.Caption = GetText("Add a new Virtual Host:")
276     lblNewvHostName.Caption = GetText("What is the name of this Virtual Host?")
280     lblNewvHostDomain.Caption = GetText("What is the domain for this Virtual Host?")
284     lblNewvHostRoot.Caption = GetText("Where is the root folder for this Virtual Host?")
288     lblNewvHostLogs.Caption = GetText("Where do you want to keep the log for this Virtual Host?")
292     cmdNewvHostOK.Caption = GetText("&OK")
296     cmdNewvHostCancel.Caption = GetText("&Cancel")
300     lblDynDNSTitle.Caption = GetText("From here you can enable updates && maintance of you DynDNS.org account. To use this feature you must have a acount and setup a Dynamic DNS host. You can not add a new host via the system.")
304     lblDynDNSCurrentIP.Caption = GetText("Current IP:")
308     lblDynDNSLastUpdate.Caption = GetText("Last Update:")
312     lblDynDNSLastResult.Caption = GetText("Last Update Result:")
316     lblDynDNSHostname.Caption = GetText("DynDNS.org Hostname:")
320     lblDynDNSUsername.Caption = GetText("DynDNS.org Username:")
324     lblDynDNSPassword.Caption = GetText("DynDNS.org Password:")
328     cmdDynDNSUpdate.Caption = GetText("&Update")
332     chkDynDNSEnable.Caption = GetText("Enable DynDNS Updates?")
336     lblConfigAdvIPBind.Caption = GetText("What IP should the server listen to? (Default: Leave blank for all available)")
340     lblConfigBasicErrorLog.Caption = GetText("Where do you want to store the server error log?")
    
344     If LoadConfigData = False Then
348         RetVal = MsgBox(GetText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
352         Select Case RetVal
                Case vbAbort
356                 End
360             Case vbRetry
364                 If LoadConfigData = False Then
368                     MsgBox GetText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
372                     End
                    End If
376             Case vbIgnore
380                 MsgBox GetText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    
384     With vbaSideBar
388         .Redraw = False
392         Set cBar = .Bars.Add(, "status", GetText("System Status"))
396         Set cItem = cBar.Items.Add(, "status", GetText("Current Status"), 0)
        
400         Set cBar = .Bars.Add(, "config", GetText("Configuration"))
404         Set cItem = cBar.Items.Add(, "basic", GetText("Basic"), 0)
408         Set cItem = cBar.Items.Add(, "advanced", GetText("Advanced"), 0)
412         Set cItem = cBar.Items.Add(, "vhost", GetText("Virtual Host"), 0)
416         Set cItem = cBar.Items.Add(, "cgi", GetText("CGI"), 0)
            'I'm not going to show this for now pending more development.
420         Set cItem = cBar.Items.Add(, "dyndns", GetText("Dynamic DNS"), 0)
        
424         Set cBar = .Bars.Add(, "logs", GetText("System Logs"))
428         Set cItem = cBar.Items.Add(, "logs", GetText("View Logs"), 0)
432         .Height = Me.Height
436         .Redraw = True
        End With

440     fraStatus.ZOrder 0
444     vbaSideBar.ZOrder 0
448     tmrStatus_Timer
        '<EhFooter>
        Exit Sub

Form_Load_Err:
452     DisplayErrMsg Err.Description, "WinUI.frmMain.Form_Load", Erl, False
456     Resume Next
        '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>
    Dim lngRetVal As Long
100     If blnDirty = True Then
104         lngRetVal = MsgBox(GetText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
108         If lngRetVal = vbYes Then
112             If SaveConfigData(strConfigFile) = False Then
116                 MsgBox GetText("Data was not saved, no idea why...")
                End If
            End If
        End If
120     Me.Visible = False
124     DoEvents
        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
128     DisplayErrMsg Err.Description, "WinUI.frmMain.Form_QueryUnload", Erl, False
132     Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     LoadUser32 False
104     End
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.Form_Unload", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub lblUpdateStatus_Click()
        '<EhHeader>
        On Error GoTo lblUpdateStatus_Click_Err
        '</EhHeader>
100     If Update.Available = True Then
104         Load frmUpdate
108         frmUpdate.Show
        End If
        '<EhFooter>
        Exit Sub

lblUpdateStatus_Click_Err:
112     DisplayErrMsg Err.Description, "WinUI.frmMain.lblUpdateStatus_Click", Erl, False
116     Resume Next
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
116     txtCGIInterp.Text = Config.CGI((lstCGI.ListIndex + 1), 1)
120     txtCGIExt.Text = Config.CGI((lstCGI.ListIndex + 1), 2)
        '<EhFooter>
        Exit Sub

lstCGI_Click_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.lstCGI_Click", Erl, False
128     Resume Next
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
128     txtvHostName.Text = Config.vHost((lstvHosts.ListIndex + 1)).Name
132     txtvHostDomain.Text = Config.vHost((lstvHosts.ListIndex + 1)).Domain
136     txtvHostRoot.Text = Config.vHost((lstvHosts.ListIndex + 1)).Root
140     txtvHostLog.Text = Config.vHost((lstvHosts.ListIndex + 1)).Log
        '<EhFooter>
        Exit Sub

lstvHosts_Click_Err:
144     DisplayErrMsg Err.Description, "WinUI.frmMain.lstvHosts_Click", Erl, False
148     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuFileExit_Click", Erl, False
108     Resume Next
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
112             Print #1, GetConfigReport
116         Close 1
        End If
120     Set cDlg = Nothing
        '<EhFooter>
        Exit Sub

mnuFileExport_Click_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuFileExport_Click", Erl, False
128     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuFileReload_Click()
        '<EhHeader>
        On Error GoTo mnuFileReload_Click_Err
        '</EhHeader>
    Dim RetVal As Long
100     RetVal = MsgBox(GetText("This will reset any changes you make.\r\rDo you want to continue?"), vbYesNo + vbQuestion)
104     If RetVal = vbYes Then
108         If LoadConfigData = False Then
112             RetVal = MsgBox(GetText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
116             Select Case RetVal
                    Case vbAbort
120                     Unload Me
124                 Case vbRetry
128                     If LoadConfigData = False Then
132                         MsgBox GetText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                        End If
136                 Case vbIgnore
140                     MsgBox GetText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
                End Select
            End If
        End If
        '<EhFooter>
        Exit Sub

mnuFileReload_Click_Err:
144     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuFileReload_Click", Erl, False
148     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuFileSave_Click()
        '<EhHeader>
        On Error GoTo mnuFileSave_Click_Err
        '</EhHeader>
100     If SaveConfigData(strConfigFile) = False Then
104         MsgBox GetText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox GetText("You data has been saved./r/rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
        '<EhFooter>
        Exit Sub

mnuFileSave_Click_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuFileSave_Click", Erl, False
120     Resume Next
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
108     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuHelpAbout_Click", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuHelpForum_Click()
        '<EhHeader>
        On Error GoTo mnuHelpForum_Click_Err
        '</EhHeader>
100     OpenURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
        '<EhFooter>
        Exit Sub

mnuHelpForum_Click_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuHelpForum_Click", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuHelpHomePage_Click()
        '<EhHeader>
        On Error GoTo mnuHelpHomePage_Click_Err
        '</EhHeader>
100     OpenURL "http://swebs.sourceforge.net/html/index.php"
        '<EhFooter>
        Exit Sub

mnuHelpHomePage_Click_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuHelpHomePage_Click", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuHelpRegister_Click()
        '<EhHeader>
        On Error GoTo mnuHelpRegister_Click_Err
        '</EhHeader>
100     StartRegistration
        '<EhFooter>
        Exit Sub

mnuHelpRegister_Click_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuHelpRegister_Click", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub mnuHelpUpdate_Click()
        '<EhHeader>
        On Error GoTo mnuHelpUpdate_Click_Err
        '</EhHeader>
100     AppStatus True, GetText("Retrieving Update Information") & "..."
104     GetUpdateInfo
108     If Update.Available = True Then
112         lblUpdateStatus.Caption = GetText("New Version Available")
116         lblUpdateStatus.Font.Underline = True
120         lblUpdateStatus.ForeColor = vbBlue
124         lblUpdateStatus.MousePointer = vbCustom
128         Load frmUpdate
132         frmUpdate.Show
        Else
136         MsgBox GetText("You have the most current version available."), vbOKOnly + vbInformation
        End If
140     AppStatus False
        '<EhFooter>
        Exit Sub

mnuHelpUpdate_Click_Err:
144     DisplayErrMsg Err.Description, "WinUI.frmMain.mnuHelpUpdate_Click", Erl, False
148     Resume Next
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
108         Case icConnecting
112             DoEvents
116         Case icConnected
120             DoEvents
124         Case icRequesting
128             DoEvents
132         Case icRequestSent
136             DoEvents
140         Case icReceivingResponse
144             DoEvents
148         Case icResponseReceived
152             DoEvents
156         Case icDisconnecting
160             DoEvents
164         Case icDisconnected
168             DoEvents
172         Case icError
176             DoEvents
180         Case icResponseCompleted
184             strResult = netDynDNS.GetChunk(1024, icString)
188             DynDNS.LastIP = DynDNS.CurrentIP
192             DynDNS.LastUpdate = Now
196             DynDNS.LastResult = strResult
200             txtDynDNSLastUpdate.Text = DynDNS.LastUpdate
204             txtDynDNSLastResult.Text = DynDNS.LastResult
            
208             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", DynDNS.Hostname
212             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", DynDNS.LastIP
216             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", DynDNS.LastResult
220             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", DynDNS.LastUpdate
224             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", DynDNS.Password
228             SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", DynDNS.UserName
232             If DynDNS.Enabled = True Then
236                 SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
                Else
240                 SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
                End If
244             cmdDynDNSUpdate.Enabled = False
248             AppStatus False
252             MsgBox "Update completed. DynDNS.org returned:" & vbCrLf & vbCrLf & Chr(9) & strResult, vbInformation 'this line will go away soon, thus no GT
        End Select
        '<EhFooter>
        Exit Sub

netDynDNS_StateChanged_Err:
256     DisplayErrMsg Err.Description, "WinUI.frmMain.netDynDNS_StateChanged", Erl, False
260     Resume Next
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
            lblSrvStatusCur.Caption = GetText("Stopped")
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Start Pending"
            lblSrvStatusCur.Caption = GetText("Start Pending")
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Stop Pending"
            lblSrvStatusCur.Caption = GetText("Stop Pending")
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = GetText("Running")
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
        Case "Continue Pending"
            lblSrvStatusCur.Caption = GetText("Continue Pending")
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Pause Pending"
            lblSrvStatusCur.Caption = GetText("Pause Pending")
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = False
        Case "Paused"
            lblSrvStatusCur.Caption = GetText("Paused")
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
    End Select
End Sub

Private Sub AppStatus(blnBusy As Boolean, Optional strMessage As String = "Ready...")
        '<EhHeader>
        On Error GoTo AppStatus_Err
        '</EhHeader>
100     If blnBusy = True Then
104         Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
        Else
108         Screen.MousePointer = vbDefault  '0 default
        End If
112     lblAppStatus.Caption = GetText(strMessage)
116     DoEvents 'i'm not sure if this will stay, causes the lbl to flash for fast operations...
        '<EhFooter>
        Exit Sub

AppStatus_Err:
120     DisplayErrMsg Err.Description, "WinUI.frmMain.AppStatus", Erl, False
124     Resume Next
        '</EhFooter>
End Sub

Private Function LoadConfigData() As Boolean
        '<EhHeader>
        On Error GoTo LoadConfigData_Err
        '</EhHeader>
    Dim i As Long
    Dim strTemp As String
    Dim strResult As String
    
100     AppStatus True, GetText("Loading Configuration Data") & "..."
104     LoadConfigData = GetConfigData(strConfigFile)
    
        'Setup the form...
108     txtServerName.Text = Config.ServerName
112     txtPort.Text = Config.Port
116     txtWebroot.Text = Config.WebRoot
120     txtMaxConnect.Text = Config.MaxConnections
124     txtLogFile.Text = Config.LogFile
128     txtConfigAdvIPBind.Text = Config.ListeningAddress
132     txtAllowIndex.Text = Config.AllowIndex
136     txtErrorPages.Text = Config.ErrorPages
140     txtConfigBasicErrorLog.Text = Config.ErrorLog
    
144     For i = 1 To UBound(Config.Index)
148         strTemp = strTemp & Config.Index(i) & " "
        Next
152     txtIndexFiles.Text = Trim$(strTemp)
156     If Config.CGI(1, 2) <> "" Then
160         lstCGI.Clear
164         For i = 1 To UBound(Config.CGI)
168             lstCGI.AddItem Config.CGI(i, 2)
            Next
        Else
172         lstCGI.Enabled = False
        End If
176     If Config.vHost(1).Name <> "" Then
180         lstvHosts.Clear
184         For i = 1 To UBound(Config.vHost)
188             lstvHosts.AddItem Config.vHost(i).Name
            Next
        Else
192         lstvHosts.Enabled = False
        End If
196     cmbViewLogFiles.Clear
200     If Dir$(Config.LogFile) <> "" Then
204         cmbViewLogFiles.AddItem Config.LogFile
        End If
208     If Dir$(Config.ErrorLog) <> "" Then
212         cmbViewLogFiles.AddItem Config.ErrorLog
        End If
216     For i = 1 To UBound(Config.vHost)
220         If Dir$(Config.vHost(i).Log) <> "" Then
224             cmbViewLogFiles.AddItem Config.vHost(i).Log
            End If
        Next
    
        'we now only check for updates every 24 hours, this could confuse some people.
        'but this should make loading faster.
228     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "LastUpdateCheck")
232     If strResult = "" Then
236         strResult = CDate(1.1)
        End If
240     If DateDiff("h", CDate(strResult), Now) >= 24 Then
244         GetUpdateInfo
248         If Update.Available = True Then
252             lblUpdateStatus.Caption = GetText("New Version Available")
            Else
256             lblUpdateStatus.Caption = GetText("No Updates Available")
260             lblUpdateStatus.Font.Underline = False
264             lblUpdateStatus.ForeColor = vbButtonText
268             lblUpdateStatus.MousePointer = vbDefault
272             SaveRegistryString &H80000002, "SOFTWARE\SWS", "LastUpdateCheck", Now
            End If
        Else
276         lblUpdateStatus.Caption = GetText("No Updates Available")
280         lblUpdateStatus.Font.Underline = False
284         lblUpdateStatus.ForeColor = vbButtonText
288         lblUpdateStatus.MousePointer = vbDefault
        End If
    
292     UpdateStats
    
296     DynDNS.CurrentIP = GetLocalIP
300     txtDynDNSCurrentIP.Text = DynDNS.CurrentIP
304     txtDynDNSHostname.Text = DynDNS.Hostname
308     txtDynDNSUsername.Text = DynDNS.UserName
312     txtDynDNSLastUpdate.Text = DynDNS.LastUpdate
316     txtDynDNSLastUpdate.Enabled = False
320     txtDynDNSLastResult.Text = DynDNS.LastResult
324     txtDynDNSLastResult.Enabled = False
328     txtDynDNSPassword.Text = DynDNS.Password
332     If DynDNS.Enabled = True Then
336         chkDynDNSEnable.Value = vbChecked
        End If
340     If DynDNS.CurrentIP <> DynDNS.LastIP Or DateDiff("d", CDate(DynDNS.LastUpdate), Now) >= 28 Then
344         cmdDynDNSUpdate.Enabled = True
        Else
348         cmdDynDNSUpdate.Enabled = False
        End If
    
352     If blnRegistered = True Then
356         mnuHelpRegister.Enabled = False
            'netMain.OpenURL "http://swebs.sf.net/register/regupdate.php?email=" & UrlEncode(GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")) & "&ver=" & UrlEncode(strInstalledVer)
        End If
    
360     AppStatus False
        '<EhFooter>
        Exit Function

LoadConfigData_Err:
364     DisplayErrMsg Err.Description, "WinUI.frmMain.LoadConfigData", Erl, False
368     Resume Next
        '</EhFooter>
End Function

Private Sub txtAllowIndex_Change()
        '<EhHeader>
        On Error GoTo txtAllowIndex_Change_Err
        '</EhHeader>
100     Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
        '<EhFooter>
        Exit Sub

txtAllowIndex_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtAllowIndex_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtAllowIndex_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtAllowIndex_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtAllowIndex_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtAllowIndex_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtAllowIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtAllowIndex_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtAllowIndex_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtAllowIndex_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIExt_Change()
        '<EhHeader>
        On Error GoTo txtCGIExt_Change_Err
        '</EhHeader>
100     If lstCGI.ListIndex <> -1 Then
104         Config.CGI((lstCGI.ListIndex + 1), 2) = txtCGIExt.Text
        End If
        '<EhFooter>
        Exit Sub

txtCGIExt_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIExt_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIExt_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtCGIExt_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtCGIExt_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIExt_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIExt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtCGIExt_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtCGIExt_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIExt_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIInterp_Change()
        '<EhHeader>
        On Error GoTo txtCGIInterp_Change_Err
        '</EhHeader>
100     If lstCGI.ListIndex <> -1 Then
104         Config.CGI((lstCGI.ListIndex + 1), 1) = txtCGIInterp.Text
        End If
        '<EhFooter>
        Exit Sub

txtCGIInterp_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIInterp_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIInterp_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtCGIInterp_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtCGIInterp_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIInterp_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtCGIInterp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtCGIInterp_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtCGIInterp_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtCGIInterp_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigAdvIPBind_Change()
        '<EhHeader>
        On Error GoTo txtConfigAdvIPBind_Change_Err
        '</EhHeader>
100     Config.ListeningAddress = txtConfigAdvIPBind.Text
        '<EhFooter>
        Exit Sub

txtConfigAdvIPBind_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigAdvIPBind_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigAdvIPBind_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtConfigAdvIPBind_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtConfigAdvIPBind_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigAdvIPBind_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigAdvIPBind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtConfigAdvIPBind_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtConfigAdvIPBind_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigAdvIPBind_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigBasicErrorLog_Change()
        '<EhHeader>
        On Error GoTo txtConfigBasicErrorLog_Change_Err
        '</EhHeader>
100     Config.ErrorLog = txtConfigBasicErrorLog.Text
        '<EhFooter>
        Exit Sub

txtConfigBasicErrorLog_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigBasicErrorLog_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigBasicErrorLog_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtConfigBasicErrorLog_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtConfigBasicErrorLog_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigBasicErrorLog_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtConfigBasicErrorLog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtConfigBasicErrorLog_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtConfigBasicErrorLog_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtConfigBasicErrorLog_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSCurrentIP_Change()
        '<EhHeader>
        On Error GoTo txtDynDNSCurrentIP_Change_Err
        '</EhHeader>
100     DynDNS.CurrentIP = txtDynDNSCurrentIP.Text
104     If DynDNS.CurrentIP <> DynDNS.LastIP Or DateDiff("d", CDate(DynDNS.LastUpdate), Now) >= 28 Then
108         cmdDynDNSUpdate.Enabled = True
        Else
112         cmdDynDNSUpdate.Enabled = False
        End If
        '<EhFooter>
        Exit Sub

txtDynDNSCurrentIP_Change_Err:
116     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSCurrentIP_Change", Erl, False
120     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSCurrentIP_KeyPress", Erl, False
108     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSCurrentIP_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSHostname_Change()
        '<EhHeader>
        On Error GoTo txtDynDNSHostname_Change_Err
        '</EhHeader>
100     DynDNS.Hostname = txtDynDNSHostname.Text
        '<EhFooter>
        Exit Sub

txtDynDNSHostname_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSHostname_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSHostname_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtDynDNSHostname_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSHostname_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSHostname_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSHostname_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtDynDNSHostname_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSHostname_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSHostname_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSPassword_Change()
        '<EhHeader>
        On Error GoTo txtDynDNSPassword_Change_Err
        '</EhHeader>
100     DynDNS.Password = txtDynDNSPassword.Text
        '<EhFooter>
        Exit Sub

txtDynDNSPassword_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSPassword_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSPassword_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtDynDNSPassword_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSPassword_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSPassword_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSPassword_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtDynDNSPassword_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSPassword_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSPassword_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSUsername_Change()
        '<EhHeader>
        On Error GoTo txtDynDNSUsername_Change_Err
        '</EhHeader>
100     DynDNS.UserName = txtDynDNSUsername.Text
        '<EhFooter>
        Exit Sub

txtDynDNSUsername_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSUsername_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSUsername_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtDynDNSUsername_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSUsername_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSUsername_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtDynDNSUsername_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtDynDNSUsername_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtDynDNSUsername_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtDynDNSUsername_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtErrorPages_Change()
        '<EhHeader>
        On Error GoTo txtErrorPages_Change_Err
        '</EhHeader>
100     Config.ErrorPages = txtErrorPages.Text
        '<EhFooter>
        Exit Sub

txtErrorPages_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtErrorPages_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtErrorPages_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtErrorPages_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtErrorPages_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtErrorPages_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtErrorPages_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtErrorPages_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtErrorPages_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtErrorPages_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtIndexFiles_Change()
        '<EhHeader>
        On Error GoTo txtIndexFiles_Change_Err
        '</EhHeader>
    Dim strTmpArray() As String
    Dim lngRecCount As Long
    Dim i As Long
100     strTmpArray = Split(Trim$(txtIndexFiles.Text), " ")
104     If UBound(strTmpArray) >= 1 Then
108         ReDim Config.Index(1 To (UBound(strTmpArray) + 1))
112         lngRecCount = UBound(strTmpArray)
116         For i = 0 To lngRecCount
120             Config.Index(i + 1) = strTmpArray(i)
            Next
        End If
        '<EhFooter>
        Exit Sub

txtIndexFiles_Change_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.txtIndexFiles_Change", Erl, False
128     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtIndexFiles_KeyPress", Erl, False
108     Resume Next
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
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtIndexFiles_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtLogFile_Change()
        '<EhHeader>
        On Error GoTo txtLogFile_Change_Err
        '</EhHeader>
100     Config.LogFile = Trim$(txtLogFile.Text)
        '<EhFooter>
        Exit Sub

txtLogFile_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtLogFile_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtLogFile_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtLogFile_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtLogFile_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtLogFile_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtLogFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtLogFile_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtLogFile_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtLogFile_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtMaxConnect_Change()
        '<EhHeader>
        On Error GoTo txtMaxConnect_Change_Err
        '</EhHeader>
100     Config.MaxConnections = Int(Val(txtMaxConnect.Text))
        '<EhFooter>
        Exit Sub

txtMaxConnect_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtMaxConnect_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtMaxConnect_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtMaxConnect_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtMaxConnect_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtMaxConnect_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtMaxConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtMaxConnect_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtMaxConnect_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtMaxConnect_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtPort_Change()
        '<EhHeader>
        On Error GoTo txtPort_Change_Err
        '</EhHeader>
100     Config.Port = Int(Val(txtPort.Text))
        '<EhFooter>
        Exit Sub

txtPort_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtPort_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtPort_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtPort_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtPort_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtPort_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtPort_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtPort_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtPort_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtServerName_Change()
        '<EhHeader>
        On Error GoTo txtServerName_Change_Err
        '</EhHeader>
100     Config.ServerName = Trim$(txtServerName.Text)
        '<EhFooter>
        Exit Sub

txtServerName_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtServerName_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtServerName_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtServerName_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtServerName_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtServerName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtServerName_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtServerName_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtServerName_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostDomain_Change()
        '<EhHeader>
        On Error GoTo txtvHostDomain_Change_Err
        '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         Config.vHost((lstvHosts.ListIndex + 1)).Domain = txtvHostDomain.Text
        End If
        '<EhFooter>
        Exit Sub

txtvHostDomain_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostDomain_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostDomain_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtvHostDomain_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostDomain_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostDomain_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostDomain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtvHostDomain_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostDomain_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostDomain_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostLog_Change()
        '<EhHeader>
        On Error GoTo txtvHostLog_Change_Err
        '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         Config.vHost((lstvHosts.ListIndex + 1)).Log = txtvHostLog.Text
        End If
        '<EhFooter>
        Exit Sub

txtvHostLog_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostLog_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostLog_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtvHostLog_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostLog_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostLog_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostLog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtvHostLog_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostLog_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostLog_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostName_Change()
        '<EhHeader>
        On Error GoTo txtvHostName_Change_Err
        '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         Config.vHost((lstvHosts.ListIndex + 1)).Name = txtvHostName.Text
        End If
        '<EhFooter>
        Exit Sub

txtvHostName_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostName_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostName_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtvHostName_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostName_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostName_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtvHostName_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostName_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostName_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostRoot_Change()
        '<EhHeader>
        On Error GoTo txtvHostRoot_Change_Err
        '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         Config.vHost((lstvHosts.ListIndex + 1)).Root = txtvHostRoot.Text
        End If
        '<EhFooter>
        Exit Sub

txtvHostRoot_Change_Err:
108     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostRoot_Change", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostRoot_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtvHostRoot_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostRoot_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostRoot_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtvHostRoot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtvHostRoot_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtvHostRoot_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtvHostRoot_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtWebroot_Change()
        '<EhHeader>
        On Error GoTo txtWebroot_Change_Err
        '</EhHeader>
100     Config.WebRoot = Trim$(txtWebroot.Text)
        '<EhFooter>
        Exit Sub

txtWebroot_Change_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtWebroot_Change", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtWebroot_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtWebroot_KeyPress_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtWebroot_KeyPress_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtWebroot_KeyPress", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub txtWebroot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo txtWebroot_MouseUp_Err
        '</EhHeader>
100     blnDirty = True
        '<EhFooter>
        Exit Sub

txtWebroot_MouseUp_Err:
104     DisplayErrMsg Err.Description, "WinUI.frmMain.txtWebroot_MouseUp", Erl, False
108     Resume Next
        '</EhFooter>
End Sub

Private Sub GetUpdateInfo()
        '<EhHeader>
        On Error GoTo GetUpdateInfo_Err
        '</EhHeader>
    Dim strdata As String

        'get data from server
    
100     If GetNetStatus = True Then
104         strdata = Replace(netMain.OpenURL("http://swebs.sf.net/upgrade.xml", icString), vbLf, vbCrLf)
        End If
    
108     Call GetUpdateStatus(strdata)
        '<EhFooter>
        Exit Sub

GetUpdateInfo_Err:
112     DisplayErrMsg Err.Description, "WinUI.frmMain.GetUpdateInfo", Erl, False
116     Resume Next
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
168     DisplayErrMsg Err.Description, "WinUI.frmMain.vbaSideBar_ItemClick", Erl, False
172     Resume Next
        '</EhFooter>
End Sub

Private Sub UpdateStats()
        '<EhHeader>
        On Error GoTo UpdateStats_Err
        '</EhHeader>
100     GetStatsData
104     lblStatsLastRestart.Caption = GetText("Last Restart") & ": " & Stats.LastRestart
108     lblStatsRequestCount.Caption = GetText("Request Count") & ": " & Stats.RequestCount
112     lblStatsBytesSent.Caption = GetText("Total Bytes Sent") & ": " & Format$(Stats.TotalBytesSent, "###,###,###,###,##0")
116     lblCurVersion.Caption = GetText("Current Version") & ": " & strInstalledVer
120     lblUpdateVersion.Caption = GetText("Update Version") & ": " & IIf(Update.Version <> "", Update.Version, strInstalledVer)
        '<EhFooter>
        Exit Sub

UpdateStats_Err:
124     DisplayErrMsg Err.Description, "WinUI.frmMain.UpdateStats", Erl, False
128     Resume Next
        '</EhFooter>
End Sub

Private Function GetLocalIP() As String
        '<EhHeader>
        On Error GoTo GetLocalIP_Err
        '</EhHeader>
    Dim strResult As String

100     If GetNetStatus = True Then
104         strResult = netMain.OpenURL("http://checkip.dyndns.org/")
108         strResult = Replace(strResult, vbCr, "")
112         strResult = Replace(strResult, vbLf, "")
116         strResult = Mid(strResult, InStr(1, strResult, "Current IP Address: "), (InStr(1, strResult, "</body>") - 1) - InStr(1, strResult, "Current IP Address: ") + 1)
120         strResult = Replace(strResult, "Current IP Address: ", "")
124         GetLocalIP = strResult
        Else
128         GetLocalIP = "127.0.0.1"
        End If
        '<EhFooter>
        Exit Function

GetLocalIP_Err:
132     DisplayErrMsg Err.Description, "WinUI.frmMain.GetLocalIP", Erl, False
136     Resume Next
        '</EhFooter>
End Function
