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
         Caption         =   $"frmMain.frx":0CCA
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
               MouseIcon       =   "frmMain.frx":0D8A
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
         Picture         =   "frmMain.frx":1094
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
   Begin VB.Frame fraLogs 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   40
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtViewLogFiles 
         Enabled         =   0   'False
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Text            =   "frmMain.frx":1D5E
         Top             =   480
         Width           =   6735
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":1D84
         Left            =   120
         List            =   "frmMain.frx":1D86
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   120
         Width           =   6735
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
         ItemData        =   "frmMain.frx":1D88
         Left            =   120
         List            =   "frmMain.frx":1D8A
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
         ItemData        =   "frmMain.frx":1D8C
         Left            =   120
         List            =   "frmMain.frx":1D8E
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
   Begin VB.Frame fraConfigBasic 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   5
         Left            =   6360
         ScaleHeight     =   1215
         ScaleWidth      =   255
         TabIndex        =   56
         Top             =   2160
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
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtWebroot 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox txtLogFile 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   6015
      End
      Begin VB.Label lblLogFile 
         Caption         =   "This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here."
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   6135
      End
      Begin VB.Label lblServerName 
         Caption         =   "What is the name of your server?"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label lblPort 
         Caption         =   "What port do you want to use? (Default is 80)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label lblWebroot 
         Caption         =   $"frmMain.frx":1D90
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1680
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
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtErrorPages 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label lblMaxConnect 
         Caption         =   "What is the maximum number of connections that your server can handle at any one time."
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label lblIndexFiles 
         Caption         =   $"frmMain.frx":1E34
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label lblAllowIndex 
         Caption         =   "This allows the server print out a list of all the files in the folder if no index file can be found."
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   6135
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
    blnDirty = True
    If chkDynDNSEnable.Value = vbChecked Then
        DynDNS.Enabled = True
    Else
        DynDNS.Enabled = False
    End If
End Sub

Private Sub cmbViewLogFiles_Click()
Dim strLog As String
    
    AppStatus True, GetText("Loading Log File") & "..."
    If Dir$(cmbViewLogFiles.Text) <> "" Then
        Open cmbViewLogFiles.Text For Random As 1 Len = FileLen(cmbViewLogFiles.Text)
            Get #1, , strLog
        Close 1
    Else
        DoEvents
        MsgBox GetText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
    End If
    AppStatus False
End Sub

Private Sub cmdApply_Click()
    If SaveConfigData(strConfigFile) = False Then
        MsgBox GetText("Data was not saved, no idea why...")
    Else
        blnDirty = False
        MsgBox GetText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdBrowseCGIInterp_Click()
Dim cDlg As cCommonDialog
Dim strFile As String
Dim strStartDir As String

    Set cDlg = New cCommonDialog
    strStartDir = Mid$(Config.CGI((lstCGI.ListIndex + 1), 1), 1, (Len(Config.CGI((lstCGI.ListIndex + 1), 1)) - InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\")))
    If cDlg.VBGetOpenFileName(strFile, , True, , , , "Executable Files (*.exe)|*.exe", , strStartDir, , "exe") Then
        txtCGIInterp.Text = strFile
    End If
    Set cDlg = Nothing
End Sub

Private Sub cmdBrowseErrorPages_Click()
Dim strPath As String
    blnDirty = True
    strPath = BrowseForFolder(Me, , True, Config.ErrorPages)
    If strPath <> "" Then
        txtErrorPages.Text = strPath
    End If
End Sub

Private Sub cmdBrowseNewCGIInterp_Click()
Dim cDlg As cCommonDialog
Dim strFile As String

    Set cDlg = New cCommonDialog
    If cDlg.VBGetOpenFileName(strFile, , True, , , , "Executable Files (*.exe)|*.exe", , , , "exe") Then
        txtNewCGIInterp.Text = strFile
    End If
    Set cDlg = Nothing
End Sub

Private Sub cmdBrowseNewvHostLogs_Click()
Dim cDlg As cCommonDialog
Dim strFile As String

    Set cDlg = New cCommonDialog
    If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
        txtNewvHostLogs.Text = strFile
    End If
    Set cDlg = Nothing
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
Dim strPath As String
    strPath = BrowseForFolder(Me, , True, Config.WebRoot)
    If strPath <> "" Then
        txtNewvHostRoot.Text = strPath
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
Dim cDlg As cCommonDialog
Dim strFile As String
Dim strStartDir As String

    Set cDlg = New cCommonDialog
    blnDirty = True
    strStartDir = Mid$(Config.vHost((lstvHosts.ListIndex + 1), 4), (InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\") + 1))
    If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
        txtvHostLog.Text = strFile
    End If
    Set cDlg = Nothing
End Sub

Private Sub cmdBrowsevHostRoot_Click()
Dim strPath As String
    strPath = BrowseForFolder(Me, , True, Config.vHost((lstvHosts.ListIndex + 1), 3))
    If strPath <> "" Then
        txtvHostRoot.Text = strPath
    End If
End Sub

Private Sub cmdBrowseLogFile_Click()
Dim cDlg As cCommonDialog
Dim strFile As String
Dim strStartDir As String

    Set cDlg = New cCommonDialog
    blnDirty = True
    strStartDir = Mid$(Config.LogFile, (InStrRev(Config.LogFile, "\") + 1))
    If cDlg.VBGetSaveFileName(strFile, , , "Log Files (*.log)|*.log|All Files (*.*)|*.*") Then
        txtLogFile.Text = strFile
    End If
    Set cDlg = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCGINew_Click()
    fraNewCGI.ZOrder 0
    vbaSideBar.ZOrder 0
End Sub

Private Sub cmdCGIRemove_Click()
Dim lngRetVal As Long
Dim i As Long

    If lstCGI.ListIndex >= 0 Then
        lngRetVal = MsgBox(GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
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

Private Sub cmdDynDNSUpdate_Click()
    
    AppStatus True, "Updating DNS Information..."
    netDynDNS.URL = "http://members.dyndns.org"
    netDynDNS.Document = "/nic/update?system=dyndns&hostname=" & DynDNS.Hostname & "&myip=" & DynDNS.CurrentIP & "&wildcard=NOCHG"
    netDynDNS.UserName = DynDNS.UserName
    netDynDNS.Password = DynDNS.Password
    netDynDNS.Execute , "GET", , "User-Agent: SWEBS WinUI " & strInstalledVer & " <plenojure@users.sf.net>"
    
End Sub

Private Sub cmdNewCGICancel_Click()
    fraNewCGI.ZOrder 1
    txtNewCGIInterp.Text = ""
    txtNewCGIExt.Text = ""
End Sub

Private Sub cmdNewCGIOK_Click()
Dim i As Long

    If txtNewCGIInterp.Text <> "" And txtNewCGIExt.Text <> "" Then
        blnDirty = True
        AddNewCGI txtNewCGIExt.Text, txtNewCGIInterp.Text
        If Config.CGI(1, 2) <> "" Then
            lstCGI.Clear
            For i = 1 To UBound(Config.CGI)
                lstCGI.AddItem Config.CGI(i, 2)
            Next
        Else
            lstCGI.Enabled = False
        End If
        fraNewCGI.ZOrder 1
        txtNewCGIInterp.Text = ""
        txtNewCGIExt.Text = ""
    Else
        MsgBox GetText("Please fill all fields.")
    End If
End Sub

Private Sub cmdNewvHostCancel_Click()
    fraNewvHost.ZOrder 1
    txtNewvHostName.Text = ""
    txtNewvHostDomain.Text = ""
    txtNewvHostRoot.Text = ""
    txtNewvHostLogs.Text = ""
End Sub

Private Sub cmdNewvHostOK_Click()
Dim i As Long

    If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
        blnDirty = True
        AddNewvHost txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text
        lstvHosts.Clear
        If Config.vHost(1, 1) <> "" Then
            For i = 1 To UBound(Config.vHost)
                lstvHosts.AddItem Config.vHost(i, 1)
            Next
            lstvHosts.Enabled = True
        Else
            lstvHosts.Enabled = False
        End If
        fraNewvHost.ZOrder 1
        txtNewvHostName.Text = ""
        txtNewvHostDomain.Text = ""
        txtNewvHostRoot.Text = ""
        txtNewvHostLogs.Text = ""
    Else
        MsgBox GetText("Please fill all fields.")
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSrvRestart_Click()
    AppStatus True, GetText("Restarting Service") & "..."
    ServiceStop "", "SWEBS Web Server"
    Do Until ServiceStatus("", "SWEBS Web Server") = "Stopped"
        DoEvents
    Loop
    ServiceStart "", "SWEBS Web Server"
    UpdateStats
    AppStatus False
End Sub

Private Sub cmdSrvStart_Click()
    AppStatus True, GetText("Starting Service") & "..."
    ServiceStart "", "SWEBS Web Server"
    UpdateStats
    AppStatus False
End Sub

Private Sub cmdSrvStop_Click()
    AppStatus True, GetText("Stopping Service") & "..."
    ServiceStop "", "SWEBS Web Server"
    AppStatus False
End Sub

Private Sub cmdvHostNew_Click()
    fraNewvHost.ZOrder 0
    vbaSideBar.ZOrder 0
End Sub

Private Sub cmdvHostRemove_Click()
Dim lngRetVal As Long
Dim i As Long

    If lstvHosts.ListIndex >= 0 Then
        lngRetVal = MsgBox(GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
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
Dim cBar As cExplorerBar
Dim cItem As cExplorerBarItem

    'setup the translated strings...
    mnuFile.Caption = GetText("&File")
    mnuFileSave.Caption = GetText("Save Data") & "..."
    mnuFileExport.Caption = GetText("Export Setings") & "..."
    mnuFileExit.Caption = GetText("E&xit")
    mnuHelp.Caption = GetText("&Help")
    mnuHelpHomePage.Caption = GetText("SWEBS Home Page") & "..."
    mnuHelpForum.Caption = GetText("SWEBS Forum") & "..."
    mnuHelpUpdate.Caption = GetText("Check For Update") & "..."
    mnuHelpRegister.Caption = GetText("Register") & "..."
    mnuHelpAbout.Caption = GetText("&About") & "..."
    cmdOK.Caption = GetText("&OK")
    cmdApply.Caption = GetText("&Apply")
    cmdCancel.Caption = GetText("&Cancel")
    fraSrvStatus.Caption = GetText("Current Service Status:")
    lblSrvStatus.Caption = GetText("Status:")
    cmdSrvStart.Caption = GetText("S&tart")
    cmdSrvStop.Caption = GetText("St&op")
    cmdSrvRestart.Caption = GetText("R&estart")
    fraUpdate.Caption = GetText("Update Status:")
    fraBasicStats.Caption = GetText("Basic Stats:")
    lblMaxConnect.Caption = GetText("What is the maximum number of connections that your server can handle at any one time.")
    lblAllowIndex.Caption = GetText("This allows the server print out a list of all the files in the folder if no index file can be found.")
    lblIndexFiles.Caption = GetText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
    lblErrorPages.Caption = GetText("Where is the location of the folder which stores pages to be used when the server receives an error.")
    lblServerName.Caption = GetText("What is the name of your server?")
    lblPort.Caption = GetText("What port do you want to use? (Default is 80)")
    lblWebroot.Caption = GetText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
    lblLogFile.Caption = GetText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
    lblCGIInterp.Caption = GetText("Where is the executable that will interpret these CGI scripts?")
    lblCGIExt.Caption = GetText("What is the extension that is mapped to this interpreter.")
    cmdCGINew.Caption = GetText("Add New...")
    cmdCGIRemove.Caption = GetText("Remove...")
    cmdvHostNew.Caption = GetText("Add New...")
    cmdvHostRemove.Caption = GetText("Remove...")
    lblvHostName.Caption = GetText("What is the name of this Virtual Host?")
    lblvHostDomain.Caption = GetText("What is it's domain name?")
    lblvHostRoot.Caption = GetText("This is the root directory where files are kept for this Virtual Host.")
    lblvHostLog.Caption = GetText("Where do you want to keep the log file for this Virtual Host?")
    lblNewCGITitle.Caption = GetText("Add a new CGI interpreter:")
    lblNewCGIInterp.Caption = GetText("Where is the executable that will interpret this script type?")
    lblNewCGIExt.Caption = GetText("What is the file extension for this file type?")
    cmdNewCGIOK.Caption = GetText("&OK")
    cmdNewCGICancel.Caption = GetText("&Cancel")
    lblNewvHostTitle.Caption = GetText("Add a new Virtual Host:")
    lblNewvHostName.Caption = GetText("What is the name of this Virtual Host?")
    lblNewvHostDomain.Caption = GetText("What is the domain for this Virtual Host?")
    lblNewvHostRoot.Caption = GetText("Where is the root folder for this Virtual Host?")
    lblNewvHostLogs.Caption = GetText("Where do you want to keep the log for this Virtual Host?")
    cmdNewvHostOK.Caption = GetText("&OK")
    cmdNewvHostCancel.Caption = GetText("&Cancel")
    lblDynDNSTitle.Caption = GetText("From here you can enable updates && maintance of you DynDNS.org account. To use this feature you must have a acount and setup a Dynamic DNS host. You can not add a new host via the system.")
    lblDynDNSCurrentIP.Caption = GetText("Current IP:")
    lblDynDNSLastUpdate.Caption = GetText("Last Update:")
    lblDynDNSLastResult.Caption = GetText("Last Update Result:")
    lblDynDNSHostname.Caption = GetText("DynDNS.org Hostname:")
    lblDynDNSUsername.Caption = GetText("DynDNS.org Username:")
    lblDynDNSPassword.Caption = GetText("DynDNS.org Password:")
    cmdDynDNSUpdate.Caption = GetText("&Update")
    chkDynDNSEnable.Caption = GetText("Enable DynDNS Updates?")
    
    If LoadConfigData = False Then
        RetVal = MsgBox(GetText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                End
            Case vbRetry
                If LoadConfigData = False Then
                    MsgBox GetText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                    End
                End If
            Case vbIgnore
                MsgBox GetText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
        End Select
    End If
    
    With vbaSideBar
        .Redraw = False
        Set cBar = .Bars.Add(, "status", GetText("System Status"))
        Set cItem = cBar.Items.Add(, "status", GetText("Current Status"), 0)
        
        Set cBar = .Bars.Add(, "config", GetText("Configuration"))
        Set cItem = cBar.Items.Add(, "basic", GetText("Basic"), 0)
        Set cItem = cBar.Items.Add(, "advanced", GetText("Advanced"), 0)
        Set cItem = cBar.Items.Add(, "vhost", GetText("Virtual Host"), 0)
        Set cItem = cBar.Items.Add(, "cgi", GetText("CGI"), 0)
        Set cItem = cBar.Items.Add(, "dyndns", GetText("DynDNS.org"), 0)
        
        Set cBar = .Bars.Add(, "logs", GetText("System Logs"))
        Set cItem = cBar.Items.Add(, "logs", GetText("View Logs"), 0)
        .Height = Me.Height
        .Redraw = True
    End With

    fraStatus.ZOrder 0
    vbaSideBar.ZOrder 0
    tmrStatus_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lngRetVal As Long
    If blnDirty = True Then
        lngRetVal = MsgBox(GetText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
        If lngRetVal = vbYes Then
            If SaveConfigData(strConfigFile) = False Then
                MsgBox GetText("Data was not saved, no idea why...")
            End If
        End If
    End If
    Me.Visible = False
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LoadUser32 False
    End
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
Dim cDlg As cCommonDialog
Dim strFile As String

    Set cDlg = New cCommonDialog
    If cDlg.VBGetSaveFileName(strFile, , , "Text Files (*.txt)|*.txt|All Files (*.*)|*.*") Then
        Open strFile For Append As 1
            Print #1, GetConfigReport
        Close 1
    End If
    Set cDlg = Nothing
End Sub

Private Sub mnuFileReload_Click()
Dim RetVal As Long
    RetVal = MsgBox(GetText("This will reset any changes you make.\r\rDo you want to continue?"), vbYesNo + vbQuestion)
    If RetVal = vbYes Then
        If LoadConfigData = False Then
            RetVal = MsgBox(GetText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
            Select Case RetVal
                Case vbAbort
                    Unload Me
                Case vbRetry
                    If LoadConfigData = False Then
                        MsgBox GetText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                    End If
                Case vbIgnore
                    MsgBox GetText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    End If
End Sub

Private Sub mnuFileSave_Click()
    If SaveConfigData(strConfigFile) = False Then
        MsgBox GetText("Data was not saved, no idea why...")
    Else
        blnDirty = False
        MsgBox GetText("You data has been saved./r/rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpForum_Click()
    OpenURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
End Sub

Private Sub mnuHelpHomePage_Click()
    OpenURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub mnuHelpRegister_Click()
    StartRegistration
End Sub

Private Sub mnuHelpUpdate_Click()
    AppStatus True, GetText("Retrieving Update Information") & "..."
    GetUpdateInfo
    If Update.Available = True Then
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox GetText("You have the most current version available."), vbOKOnly + vbInformation
    End If
    AppStatus False
End Sub

Private Sub netDynDNS_StateChanged(ByVal State As Integer)
Dim strResult As String

    Select Case State
        Case icHostResolved
            DoEvents
        Case icConnecting
            DoEvents
        Case icConnected
            DoEvents
        Case icRequesting
            DoEvents
        Case icRequestSent
            DoEvents
        Case icReceivingResponse
            DoEvents
        Case icResponseReceived
            DoEvents
        Case icDisconnecting
            DoEvents
        Case icDisconnected
            DoEvents
        Case icError
            DoEvents
        Case icResponseCompleted
            strResult = netDynDNS.GetChunk(1024, icString)
            DynDNS.LastIP = DynDNS.CurrentIP
            DynDNS.LastUpdate = Now
            DynDNS.LastResult = strResult
            txtDynDNSLastUpdate.Text = DynDNS.LastUpdate
            txtDynDNSLastResult.Text = DynDNS.LastResult
            
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", DynDNS.Hostname
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", DynDNS.LastIP
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", DynDNS.LastResult
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", DynDNS.LastUpdate
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", DynDNS.Password
            SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", DynDNS.UserName
            If DynDNS.Enabled = True Then
                SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
            Else
                SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
            End If
            cmdDynDNSUpdate.Enabled = False
            AppStatus False
            MsgBox "Update completed. DynDNS.org returned:" & vbCrLf & vbCrLf & Chr(9) & strResult, vbInformation 'this line will go away soon, thus no GT
    End Select
End Sub

Private Sub tmrStatus_Timer()
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
    If blnBusy = True Then
        Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
    Else
        Screen.MousePointer = vbDefault  '0 default
    End If
    lblAppStatus.Caption = GetText(strMessage)
    DoEvents 'i'm not sure if this will stay, causes the lbl to flash for fast operations...
End Sub

Private Function LoadConfigData() As Boolean
Dim i As Long
Dim strTemp As String
    
    AppStatus True, GetText("Loading Configuration Data") & "..."
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
    
    GetUpdateInfo
    If Update.Available = True Then
        lblUpdateStatus.Caption = GetText("New Version Available")
    Else
        lblUpdateStatus.Caption = GetText("No Updates Available")
        lblUpdateStatus.Font.Underline = False
        lblUpdateStatus.ForeColor = vbButtonText
        lblUpdateStatus.MousePointer = vbDefault
    End If
    
    GetStatsData
    UpdateStats
    
    DynDNS.CurrentIP = GetLocalIP
    txtDynDNSCurrentIP.Text = DynDNS.CurrentIP
    txtDynDNSHostname.Text = DynDNS.Hostname
    txtDynDNSUsername.Text = DynDNS.UserName
    txtDynDNSLastUpdate.Text = DynDNS.LastUpdate
    txtDynDNSLastUpdate.Enabled = False
    txtDynDNSLastResult.Text = DynDNS.LastResult
    txtDynDNSLastResult.Enabled = False
    txtDynDNSPassword.Text = DynDNS.Password
    If DynDNS.Enabled = True Then
        chkDynDNSEnable.Value = vbChecked
    End If
    If DynDNS.CurrentIP <> DynDNS.LastIP Or DateDiff("d", CDate(DynDNS.LastUpdate), Now) >= 28 Then
        cmdDynDNSUpdate.Enabled = True
    Else
        cmdDynDNSUpdate.Enabled = False
    End If
    
    If blnRegistered = True Then
        mnuHelpRegister.Enabled = False
        'netMain.OpenURL "http://swebs.sf.net/register/regupdate.php?email=" & UrlEncode(GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")) & "&ver=" & UrlEncode(strInstalledVer)
    End If
    
    AppStatus False
End Function

Private Sub txtAllowIndex_Change()
    Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
End Sub

Private Sub txtAllowIndex_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtAllowIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtCGIExt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtCGIInterp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtDynDNSCurrentIP_Change()
    DynDNS.CurrentIP = txtDynDNSCurrentIP.Text
    If DynDNS.CurrentIP <> DynDNS.LastIP Or DateDiff("d", CDate(DynDNS.LastUpdate), Now) >= 28 Then
        cmdDynDNSUpdate.Enabled = True
    Else
        cmdDynDNSUpdate.Enabled = False
    End If
End Sub

Private Sub txtDynDNSCurrentIP_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtDynDNSCurrentIP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtDynDNSHostname_Change()
    DynDNS.Hostname = txtDynDNSHostname.Text
End Sub

Private Sub txtDynDNSHostname_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtDynDNSHostname_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtDynDNSPassword_Change()
    DynDNS.Password = txtDynDNSPassword.Text
End Sub

Private Sub txtDynDNSPassword_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtDynDNSPassword_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtDynDNSUsername_Change()
    DynDNS.UserName = txtDynDNSUsername.Text
End Sub

Private Sub txtDynDNSUsername_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtDynDNSUsername_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtErrorPages_Change()
    Config.ErrorPages = txtErrorPages.Text
End Sub

Private Sub txtErrorPages_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtErrorPages_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtIndexFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtLogFile_Change()
    Config.LogFile = Trim$(txtLogFile.Text)
End Sub

Private Sub txtLogFile_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtLogFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtMaxConnect_Change()
    Config.MaxConnections = Int(Val(txtMaxConnect.Text))
End Sub

Private Sub txtMaxConnect_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtMaxConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtPort_Change()
    Config.Port = Int(Val(txtPort.Text))
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtPort_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtServerName_Change()
    Config.ServerName = Trim$(txtServerName.Text)
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtServerName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtvHostDomain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtvHostLog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtvHostName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtvHostRoot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub txtWebroot_Change()
    Config.WebRoot = Trim$(txtWebroot.Text)
End Sub

Private Sub txtWebroot_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtWebroot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDirty = True
End Sub

Private Sub GetUpdateInfo()
Dim strdata As String

    'get data from server
    
    If GetNetStatus = True Then
        strdata = Replace(netMain.OpenURL("http://swebs.sf.net/upgrade.xml", icString), vbLf, vbCrLf)
    End If
    
    Call GetUpdateStatus(strdata)
End Sub

Private Sub vbaSideBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    StopWinUpdate Me.hWnd
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
        Case "dyndns"
            fraConfigDynDns.ZOrder 0
        Case "logs"
            fraLogs.ZOrder 0
    End Select
    vbaSideBar.ZOrder 0
    StopWinUpdate
End Sub

Private Sub UpdateStats()
    GetStatsData
    lblStatsLastRestart.Caption = GetText("Last Restart") & ": " & Stats.LastRestart
    lblStatsRequestCount.Caption = GetText("Request Count") & ": " & Stats.RequestCount
    lblStatsBytesSent.Caption = GetText("Total Bytes Sent") & ": " & Format$(Stats.TotalBytesSent, "###,###,###,###,##0")
    lblCurVersion.Caption = GetText("Current Version") & ": " & strInstalledVer
    lblUpdateVersion.Caption = GetText("Update Version") & ": " & IIf(Update.Version <> "", Update.Version, strInstalledVer)
End Sub

Private Function GetLocalIP() As String
Dim strResult As String

    If GetNetStatus = True Then
        strResult = netMain.OpenURL("http://checkip.dyndns.org/")
        strResult = Replace(strResult, vbCr, "")
        strResult = Replace(strResult, vbLf, "")
        strResult = Right(strResult, InStr(1, strResult, "Current IP Address: ") + 1)
        strResult = Left(strResult, InStr(1, strResult, "<br>") - 1)
        strResult = Replace(strResult, "Current IP Address: ", "")
        GetLocalIP = strResult
    Else
        GetLocalIP = "127.0.0.1"
    End If
End Function
