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
   Begin VB.Frame fraNewISAPI 
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
         Begin VB.CommandButton cmdNewISAPICancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1080
            TabIndex        =   81
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdNewISAPIOK 
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
         Begin VB.CommandButton cmdBrowseNewISAPIInterp 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewISAPIExt 
         Height          =   285
         Left            =   1080
         TabIndex        =   76
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNewISAPIInterp 
         Height          =   285
         Left            =   1080
         TabIndex        =   74
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label lblNewISAPIIExt 
         Caption         =   "What is the file extension for this file type?"
         Height          =   255
         Left            =   840
         TabIndex        =   75
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lblNewISAPIInterp 
         Caption         =   "Where is the executable that will interpret this script type?"
         Height          =   255
         Left            =   840
         TabIndex        =   73
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label lblNewISAPITitle 
         Caption         =   "Add a new ISAPI interpreter:"
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame fraConfigISAPI 
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
         Begin VB.CommandButton cmdISAPIRemove 
            Caption         =   "Remove..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   46
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdISAPINew 
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
         Begin VB.CommandButton cmdBrowseISAPIInterp 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.ListBox lstISAPI 
         Height          =   3375
         ItemData        =   "frmMain.frx":0CCA
         Left            =   120
         List            =   "frmMain.frx":0CD1
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtISAPIInterp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtISAPIExt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblISAPIInterp 
         Caption         =   "Where is the ISAPI Plugin?"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblISAPIExt 
         Caption         =   "What is the extension that is mapped to this interpreter."
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   1080
         Width           =   4815
      End
   End
   Begin SWEBS_WinUI.ctxHookMenu ctxXPMenu 
      Left            =   5280
      Top             =   3960
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   13
      Bmp:1           =   "frmMain.frx":0CDF
      Mask:1          =   12632256
      Key:1           =   "#mnuFileSave"
      Bmp:2           =   "frmMain.frx":0DF1
      Mask:2          =   12632256
      Key:2           =   "#mnuHelpRegister"
      Bmp:3           =   "frmMain.frx":0F03
      Mask:3          =   12632256
      Key:3           =   "#mnuHelpUpdate"
      Bmp:4           =   "frmMain.frx":1015
      Mask:4          =   12632256
      Key:4           =   "#mnuFileExit"
      Bmp:5           =   "frmMain.frx":1127
      Mask:5          =   12632256
      Key:5           =   "#mnuHelpForum"
      Bmp:6           =   "frmMain.frx":1239
      Mask:6          =   13355979
      Key:6           =   "#mnuHelpHomePage"
      Bmp:7           =   "frmMain.frx":178B
      Mask:7          =   13553358
      Key:7           =   "#mnuHelpAbout"
      Bmp:8           =   "frmMain.frx":1CDD
      Mask:8          =   13355979
      Key:8           =   "#mnuFileExport"
      Bmp:9           =   "frmMain.frx":222F
      Mask:9          =   12632256
      Key:9           =   "#mnuSysTrayPopupExit"
      Bmp:10          =   "frmMain.frx":2341
      Mask:10         =   13553358
      Key:10          =   "#mnuSysTrayPopupAbout"
      Bmp:11          =   "frmMain.frx":2893
      Mask:11         =   12632256
      Key:11          =   "#mnuSysTrayPopupUpdate"
      Bmp:12          =   "frmMain.frx":29A5
      Mask:12         =   12632256
      Key:12          =   "#mnuSysTrayPopupForum"
      Bmp:13          =   "frmMain.frx":2AB7
      Mask:13         =   13355979
      Key:13          =   "#mnuSysTrayPopupHomePage"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
               MouseIcon       =   "frmMain.frx":3009
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
         Picture         =   "frmMain.frx":3313
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
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":3FDD
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Height          =   315
         ItemData        =   "frmMain.frx":405F
         Left            =   120
         List            =   "frmMain.frx":4061
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
         ItemData        =   "frmMain.frx":4063
         Left            =   120
         List            =   "frmMain.frx":4065
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
         Caption         =   $"frmMain.frx":4067
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
         Caption         =   $"frmMain.frx":410B
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
'CSEH: WinUI - Custom
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
    '<EhHeader>
    On Error GoTo cmbViewLogFiles_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmbViewLogFiles_Click")
    '</EhHeader>
    Dim strLog As String
    
100     SetStatus WinUI.GetTranslatedText("Loading Log File") & "...", True
104     If Dir$(cmbViewLogFiles.Text) <> "" Then
108         DoEvents
112         strLog = Space$(FileLen(cmbViewLogFiles.Text))
116         Open cmbViewLogFiles.Text For Binary As 1
120             Get #1, 1, strLog
124         Close 1
128         rtfViewLogFiles.Text = Replace(strLog, vbCr, "")
132         rtfViewLogFiles.RightMargin = 214748364
136         rtfViewLogFiles.SetFocus
        Else
140         DoEvents
144         MsgBox WinUI.GetTranslatedText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
        End If
148     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmbViewLogFiles_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmbViewLogFiles_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdApply_Click()
    '<EhHeader>
    On Error GoTo cmdApply_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdApply_Click")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.Save(WinUI.Server.HTTP.Config.File) = False Then
104         MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox WinUI.GetTranslatedText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdApply_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdApply_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseISAPIInterp_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseISAPIInterp_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseISAPIInterp_Click")
    '</EhHeader>
    Dim strDefaultFile As String
100     blnDirty = True
104     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
108     dlgMain.Filter = WinUI.GetTranslatedText("ISAPI Plugin Files (*.dll)|*.dll|All Files (*.*)|*.*")
112     strDefaultFile = Mid$(WinUI.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, (InStrRev(WinUI.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, "\") + 1))
116     dlgMain.FileName = strDefaultFile
120     dlgMain.InitDir = Mid$(WinUI.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, 1, (Len(WinUI.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter) - InStrRev(WinUI.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, "\")))
124     dlgMain.ShowSave
128     If dlgMain.FileName <> strDefaultFile Then
132         txtISAPIInterp.Text = dlgMain.FileName
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseISAPIInterp_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseISAPIInterp_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseErrorLog_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseErrorLog_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseErrorLog_Click")
    '</EhHeader>
    Dim strDefaultFile As String

100     blnDirty = True
104     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
108     dlgMain.Filter = WinUI.GetTranslatedText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
112     strDefaultFile = Mid$(WinUI.Server.HTTP.Config.ErrorLog, (InStrRev(WinUI.Server.HTTP.Config.ErrorLog, "\") + 1))
116     dlgMain.FileName = strDefaultFile
120     dlgMain.InitDir = WinUI.Path
124     dlgMain.ShowSave
128     If dlgMain.FileName <> strDefaultFile Then
132         txtvHostLog.Text = dlgMain.FileName
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseErrorLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseErrorLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseErrorPages_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseErrorPages_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseErrorPages_Click")
    '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.ErrorPages)
108     If strPath <> "" Then
112         txtErrorPages.Text = strPath
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseErrorPages_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseErrorPages_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewISAPIInterp_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewISAPIInterp_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseNewISAPIInterp_Click")
    '</EhHeader>
100     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
104     dlgMain.Filter = WinUI.GetTranslatedText("ISAPI Plgin Files (*.dll)|*.dll|All Files (*.*)|*.*")
108     dlgMain.ShowSave
112     If dlgMain.FileName <> "" Then
116         txtNewISAPIInterp.Text = dlgMain.FileName
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseNewISAPIInterp_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewISAPIInterp_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewvHostLogs_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewvHostLogs_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseNewvHostLogs_Click")
    '</EhHeader>
100     blnDirty = True
104     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
108     dlgMain.Filter = WinUI.GetTranslatedText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
112     dlgMain.InitDir = WinUI.Path
116     dlgMain.ShowSave
120     txtvHostLog.Text = dlgMain.FileName
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseNewvHostLogs_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewvHostLogs_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseNewvHostRoot_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseNewvHostRoot_Click")
    '</EhHeader>
    Dim strPath As String
100     strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.WebRoot)
104     If strPath <> "" Then
108         txtNewvHostRoot.Text = strPath
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseNewvHostRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseNewvHostRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseRoot_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseRoot_Click")
    '</EhHeader>
    Dim strPath As String
100     blnDirty = True
104     strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.WebRoot)
108     If strPath <> "" Then
112         txtWebroot.Text = strPath
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowsevHostLog_Click()
    '<EhHeader>
    On Error GoTo cmdBrowsevHostLog_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowsevHostLog_Click")
    '</EhHeader>
    Dim strDefaultFile As String

100     blnDirty = True
104     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
108     dlgMain.Filter = WinUI.GetTranslatedText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
112     strDefaultFile = Mid$(WinUI.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, (InStrRev(WinUI.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, "\") + 1))
116     dlgMain.FileName = strDefaultFile
120     dlgMain.InitDir = Mid$(WinUI.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, 1, (Len(WinUI.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log) - InStrRev(WinUI.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, "\")))
124     dlgMain.ShowSave
128     If dlgMain.FileName <> strDefaultFile Then
132         txtvHostLog.Text = dlgMain.FileName
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowsevHostLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowsevHostLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowsevHostRoot_Click()
    '<EhHeader>
    On Error GoTo cmdBrowsevHostRoot_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowsevHostRoot_Click")
    '</EhHeader>
    Dim strPath As String
100     strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.VirtHost((lstvHosts.ListIndex + 1)).Root)
104     If strPath <> "" Then
108         txtvHostRoot.Text = strPath
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowsevHostRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowsevHostRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBrowseLogFile_Click()
    '<EhHeader>
    On Error GoTo cmdBrowseLogFile_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdBrowseLogFile_Click")
    '</EhHeader>
    Dim strDefaultFile As String

100     blnDirty = True
104     dlgMain.DialogTitle = WinUI.GetTranslatedText("Please select a file...")
108     dlgMain.Filter = WinUI.GetTranslatedText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
112     strDefaultFile = Mid$(WinUI.Server.HTTP.Config.LogFile, (InStrRev(WinUI.Server.HTTP.Config.LogFile, "\") + 1))
116     dlgMain.FileName = strDefaultFile
120     dlgMain.InitDir = Mid$(WinUI.Server.HTTP.Config.LogFile, 1, (Len(WinUI.Server.HTTP.Config.LogFile) - InStrRev(WinUI.Server.HTTP.Config.LogFile, "\")))
124     dlgMain.ShowSave
128     If dlgMain.FileName <> strDefaultFile Then
132         txtLogFile.Text = dlgMain.FileName
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdBrowseLogFile_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdBrowseLogFile_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdExit_Click()
    '<EhHeader>
    On Error GoTo cmdExit_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdExit_Click")
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdExit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdExit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdISAPINew_Click()
    '<EhHeader>
    On Error GoTo cmdISAPINew_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdISAPINew_Click")
    '</EhHeader>
100     fraNewISAPI.ZOrder 0
104     vbaSideBar.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdISAPINew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdISAPINew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdISAPIRemove_Click()
    '<EhHeader>
    On Error GoTo cmdISAPIRemove_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdISAPIRemove_Click")
    '</EhHeader>
    Dim lngRetVal As Long
    Dim vItem As Variant
    Dim i As Long

100     If lstISAPI.ListIndex >= 0 Then
104         lngRetVal = MsgBox(WinUI.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
108         If lngRetVal = vbYes Then
112             blnDirty = True
116             WinUI.Server.HTTP.Config.ISAPI.Remove (lstISAPI.Text)
120             lstISAPI.Clear
124             If WinUI.Server.HTTP.Config.ISAPI.Count > 0 Then
128                 For Each vItem In WinUI.Server.HTTP.Config.ISAPI
132                     lstISAPI.AddItem vItem.Extension
136                     lstISAPI.Enabled = True
                    Next
                Else
140                 lstISAPI.Enabled = False
144                 cmdBrowseISAPIInterp.Enabled = False
148                 cmdISAPIRemove.Enabled = False
152                 txtISAPIInterp.Enabled = False
156                 txtISAPIExt.Enabled = False
160                 txtISAPIInterp.Text = ""
164                 txtISAPIExt.Text = ""
                End If
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdISAPIRemove_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdISAPIRemove_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNewISAPICancel_Click()
    '<EhHeader>
    On Error GoTo cmdNewISAPICancel_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdNewISAPICancel_Click")
    '</EhHeader>
100     fraNewISAPI.ZOrder 1
104     txtNewISAPIInterp.Text = ""
108     txtNewISAPIExt.Text = ""
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdNewISAPICancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewISAPICancel_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNewISAPIOK_Click()
    '<EhHeader>
    On Error GoTo cmdNewISAPIOK_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdNewISAPIOK_Click")
    '</EhHeader>
    Dim vItem As Variant
    Dim i As Long

100     If txtNewISAPIInterp.Text <> "" And txtNewISAPIExt.Text <> "" Then
104         blnDirty = True
108         WinUI.Server.HTTP.Config.ISAPI.Add txtNewISAPIInterp.Text, txtNewISAPIExt.Text, txtNewISAPIExt.Text
112         If WinUI.Server.HTTP.Config.ISAPI.Count > 0 Then
116             lstISAPI.Clear
120             For Each vItem In WinUI.Server.HTTP.Config.ISAPI
124                 lstISAPI.AddItem vItem.Extension
128                 lstISAPI.Enabled = True
                Next
            Else
132             lstISAPI.Enabled = False
            End If
136         fraNewISAPI.ZOrder 1
140         txtNewISAPIInterp.Text = ""
144         txtNewISAPIExt.Text = ""
        Else
148         MsgBox WinUI.GetTranslatedText("Please fill all fields.")
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdNewISAPIOK_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewISAPIOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNewvHostCancel_Click()
    '<EhHeader>
    On Error GoTo cmdNewvHostCancel_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdNewvHostCancel_Click")
    '</EhHeader>
100     fraNewvHost.ZOrder 1
104     txtNewvHostName.Text = ""
108     txtNewvHostDomain.Text = ""
112     txtNewvHostRoot.Text = ""
116     txtNewvHostLogs.Text = ""
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdNewvHostCancel_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewvHostCancel_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNewvHostOK_Click()
    '<EhHeader>
    On Error GoTo cmdNewvHostOK_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdNewvHostOK_Click")
    '</EhHeader>
    Dim vItem As Variant
    Dim i As Long

100     If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
104         blnDirty = True
108         WinUI.Server.HTTP.Config.VirtHost.Add txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text, txtNewvHostName.Text
112         lstvHosts.Clear
116         If WinUI.Server.HTTP.Config.VirtHost.Count > 0 Then
120             For Each vItem In WinUI.Server.HTTP.Config.VirtHost
124                 lstvHosts.AddItem vItem.HostName
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
156         MsgBox WinUI.GetTranslatedText("Please fill all fields.")
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdNewvHostOK_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdNewvHostOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdOK_Click")
    '</EhHeader>
100     If blnDirty <> False Then
104         If WinUI.Server.HTTP.Config.Save(WinUI.Server.HTTP.Config.File) = False Then
108             MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
            Else
112             blnDirty = False
116             WinUI.Server.HTTP.StopServer
120             DoEvents
124             WinUI.Server.HTTP.StartServer
128             UpdateStats
132             Me.Hide
            End If
        Else
136         Me.WindowState = vbMinimized
140         Me.Hide
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdOK_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvRestart_Click()
    '<EhHeader>
    On Error GoTo cmdSrvRestart_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdSrvRestart_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Restarting Service") & "...", True
104     WinUI.Server.HTTP.StopServer
108     DoEvents
112     WinUI.Server.HTTP.StartServer
116     UpdateStats
120     SetStatus WinUI.GetTranslatedText("Ready") & "..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdSrvRestart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvRestart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvStart_Click()
    '<EhHeader>
    On Error GoTo cmdSrvStart_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdSrvStart_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Starting Service") & "...", True
104     WinUI.Server.HTTP.StartServer
108     UpdateStats
112     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdSrvStart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvStart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSrvStop_Click()
    '<EhHeader>
    On Error GoTo cmdSrvStop_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdSrvStop_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Stopping Service") & "...", True
104     WinUI.Server.HTTP.StopServer
108     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdSrvStop_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdSrvStop_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdvHostNew_Click()
    '<EhHeader>
    On Error GoTo cmdvHostNew_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdvHostNew_Click")
    '</EhHeader>
100     fraNewvHost.ZOrder 0
104     vbaSideBar.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdvHostNew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdvHostNew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdvHostRemove_Click()
    '<EhHeader>
    On Error GoTo cmdvHostRemove_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.cmdvHostRemove_Click")
    '</EhHeader>
    Dim lngRetVal As Long
    Dim blnMore As Boolean
    Dim vItem As Variant

100     If lstvHosts.ListIndex >= 0 Then
104         lngRetVal = MsgBox(WinUI.GetTranslatedText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
108         If lngRetVal = vbYes Then
112             blnDirty = True
116             WinUI.Server.HTTP.Config.VirtHost.Remove lstvHosts.Text
120             txtvHostName.Text = ""
124             txtvHostDomain.Text = ""
128             txtvHostRoot.Text = ""
132             txtvHostLog.Text = ""
136             lstvHosts.Clear
140             For Each vItem In WinUI.Server.HTTP.Config.VirtHost
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
    WinUI.Debuger.CallStack.Pop
    Exit Sub

cmdvHostRemove_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.cmdvHostRemove_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.Form_Load")
    '</EhHeader>
    Dim RetVal As Long
    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
    
        'setup the translated strings...
100     SetStatus "Loading Translated Strings..."
    
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
152     cmdExit.Caption = WinUI.GetTranslatedText("E&xit")
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
216     lblISAPIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret these CGI scripts?")
220     lblISAPIExt.Caption = WinUI.GetTranslatedText("What is the extension that is mapped to this interpreter.")
224     cmdISAPINew.Caption = WinUI.GetTranslatedText("Add New...")
228     cmdISAPIRemove.Caption = WinUI.GetTranslatedText("Remove...")
232     cmdvHostNew.Caption = WinUI.GetTranslatedText("Add New...")
236     cmdvHostRemove.Caption = WinUI.GetTranslatedText("Remove...")
240     lblvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
244     lblvHostDomain.Caption = WinUI.GetTranslatedText("What is it's domain name?")
248     lblvHostRoot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept for this Virtual Host.")
252     lblvHostLog.Caption = WinUI.GetTranslatedText("Where do you want to keep the log file for this Virtual Host?")
256     lblNewISAPITitle.Caption = WinUI.GetTranslatedText("Add a new CGI interpreter:")
260     lblNewISAPIInterp.Caption = WinUI.GetTranslatedText("Where is the executable that will interpret this script type?")
264     lblNewISAPIIExt.Caption = WinUI.GetTranslatedText("What is the file extension for this file type?")
268     cmdNewISAPIOK.Caption = WinUI.GetTranslatedText("&OK")
272     cmdNewISAPICancel.Caption = WinUI.GetTranslatedText("&Cancel")
276     lblNewvHostTitle.Caption = WinUI.GetTranslatedText("Add a new Virtual Host:")
280     lblNewvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
284     lblNewvHostDomain.Caption = WinUI.GetTranslatedText("What is the domain for this Virtual Host?")
288     lblNewvHostRoot.Caption = WinUI.GetTranslatedText("Where is the root folder for this Virtual Host?")
292     lblNewvHostLogs.Caption = WinUI.GetTranslatedText("Where do you want to keep the log for this Virtual Host?")
296     cmdNewvHostOK.Caption = WinUI.GetTranslatedText("&OK")
300     cmdNewvHostCancel.Caption = WinUI.GetTranslatedText("&Cancel")
304     lblConfigAdvIPBind.Caption = WinUI.GetTranslatedText("What IP should the server listen to? (Default: Leave blank for all available)")
308     lblConfigBasicErrorLog.Caption = WinUI.GetTranslatedText("Where do you want to store the server error log?")
    
312     If LoadConfigData = False Then
316         RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
320         Select Case RetVal
                Case vbAbort
324                 End
328             Case vbRetry
332                 If LoadConfigData = False Then
336                     MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
340                     End
                    End If
344             Case vbIgnore
348                 MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    
352     With vbaSideBar
356         .Redraw = False
360         Set cBar = .Bars.Add(, "status", WinUI.GetTranslatedText("System Status"))
364         Set cItem = cBar.Items.Add(, "status", WinUI.GetTranslatedText("Current Status"), 0)
        
368         Set cBar = .Bars.Add(, "config", WinUI.GetTranslatedText("Configuration"))
372         Set cItem = cBar.Items.Add(, "basic", WinUI.GetTranslatedText("Basic"), 0)
376         Set cItem = cBar.Items.Add(, "advanced", WinUI.GetTranslatedText("Advanced"), 0)
380         Set cItem = cBar.Items.Add(, "vhost", WinUI.GetTranslatedText("Virtual Host"), 0)
384         Set cItem = cBar.Items.Add(, "isapi", WinUI.GetTranslatedText("ISAPI Plugins"), 0)
        
388         Set cBar = .Bars.Add(, "logs", WinUI.GetTranslatedText("System Logs"))
392         Set cItem = cBar.Items.Add(, "logs", WinUI.GetTranslatedText("View Logs"), 0)
396         .Height = Me.Height
400         .Redraw = True
        End With
    
404     Set SysTray = New cSysTray
408     Set SysTray.SourceWindow = Me
412     SysTray.IconInSysTray
416     SysTray.ToolTip = WinUI.GetTranslatedText("SWEBS Web Server") & " " & WinUI.Version
420     SysTray.Icon = Me.Icon

424     fraStatus.ZOrder 0
428     vbaSideBar.ZOrder 0
432     tmrStatus_Timer
436     SetStatus WinUI.GetTranslatedText("Ready") & "..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '<EhHeader>
    On Error GoTo Form_QueryUnload_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.Form_QueryUnload")
    '</EhHeader>
    Dim lngRetVal As Long

100     If UnloadMode <> vbFormControlMenu Then
104         If blnDirty = True Then
108             lngRetVal = MsgBox(WinUI.GetTranslatedText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
112             If lngRetVal = vbYes Then
116                 If WinUI.Server.HTTP.Config.Save(WinUI.Server.HTTP.Config.File) = False Then
120                     MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
124                     Cancel = True
                    End If
                End If
            End If
        Else
128         Cancel = True
132         Me.WindowState = vbMinimized
136         Me.Hide
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_QueryUnload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_QueryUnload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.Form_Unload")
    '</EhHeader>
    Dim i As Long

100     Me.Hide
104     PostMessage Me.hwnd, 0&, 0&, 0&
108     DoEvents
112     SysTray.RemoveFromSysTray
116     Set SysTray = Nothing
120     DoEvents
124     For i = Forms.Count - 1 To 0 Step -1
128         Unload Forms(i)
        Next
132     WinUI.Util.LoadUser32 False
136     Set WinUI = Nothing
140     SetExceptionFilter False
144     End
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_Unload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_Unload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblUpdateStatus_Click()
    '<EhHeader>
    On Error GoTo lblUpdateStatus_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblUpdateStatus_Click")
    '</EhHeader>
100     If WinUI.Update.IsAvailable = True Then
104         Load frmUpdate
108         frmUpdate.Show
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblUpdateStatus_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblUpdateStatus_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstISAPI_Click()
    '<EhHeader>
    On Error GoTo lstISAPI_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lstISAPI_Click")
    '</EhHeader>
100     cmdBrowseISAPIInterp.Enabled = True
104     cmdISAPIRemove.Enabled = True
108     txtISAPIInterp.Enabled = True
112     txtISAPIExt.Enabled = True
116     txtISAPIInterp.Text = WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter
120     txtISAPIExt.Text = WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lstISAPI_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lstISAPI_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstvHosts_Click()
    '<EhHeader>
    On Error GoTo lstvHosts_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lstvHosts_Click")
    '</EhHeader>
100     cmdBrowsevHostRoot.Enabled = True
104     cmdBrowsevHostLog.Enabled = True
108     cmdvHostRemove.Enabled = True
112     txtvHostName.Enabled = True
116     txtvHostDomain.Enabled = True
120     txtvHostRoot.Enabled = True
124     txtvHostLog.Enabled = True
128     txtvHostName.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName
132     txtvHostDomain.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain
136     txtvHostRoot.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root
140     txtvHostLog.Text = WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lstvHosts_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lstvHosts_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileExit_Click()
    '<EhHeader>
    On Error GoTo mnuFileExit_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuFileExit_Click")
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuFileExit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileExit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileExport_Click()
        'this needs some kind of error control, file checks, etc..
    '<EhHeader>
    On Error GoTo mnuFileExport_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuFileExport_Click")
    '</EhHeader>
100     dlgMain.DialogTitle = "Please select a file..."
104     dlgMain.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
108     dlgMain.ShowSave
112     If dlgMain.FileName <> "" Then
116         Open dlgMain.FileName For Append As 1
120             Print #1, WinUI.Server.HTTP.Config.Report
124         Close 1
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuFileExport_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileExport_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileReload_Click()
    '<EhHeader>
    On Error GoTo mnuFileReload_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuFileReload_Click")
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
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuFileReload_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileReload_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileSave_Click()
    '<EhHeader>
    On Error GoTo mnuFileSave_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuFileSave_Click")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.Save(WinUI.Server.HTTP.Config.File) = False Then
104         MsgBox WinUI.GetTranslatedText("Data was not saved, no idea why...")
        Else
108         blnDirty = False
112         MsgBox WinUI.GetTranslatedText("You data has been saved./r/rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuFileSave_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuFileSave_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpAbout_Click()
    '<EhHeader>
    On Error GoTo mnuHelpAbout_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpAbout_Click")
    '</EhHeader>
100     Load frmAbout
104     frmAbout.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpAbout_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpAbout_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpEventViewer_Click()
    '<EhHeader>
    On Error GoTo mnuHelpEventViewer_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpEventViewer_Click")
    '</EhHeader>
100     Load frmEventView
104     frmEventView.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpEventViewer_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpEventViewer_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpForum_Click()
    '<EhHeader>
    On Error GoTo mnuHelpForum_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpForum_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpForum_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpForum_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpHomePage_Click()
    '<EhHeader>
    On Error GoTo mnuHelpHomePage_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpHomePage_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpHomePage_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpHomePage_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpRegister_Click()
    '<EhHeader>
    On Error GoTo mnuHelpRegister_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpRegister_Click")
    '</EhHeader>
100     WinUI.Registration.Start
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpRegister_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpRegister_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuHelpUpdate_Click()
    '<EhHeader>
    On Error GoTo mnuHelpUpdate_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuHelpUpdate_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Retrieving Update Information") & "...", True
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
140     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuHelpUpdate_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuHelpUpdate_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupAbout_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupAbout_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupAbout_Click")
    '</EhHeader>
100     Load frmAbout
104     frmAbout.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupAbout_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupAbout_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupExit_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupExit_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupExit_Click")
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupExit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupExit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupForum_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupForum_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupForum_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupForum_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupForum_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupHomePage_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupHomePage_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupHomePage_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupHomePage_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupHomePage_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupOpenCC_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupOpenCC_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupOpenCC_Click")
    '</EhHeader>
100     Me.WindowState = vbNormal
104     Me.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupOpenCC_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupOpenCC_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSysTrayPopupUpdate_Click()
    '<EhHeader>
    On Error GoTo mnuSysTrayPopupUpdate_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.mnuSysTrayPopupUpdate_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Retrieving Update Information") & "...", True
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
140     SetStatus WinUI.GetTranslatedText("Ready") & "..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

mnuSysTrayPopupUpdate_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.mnuSysTrayPopupUpdate_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub SysTray_LButtonDblClk()
    '<EhHeader>
    On Error GoTo SysTray_LButtonDblClk_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.SysTray_LButtonDblClk")
    '</EhHeader>
100     Me.WindowState = vbNormal
104     Me.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

SysTray_LButtonDblClk_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.SysTray_LButtonDblClk", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub SysTray_RButtonUp()
    '<EhHeader>
    On Error GoTo SysTray_RButtonUp_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.SysTray_RButtonUp")
    '</EhHeader>
100     SetForegroundWindow Me.hwnd
104     PopupMenu mnuSysTrayPopup, , , , mnuSysTrayPopupOpenCC
108     PostMessage Me.hwnd, 0&, 0&, 0&
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

SysTray_RButtonUp_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.SysTray_RButtonUp", Erl, False
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

    strSrvStatusCur = WinUI.Server.HTTP.Status
    lblSrvStatusCur.Font.Bold = False
    Select Case strSrvStatusCur
        Case "Stopped"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Stopped")
            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Stopped"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = True
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Running")
            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            cmdSrvStart.Enabled = False
            cmdSrvStop.Enabled = True
            cmdSrvRestart.Enabled = True
    End Select
End Sub


Private Function LoadConfigData() As Boolean
    '<EhHeader>
    On Error GoTo LoadConfigData_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.LoadConfigData")
    '</EhHeader>
    Dim strTemp As String
    Dim strResult As String
    Dim vItem As Variant
    
100     WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.LoadConfigData", "Loading Config Data"
104     SetStatus WinUI.GetTranslatedText("Loading Configuration Data") & "...", True
108     LoadConfigData = WinUI.Server.HTTP.Config.LoadData
    
        'Setup the form...
112     txtServerName.Text = WinUI.Server.HTTP.Config.ServerName
116     txtPort.Text = WinUI.Server.HTTP.Config.Port
120     txtWebroot.Text = WinUI.Server.HTTP.Config.WebRoot
124     txtMaxConnect.Text = WinUI.Server.HTTP.Config.MaxConnections
128     txtLogFile.Text = WinUI.Server.HTTP.Config.LogFile
132     txtConfigAdvIPBind.Text = WinUI.Server.HTTP.Config.ListeningAddress
136     txtAllowIndex.Text = WinUI.Server.HTTP.Config.AllowIndex
140     txtErrorPages.Text = WinUI.Server.HTTP.Config.ErrorPages
144     txtConfigBasicErrorLog.Text = WinUI.Server.HTTP.Config.ErrorLog
    
148     For Each vItem In WinUI.Server.HTTP.Config.Index
152         strTemp = strTemp & vItem.FileName & " "
        Next
156     txtIndexFiles.Text = Trim$(strTemp)
    
160     lstISAPI.Enabled = False
164     lstISAPI.Clear
168     For Each vItem In WinUI.Server.HTTP.Config.ISAPI
172         lstISAPI.AddItem vItem.Extension
176         lstISAPI.Enabled = True
        Next
    
180     lstvHosts.Enabled = False
184     lstvHosts.Clear
188     For Each vItem In WinUI.Server.HTTP.Config.VirtHost
192         lstvHosts.AddItem vItem.HostName
196         lstvHosts.Enabled = True
        Next
    
200     cmbViewLogFiles.Clear
204     If Dir$(WinUI.Server.HTTP.Config.LogFile) <> "" Then
208         cmbViewLogFiles.AddItem WinUI.Server.HTTP.Config.LogFile
        End If
212     If Dir$(WinUI.Server.HTTP.Config.ErrorLog) <> "" Then
216         cmbViewLogFiles.AddItem WinUI.Server.HTTP.Config.ErrorLog
        End If
220     For Each vItem In WinUI.Server.HTTP.Config.VirtHost
224         If Dir$(vItem.Log) <> "" Then
228             cmbViewLogFiles.AddItem vItem.Log
            End If
        Next
    
        'we now only check for updates every 24 hours, this could confuse some people.
        'but this should make loading faster.
232     SetStatus "Checking For Updates...", True
236     strResult = WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "LastUpdateCheck")
240     If strResult = "" Then
244         strResult = CDate(1.1)
        End If
248     If DateDiff("h", CDate(strResult), Now) >= 24 Then
252         WinUI.Update.Check
256         If WinUI.Update.IsAvailable = True Then
260             lblUpdateStatus.Caption = WinUI.GetTranslatedText("New Version Available")
            Else
264             lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
268             lblUpdateStatus.Font.Underline = False
272             lblUpdateStatus.ForeColor = vbButtonText
276             lblUpdateStatus.MousePointer = vbDefault
280             WinUI.Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "LastUpdateCheck", Now
            End If
        Else
284         lblUpdateStatus.Caption = WinUI.GetTranslatedText("No Updates Available")
288         lblUpdateStatus.Font.Underline = False
292         lblUpdateStatus.ForeColor = vbButtonText
296         lblUpdateStatus.MousePointer = vbDefault
        End If
    
300     UpdateStats
        
304     If WinUI.Registration.IsRegistered = True Then
308         SetStatus "Updating Registration..."
312         mnuHelpRegister.Enabled = False
316         WinUI.Registration.Renew
        End If
    
320     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

LoadConfigData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.LoadConfigData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Private Sub txtAllowIndex_Change()
    '<EhHeader>
    On Error GoTo txtAllowIndex_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtAllowIndex_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.AllowIndex <> IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false") Then
104         WinUI.Server.HTTP.Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtAllowIndex_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtAllowIndex_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtISAPIExt_Change()
    '<EhHeader>
    On Error GoTo txtISAPIExt_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtISAPIExt_Change")
    '</EhHeader>
100     If lstISAPI.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension <> txtISAPIExt.Text Then
108             WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension = txtISAPIExt.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtISAPIExt_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtISAPIExt_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtISAPIInterp_Change()
    '<EhHeader>
    On Error GoTo txtISAPIInterp_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtISAPIInterp_Change")
    '</EhHeader>
100     If lstISAPI.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter <> txtISAPIInterp.Text Then
108             WinUI.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter = txtISAPIInterp.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtISAPIInterp_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtISAPIInterp_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtConfigAdvIPBind_Change()
    '<EhHeader>
    On Error GoTo txtConfigAdvIPBind_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtConfigAdvIPBind_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.ListeningAddress = txtConfigAdvIPBind.Text Then
104         WinUI.Server.HTTP.Config.ListeningAddress = txtConfigAdvIPBind.Text
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtConfigAdvIPBind_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtConfigAdvIPBind_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtConfigBasicErrorLog_Change()
    '<EhHeader>
    On Error GoTo txtConfigBasicErrorLog_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtConfigBasicErrorLog_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.ErrorLog <> txtConfigBasicErrorLog.Text Then
104         WinUI.Server.HTTP.Config.ErrorLog = txtConfigBasicErrorLog.Text
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtConfigBasicErrorLog_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtConfigBasicErrorLog_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtErrorPages_Change()
    '<EhHeader>
    On Error GoTo txtErrorPages_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtErrorPages_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.ErrorPages <> txtErrorPages.Text Then
104         WinUI.Server.HTTP.Config.ErrorPages = txtErrorPages.Text
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtErrorPages_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtErrorPages_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtIndexFiles_Change()
    '<EhHeader>
    On Error GoTo txtIndexFiles_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtIndexFiles_Change")
    '</EhHeader>
    Dim strTmpArray() As String
    Dim lngRecCount As Long
    Dim i As Long
100     strTmpArray = Split(Trim$(txtIndexFiles.Text), " ")
104     If Not IsEmpty(strTmpArray) Then
108         WinUI.Server.HTTP.Config.Index.Clear
112         lngRecCount = UBound(strTmpArray)
116         For i = 0 To lngRecCount
120             WinUI.Server.HTTP.Config.Index.Add strTmpArray(i)
            Next
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtIndexFiles_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtIndexFiles_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo txtIndexFiles_KeyPress_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtIndexFiles_KeyPress")
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtIndexFiles_KeyPress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_KeyPress", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtIndexFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo txtIndexFiles_MouseUp_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtIndexFiles_MouseUp")
    '</EhHeader>
100     blnDirty = True
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtIndexFiles_MouseUp_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtIndexFiles_MouseUp", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtLogFile_Change()
    '<EhHeader>
    On Error GoTo txtLogFile_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtLogFile_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.LogFile <> Trim$(txtLogFile.Text) Then
104         WinUI.Server.HTTP.Config.LogFile = Trim$(txtLogFile.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtLogFile_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtLogFile_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtMaxConnect_Change()
    '<EhHeader>
    On Error GoTo txtMaxConnect_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtMaxConnect_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.MaxConnections <> Int(Val(txtMaxConnect.Text)) Then
104         WinUI.Server.HTTP.Config.MaxConnections = Int(Val(txtMaxConnect.Text))
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtMaxConnect_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtMaxConnect_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtPort_Change()
    '<EhHeader>
    On Error GoTo txtPort_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtPort_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.Port <> Int(Val(txtPort.Text)) Then
104         WinUI.Server.HTTP.Config.Port = Int(Val(txtPort.Text))
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtPort_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtPort_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtServerName_Change()
    '<EhHeader>
    On Error GoTo txtServerName_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtServerName_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.ServerName <> Trim$(txtServerName.Text) Then
104         WinUI.Server.HTTP.Config.ServerName = Trim$(txtServerName.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtServerName_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtServerName_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostDomain_Change()
    '<EhHeader>
    On Error GoTo txtvHostDomain_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtvHostDomain_Change")
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain <> txtvHostDomain.Text Then
108             WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain = txtvHostDomain.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtvHostDomain_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostDomain_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostLog_Change()
    '<EhHeader>
    On Error GoTo txtvHostLog_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtvHostLog_Change")
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log <> txtvHostLog.Text Then
108             WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log = txtvHostLog.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtvHostLog_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostLog_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostName_Change()
    '<EhHeader>
    On Error GoTo txtvHostName_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtvHostName_Change")
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName <> txtvHostName.Text Then
108             blnDirty = True
112             WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName = txtvHostName.Text
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtvHostName_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostName_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtvHostRoot_Change()
    '<EhHeader>
    On Error GoTo txtvHostRoot_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtvHostRoot_Change")
    '</EhHeader>
100     If lstvHosts.ListIndex <> -1 Then
104         If WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root <> txtvHostRoot.Text Then
108             WinUI.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root = txtvHostRoot.Text
112             blnDirty = True
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtvHostRoot_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtvHostRoot_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtWebroot_Change()
    '<EhHeader>
    On Error GoTo txtWebroot_Change_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.txtWebroot_Change")
    '</EhHeader>
100     If WinUI.Server.HTTP.Config.WebRoot <> Trim$(txtWebroot.Text) Then
104         WinUI.Server.HTTP.Config.WebRoot = Trim$(txtWebroot.Text)
108         blnDirty = True
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

txtWebroot_Change_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.txtWebroot_Change", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub vbaSideBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    '<EhHeader>
    On Error GoTo vbaSideBar_ItemClick_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.vbaSideBar_ItemClick")
    '</EhHeader>
100     WinUI.Util.StopWinUpdate Me.hwnd
104     Select Case itm.Key
            Case "status"
108             fraStatus.ZOrder 0
112         Case "basic"
116             fraConfigBasic.ZOrder 0
120         Case "advanced"
124             fraConfigAdv.ZOrder 0
128         Case "vhost"
132             fraConfigvHost.ZOrder 0
136         Case "isapi"
140             fraConfigISAPI.ZOrder 0
144         Case "logs"
148             fraLogs.ZOrder 0
        End Select
152     vbaSideBar.ZOrder 0
156     WinUI.Util.StopWinUpdate
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

vbaSideBar_ItemClick_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.vbaSideBar_ItemClick", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub UpdateStats()
    '<EhHeader>
    On Error GoTo UpdateStats_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.UpdateStats")
    '</EhHeader>
100     WinUI.Server.HTTP.Stats.Reload
104     lblStatsLastRestart.Caption = WinUI.GetTranslatedText("Last Restart") & ": " & WinUI.Server.HTTP.Stats.LastRestart
108     lblStatsRequestCount.Caption = WinUI.GetTranslatedText("Request Count") & ": " & WinUI.Server.HTTP.Stats.RequestCount
112     lblStatsBytesSent.Caption = WinUI.GetTranslatedText("Total Bytes Sent") & ": " & Format$(WinUI.Server.HTTP.Stats.TotalBytesSent, "###,###,###,###,##0")
116     lblCurVersion.Caption = WinUI.GetTranslatedText("Current Version") & ": " & WinUI.Version
120     lblUpdateVersion.Caption = WinUI.GetTranslatedText("Update Version") & ": " & IIf(WinUI.Update.Version <> "", WinUI.Update.Version, WinUI.Version)
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

UpdateStats_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.UpdateStats", Erl, False
    Resume Next
    '</EhFooter>
End Sub
