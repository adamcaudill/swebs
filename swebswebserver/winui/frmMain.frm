VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SWEBS Web Server - Control Center"
   ClientHeight    =   7305
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   12075
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfigvHost 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Width           =   6975
      Begin VB.ListBox lstvHosts 
         Height          =   5130
         ItemData        =   "frmMain.frx":0CCA
         Left            =   120
         List            =   "frmMain.frx":0CCC
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtvHostName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtvHostDomain 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtvHostRoot 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtvHostLog 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label lblvHostNew 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New"
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
         Left            =   2160
         MouseIcon       =   "frmMain.frx":0CCE
         MousePointer    =   99  'Custom
         TabIndex        =   107
         Top             =   5160
         Width           =   765
      End
      Begin VB.Label lblvHostRemove 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
         Enabled         =   0   'False
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
         Left            =   3270
         MouseIcon       =   "frmMain.frx":0E20
         MousePointer    =   99  'Custom
         TabIndex        =   106
         Top             =   5160
         Width           =   705
      End
      Begin VB.Label lblBrowsevHostRoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         Enabled         =   0   'False
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
         Left            =   6240
         MouseIcon       =   "frmMain.frx":0F72
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label lblBrowsevHostLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         Enabled         =   0   'False
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
         Left            =   6240
         MouseIcon       =   "frmMain.frx":10C4
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label lblvHostName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What is the name of this Virtual Host?"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblvHostDomain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What is it's domain name?"
         Height          =   255
         Left            =   2040
         TabIndex        =   28
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblvHostRoot 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This is the root directory where files are kept for this Virtual Host."
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblvHostLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Where do you want to keep the log file for this Virtual Host?"
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   2040
         Width           =   4335
      End
   End
   Begin VB.Frame fraConfigISAPI 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      Width           =   6975
      Begin VB.ListBox lstISAPI 
         Height          =   5130
         ItemData        =   "frmMain.frx":1216
         Left            =   120
         List            =   "frmMain.frx":121D
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtISAPIInterp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtISAPIExt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblBrowseISAPIInterp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         Enabled         =   0   'False
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
         Left            =   5880
         MouseIcon       =   "frmMain.frx":122B
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblISAPIRemove 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
         Enabled         =   0   'False
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
         Left            =   3255
         MouseIcon       =   "frmMain.frx":137D
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   5160
         Width           =   705
      End
      Begin VB.Label lblISAPINew 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New"
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
         Left            =   2145
         MouseIcon       =   "frmMain.frx":14CF
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   5160
         Width           =   765
      End
      Begin VB.Label lblISAPIInterp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Where is the ISAPI Plugin?"
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblISAPIExt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What is the extension that is mapped to this interpreter."
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   1080
         Width           =   4815
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   6975
      Begin VB.PictureBox picButton 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   10
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   6615
         TabIndex        =   54
         Top             =   120
         Width           =   6615
         Begin VB.Frame fraBasicStats 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Basic Stats:"
            Height          =   1095
            Left            =   0
            TabIndex        =   62
            Top             =   1200
            Width           =   3135
            Begin VB.Label lblStatsBytesSent 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Total Bytes Sent: 000,000,000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label lblStatsRequestCount 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Request Count: 000,000,000"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label lblStatsLastRestart 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Restart: 00/00/0000 00:00:00PM"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraUpdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Update Status:"
            Height          =   1095
            Left            =   3240
            TabIndex        =   55
            Top             =   0
            Width           =   3255
            Begin VB.Label lblUpdateStatus 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               MouseIcon       =   "frmMain.frx":1621
               MousePointer    =   99  'Custom
               TabIndex        =   58
               ToolTipText     =   "Click here for details."
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label lblUpdateVersion 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Update Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label lblCurVersion 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Current Version: 0.00.0000"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame fraSrvStatus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Current Service Status:"
            Height          =   1095
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   3135
            Begin VB.Label lblSrvRestart 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Restart"
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
               Left            =   2145
               MouseIcon       =   "frmMain.frx":1773
               MousePointer    =   99  'Custom
               TabIndex        =   94
               Top             =   720
               Width           =   645
            End
            Begin VB.Label lblSrvStop 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Stop"
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
               Left            =   1305
               MouseIcon       =   "frmMain.frx":18C5
               MousePointer    =   99  'Custom
               TabIndex        =   93
               Top             =   720
               Width           =   405
            End
            Begin VB.Label lblSrvStart 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start"
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
               Left            =   330
               MouseIcon       =   "frmMain.frx":1A17
               MousePointer    =   99  'Custom
               TabIndex        =   92
               Top             =   720
               Width           =   435
            End
            Begin VB.Label lblSrvStatusCur 
               BackColor       =   &H00FFFFFF&
               Caption         =   "<current-status>"
               Height          =   255
               Left            =   720
               TabIndex        =   61
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label lblSrvStatus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Status: "
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   615
            End
         End
      End
   End
   Begin VB.Frame fraLogs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   36
      Top             =   840
      Width           =   6975
      Begin RichTextLib.RichTextBox rtfViewLogFiles 
         Height          =   5055
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8916
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":1B69
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Height          =   315
         ItemData        =   "frmMain.frx":1BEB
         Left            =   120
         List            =   "frmMain.frx":1BED
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Frame fraConfigBasic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   6975
      Begin VB.TextBox txtConfigBasicErrorLog 
         Height          =   285
         Left            =   240
         TabIndex        =   69
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtWebroot 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox txtLogFile 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   3360
         Width           =   5895
      End
      Begin VB.Label lblBrowseLogFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
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
         Left            =   6240
         MouseIcon       =   "frmMain.frx":1BEF
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label lblBrowseRoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
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
         Left            =   6240
         MouseIcon       =   "frmMain.frx":1D41
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label lblBrowseErrorLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
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
         Left            =   6240
         MouseIcon       =   "frmMain.frx":1E93
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label lblConfigBasicErrorLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Where do you want to store the server error log?"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label lblLogFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here."
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   6135
      End
      Begin VB.Label lblServerName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What is the name of your server?"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What port do you want to use? (Default is 80)"
         Height          =   495
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblWebroot 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmMain.frx":1FE5
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6135
      End
   End
   Begin VB.Frame fraConfigAdv 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   11
      Top             =   840
      Width           =   6975
      Begin VB.TextBox txtConfigAdvIPBind 
         Height          =   285
         Left            =   240
         TabIndex        =   67
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtMaxConnect 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtIndexFiles 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   5655
      End
      Begin VB.TextBox txtAllowIndex 
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtErrorPages 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label lblBrowseErrorPages 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
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
         Left            =   6000
         MouseIcon       =   "frmMain.frx":2089
         MousePointer    =   99  'Custom
         TabIndex        =   95
         Top             =   3240
         Width           =   660
      End
      Begin VB.Label lblConfigAdvIPBind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What IP should the server listen to? (Default: Leave blank for all available)"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblMaxConnect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "What is the maximum number of connections that your server can handle at any one time."
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblIndexFiles 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmMain.frx":21DB
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label lblAllowIndex 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display file list if no index is found?"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblErrorPages 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Where is the location of the folder which stores pages to be used when the server receives an error."
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   5895
      End
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   120
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrStats 
      Interval        =   60000
      Left            =   600
      Top             =   6720
   End
   Begin VB.Timer tmrStatus 
      Interval        =   1000
      Left            =   1080
      Top             =   6720
   End
   Begin VB.Frame fraNewvHost 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   2520
      TabIndex        =   38
      Top             =   840
      Width           =   6855
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2280
         ScaleHeight     =   375
         ScaleWidth      =   2175
         TabIndex        =   51
         Top             =   3240
         Width           =   2175
         Begin VB.CommandButton cmdNewvHostOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdNewvHostCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1200
            TabIndex        =   52
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdBrowseNewvHostRoot 
         Caption         =   "..."
         Height          =   255
         Left            =   5880
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   2160
         Width           =   255
         Begin VB.CommandButton cmdBrowseNewvHostLogs 
            Caption         =   "..."
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.TextBox txtNewvHostLogs 
         Height          =   285
         Left            =   600
         TabIndex        =   47
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostRoot 
         Height          =   285
         Left            =   600
         TabIndex        =   45
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtNewvHostDomain 
         Height          =   285
         Left            =   600
         TabIndex        =   44
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtNewvHostName 
         Height          =   285
         Left            =   600
         TabIndex        =   41
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblNewvHostLogs 
         Caption         =   "Where do you want to keep the log for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   2520
         Width           =   5295
      End
      Begin VB.Label lblNewvHostDomain 
         Caption         =   "What is the domain for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblNewvHostRoot 
         Caption         =   "Where is the root folder for this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label lblNewvHostName 
         Caption         =   "What is the name of this Virtual Host?"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label lblNewvHostTitle 
         Caption         =   "Add a new Virtual Host:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label lblMoreInfoData 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   9840
      TabIndex        =   103
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Line lneMain 
      Index           =   6
      X1              =   9720
      X2              =   9720
      Y1              =   3840
      Y2              =   6840
   End
   Begin VB.Line lneMain 
      Index           =   11
      X1              =   11880
      X2              =   9720
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lneMain 
      Index           =   12
      X1              =   11880
      X2              =   9720
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line lneMain 
      Index           =   14
      X1              =   11880
      X2              =   11880
      Y1              =   3840
      Y2              =   6840
   End
   Begin VB.Line lneMain 
      Index           =   13
      X1              =   11880
      X2              =   9720
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblMoreInfoTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9840
      TabIndex        =   102
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label lblLogo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SWEBS Web Suite"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8430
      TabIndex        =   91
      Top             =   120
      Width           =   2805
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   11400
      Picture         =   "frmMain.frx":2289
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":2F53
      MousePointer    =   99  'Custom
      TabIndex        =   90
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":30A5
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label lblUpdates 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check For Update"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":31F7
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label lblForum 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "SWEBS Forum"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":3349
      MousePointer    =   99  'Custom
      TabIndex        =   87
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "SWEBS Web Site"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":349B
      MousePointer    =   99  'Custom
      TabIndex        =   86
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblHTTPExportSettings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export Settings"
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
      Left            =   9960
      MouseIcon       =   "frmMain.frx":35ED
      MousePointer    =   99  'Custom
      TabIndex        =   85
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label lblTools 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tools:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9840
      TabIndex        =   84
      Top             =   1080
      Width           =   510
   End
   Begin VB.Line lneMain 
      Index           =   7
      X1              =   11880
      X2              =   9720
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line lneMain 
      Index           =   10
      X1              =   11880
      X2              =   11880
      Y1              =   960
      Y2              =   3720
   End
   Begin VB.Line lneMain 
      Index           =   8
      X1              =   11880
      X2              =   9720
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line lneMain 
      Index           =   9
      X1              =   11880
      X2              =   9720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line lneMain 
      Index           =   5
      X1              =   9720
      X2              =   9720
      Y1              =   960
      Y2              =   3720
   End
   Begin VB.Line lneMain 
      Index           =   0
      X1              =   12165
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Your Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   120
      TabIndex        =   83
      Top             =   120
      Width           =   3750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00804008&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   12165
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
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
      Left            =   11535
      MouseIcon       =   "frmMain.frx":373F
      MousePointer    =   99  'Custom
      TabIndex        =   82
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label lblApply 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Apply"
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
      Left            =   10560
      MouseIcon       =   "frmMain.frx":3891
      MousePointer    =   99  'Custom
      TabIndex        =   81
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
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
      Left            =   9720
      MouseIcon       =   "frmMain.frx":39E3
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   6960
      Width           =   285
   End
   Begin VB.Label lblViewLogs 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Logs"
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
      Left            =   1290
      MouseIcon       =   "frmMain.frx":3B35
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   3480
      Width           =   885
   End
   Begin VB.Line lneMain 
      Index           =   4
      X1              =   120
      X2              =   2160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblSystemLogs 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Logs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   78
      Top             =   3120
      Width           =   1110
   End
   Begin VB.Label lblHTTPConfigISAPI 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ISAPI Plugins"
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
      Left            =   1050
      MouseIcon       =   "frmMain.frx":3C87
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   2760
      Width           =   1125
   End
   Begin VB.Label lblHTTPConfigVirtHost 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Virtual Hosts"
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
      Left            =   1080
      MouseIcon       =   "frmMain.frx":3DD9
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblHTTPConfigAdv 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advanced"
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
      Left            =   1335
      MouseIcon       =   "frmMain.frx":3F2B
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label lblHTTPConfigBasic 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Basic"
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
      Left            =   1695
      MouseIcon       =   "frmMain.frx":407D
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   2040
      Width           =   480
   End
   Begin VB.Line lneMain 
      Index           =   3
      X1              =   120
      X2              =   2160
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblServerConfig 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server Configuration:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   73
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label lblCurrentStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Status"
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
      Left            =   960
      MouseIcon       =   "frmMain.frx":41CF
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Line lneMain 
      Index           =   2
      X1              =   2160
      X2              =   120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblSystemStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   71
      Top             =   960
      Width           =   1290
   End
   Begin VB.Line lneMain 
      Index           =   1
      X1              =   2400
      X2              =   2400
      Y1              =   720
      Y2              =   7320
   End
   Begin VB.Label lblAppStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   6840
      Width           =   6735
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

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.Form_Load")
    '</EhHeader>
    Dim RetVal As Long
    
        'setup the translated strings...
100     SetStatus "Loading Translated Strings..."
    
104     lblHTTPExportSettings.Caption = WinUI.GetTranslatedText("Export Setings")
108     lblHomePage.Caption = WinUI.GetTranslatedText("SWEBS Web Site")
112     lblForum.Caption = WinUI.GetTranslatedText("SWEBS Forum")
116     lblUpdates.Caption = WinUI.GetTranslatedText("Check For Update")
120     lblRegister.Caption = WinUI.GetTranslatedText("Register")
124     lblAbout.Caption = WinUI.GetTranslatedText("&About")
128     lblOK.Caption = WinUI.GetTranslatedText("&OK")
132     lblApply.Caption = WinUI.GetTranslatedText("&Apply")
136     lblExit.Caption = WinUI.GetTranslatedText("E&xit")
140     fraSrvStatus.Caption = WinUI.GetTranslatedText("Current Service Status:")
144     lblSrvStatus.Caption = WinUI.GetTranslatedText("Status:")
148     lblSrvStart.Caption = WinUI.GetTranslatedText("S&tart")
152     lblSrvStop.Caption = WinUI.GetTranslatedText("St&op")
156     lblSrvRestart.Caption = WinUI.GetTranslatedText("R&estart")
160     fraUpdate.Caption = WinUI.GetTranslatedText("Update Status:")
164     fraBasicStats.Caption = WinUI.GetTranslatedText("Basic Stats:")
168     lblMaxConnect.Caption = WinUI.GetTranslatedText("What is the maximum number of connections that your server can handle at any one time.")
172     lblAllowIndex.Caption = WinUI.GetTranslatedText("Display file list if no index is found?")
176     lblIndexFiles.Caption = WinUI.GetTranslatedText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
180     lblErrorPages.Caption = WinUI.GetTranslatedText("Where is the location of the folder which stores pages to be used when the server receives an error.")
184     lblServerName.Caption = WinUI.GetTranslatedText("What is the name of your server?")
188     lblPort.Caption = WinUI.GetTranslatedText("What port do you want to use? (Default is 80)")
192     lblWebroot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
196     lblLogFile.Caption = WinUI.GetTranslatedText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
200     lblISAPIInterp.Caption = WinUI.GetTranslatedText("Where is the ISAPI Plugin?")
204     lblISAPIExt.Caption = WinUI.GetTranslatedText("What is the extension that is mapped to this interpreter.")
208     lblISAPINew.Caption = WinUI.GetTranslatedText("Add New...")
212     lblISAPIRemove.Caption = WinUI.GetTranslatedText("Remove...")
216     lblvHostNew.Caption = WinUI.GetTranslatedText("Add New...")
220     lblvHostRemove.Caption = WinUI.GetTranslatedText("Remove...")
224     lblvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
228     lblvHostDomain.Caption = WinUI.GetTranslatedText("What is it's domain name?")
232     lblvHostRoot.Caption = WinUI.GetTranslatedText("This is the root directory where files are kept for this Virtual Host.")
236     lblvHostLog.Caption = WinUI.GetTranslatedText("Where do you want to keep the log file for this Virtual Host?")
240     lblNewvHostTitle.Caption = WinUI.GetTranslatedText("Add a new Virtual Host:")
244     lblNewvHostName.Caption = WinUI.GetTranslatedText("What is the name of this Virtual Host?")
248     lblNewvHostDomain.Caption = WinUI.GetTranslatedText("What is the domain for this Virtual Host?")
252     lblNewvHostRoot.Caption = WinUI.GetTranslatedText("Where is the root folder for this Virtual Host?")
256     lblNewvHostLogs.Caption = WinUI.GetTranslatedText("Where do you want to keep the log for this Virtual Host?")
260     cmdNewvHostOK.Caption = WinUI.GetTranslatedText("&OK")
264     cmdNewvHostCancel.Caption = WinUI.GetTranslatedText("&Cancel")
268     lblConfigAdvIPBind.Caption = WinUI.GetTranslatedText("What IP should the server listen to? (Default: Leave blank for all available)")
272     lblConfigBasicErrorLog.Caption = WinUI.GetTranslatedText("Where do you want to store the server error log?")
    
276     If LoadConfigData = False Then
280         RetVal = MsgBox(WinUI.GetTranslatedText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
284         Select Case RetVal
                Case vbAbort
288                 End
292             Case vbRetry
296                 If LoadConfigData = False Then
300                     MsgBox WinUI.GetTranslatedText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
304                     End
                    End If
308             Case vbIgnore
312                 MsgBox WinUI.GetTranslatedText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
            End Select
        End If
    
316     Set SysTray = New cSysTray
320     Set SysTray.SourceWindow = Me
324     SysTray.IconInSysTray
328     SysTray.ToolTip = WinUI.GetTranslatedText("SWEBS Web Server") & " " & WinUI.Version
332     SysTray.Icon = Me.Icon

336     fraStatus.ZOrder 0
340     tmrStatus_Timer
344     SetStatus WinUI.GetTranslatedText("Ready") & "..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo Form_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.Form_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = ""
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.Form_MouseMove", Erl, False
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

Private Sub lblAbout_Click()
    '<EhHeader>
    On Error GoTo lblAbout_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblAbout_Click")
    '</EhHeader>
100     Load frmAbout
104     frmAbout.Show
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblAbout_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblAbout_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblAbout_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblAbout_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "Displays more information about this application and who wrote it. This also contains useful copyright and other legal information."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblAbout_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblAbout_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblApply_Click()
    '<EhHeader>
    On Error GoTo lblApply_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblApply_Click")
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

lblApply_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblApply_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowseErrorLog_Click()
    '<EhHeader>
    On Error GoTo lblBrowseErrorLog_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowseErrorLog_Click")
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

lblBrowseErrorLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowseErrorLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowseErrorPages_Click()
    '<EhHeader>
    On Error GoTo lblBrowseErrorPages_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowseErrorPages_Click")
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

lblBrowseErrorPages_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowseErrorPages_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowseISAPIInterp_Click()
    '<EhHeader>
    On Error GoTo lblBrowseISAPIInterp_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowseISAPIInterp_Click")
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

lblBrowseISAPIInterp_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowseISAPIInterp_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowseLogFile_Click()
    '<EhHeader>
    On Error GoTo lblBrowseLogFile_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowseLogFile_Click")
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

lblBrowseLogFile_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowseLogFile_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowseRoot_Click()
    '<EhHeader>
    On Error GoTo lblBrowseRoot_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowseRoot_Click")
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

lblBrowseRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowseRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowsevHostLog_Click()
    '<EhHeader>
    On Error GoTo lblBrowsevHostLog_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowsevHostLog_Click")
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

lblBrowsevHostLog_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowsevHostLog_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblBrowsevHostRoot_Click()
    '<EhHeader>
    On Error GoTo lblBrowsevHostRoot_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblBrowsevHostRoot_Click")
    '</EhHeader>
    Dim strPath As String

100     strPath = WinUI.Util.BrowseForFolder(, True, WinUI.Server.HTTP.Config.VirtHost((lstvHosts.ListIndex + 1)).Root)
104     If strPath <> "" Then
108         txtvHostRoot.Text = strPath
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblBrowsevHostRoot_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblBrowsevHostRoot_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblCurrentStatus_Click()
    '<EhHeader>
    On Error GoTo lblCurrentStatus_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblCurrentStatus_Click")
    '</EhHeader>
100     fraStatus.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblCurrentStatus_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblCurrentStatus_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblCurrentStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblCurrentStatus_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblCurrentStatus_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "This displays the most important status information about this application and your servers."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblCurrentStatus_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblCurrentStatus_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblExit_Click()
    '<EhHeader>
    On Error GoTo lblExit_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblExit_Click")
    '</EhHeader>
100     Unload Me
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblExit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblExit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblForum_Click()
    '<EhHeader>
    On Error GoTo lblForum_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblForum_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblForum_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblForum_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHomePage_Click()
    '<EhHeader>
    On Error GoTo lblHomePage_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHomePage_Click")
    '</EhHeader>
100     WinUI.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHomePage_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHomePage_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigAdv_Click()
    '<EhHeader>
    On Error GoTo lblHTTPConfigAdv_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigAdv_Click")
    '</EhHeader>
100     fraConfigAdv.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigAdv_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigAdv_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigAdv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblHTTPConfigAdv_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigAdv_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "From here you can adjust the servers advanced settings, these should options should be used with care."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigAdv_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigAdv_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigBasic_Click()
    '<EhHeader>
    On Error GoTo lblHTTPConfigBasic_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigBasic_Click")
    '</EhHeader>
100     fraConfigBasic.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigBasic_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigBasic_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblHTTPConfigBasic_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigBasic_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "From here you can change your servers basic settings, such as it's port, logging information and other simple items."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigBasic_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigBasic_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigISAPI_Click()
    '<EhHeader>
    On Error GoTo lblHTTPConfigISAPI_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigISAPI_Click")
    '</EhHeader>
100     fraConfigISAPI.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigISAPI_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigISAPI_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigISAPI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblHTTPConfigISAPI_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigISAPI_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "From here you can adjust the setiings of your ISAPI plugins, such as PHP and other scripting languages."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigISAPI_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigISAPI_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigVirtHost_Click()
    '<EhHeader>
    On Error GoTo lblHTTPConfigVirtHost_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigVirtHost_Click")
    '</EhHeader>
100     fraConfigvHost.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigVirtHost_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigVirtHost_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPConfigVirtHost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblHTTPConfigVirtHost_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPConfigVirtHost_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "This allows you to host several domains from one server. This fature is useful for hosting many websites at the same time."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblHTTPConfigVirtHost_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPConfigVirtHost_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblHTTPExportSettings_Click()
        'this needs some kind of error control, file checks, etc..
    '<EhHeader>
    On Error GoTo lblHTTPExportSettings_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblHTTPExportSettings_Click")
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

lblHTTPExportSettings_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblHTTPExportSettings_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblISAPINew_Click()
    '<EhHeader>
    On Error GoTo lblISAPINew_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblISAPINew_Click")
    '</EhHeader>
    Dim vItem As Variant
    Dim i As Long

100     Load frmNewISAPI
104     frmNewISAPI.Show vbModal
108     If WinUI.Server.HTTP.Config.ISAPI.Count > 0 Then
112         lstISAPI.Clear
116         For Each vItem In WinUI.Server.HTTP.Config.ISAPI
120             lstISAPI.AddItem vItem.Extension
124             lstISAPI.Enabled = True
            Next
        Else
128         lstISAPI.Enabled = False
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblISAPINew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblISAPINew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblISAPIRemove_Click()
    '<EhHeader>
    On Error GoTo lblISAPIRemove_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblISAPIRemove_Click")
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
144                 lblBrowseISAPIInterp.Enabled = False
148                 lblISAPIRemove.Enabled = False
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

lblISAPIRemove_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblISAPIRemove_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblOK_Click()
    '<EhHeader>
    On Error GoTo lblOK_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblOK_Click")
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

lblOK_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblRegister_Click()
    '<EhHeader>
    On Error GoTo lblRegister_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblRegister_Click")
    '</EhHeader>
100     WinUI.Registration.Start
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblRegister_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblRegister_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblSrvRestart_Click()
    '<EhHeader>
    On Error GoTo lblSrvRestart_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblSrvRestart_Click")
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

lblSrvRestart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblSrvRestart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblSrvStart_Click()
    '<EhHeader>
    On Error GoTo lblSrvStart_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblSrvStart_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Starting Service") & "...", True
104     WinUI.Server.HTTP.StartServer
108     UpdateStats
112     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblSrvStart_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblSrvStart_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblSrvStop_Click()
    '<EhHeader>
    On Error GoTo lblSrvStop_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblSrvStop_Click")
    '</EhHeader>
100     SetStatus WinUI.GetTranslatedText("Stopping Service") & "...", True
104     WinUI.Server.HTTP.StopServer
108     SetStatus "Ready..."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblSrvStop_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblSrvStop_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblUpdates_Click()
    '<EhHeader>
    On Error GoTo lblUpdates_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblUpdates_Click")
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

lblUpdates_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblUpdates_Click", Erl, False
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

Private Sub lblvHostNew_Click()
    '<EhHeader>
    On Error GoTo lblvHostNew_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblvHostNew_Click")
    '</EhHeader>
100     fraNewvHost.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblvHostNew_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblvHostNew_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblvHostRemove_Click()
    '<EhHeader>
    On Error GoTo lblvHostRemove_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblvHostRemove_Click")
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
156                 lblBrowsevHostRoot.Enabled = False
160                 lblBrowsevHostLog.Enabled = False
164                 lblvHostRemove.Enabled = False
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

lblvHostRemove_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblvHostRemove_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblViewLogs_Click()
    '<EhHeader>
    On Error GoTo lblViewLogs_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblViewLogs_Click")
    '</EhHeader>
100     fraLogs.ZOrder 0
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblViewLogs_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblViewLogs_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lblViewLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo lblViewLogs_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lblViewLogs_MouseMove")
    '</EhHeader>
100     lblMoreInfoData.Caption = "Displays the log files for your servers. From here you can view any events, errors and a log of all connections."
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

lblViewLogs_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmMain.lblViewLogs_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstISAPI_Click()
    '<EhHeader>
    On Error GoTo lstISAPI_Click_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmMain.lstISAPI_Click")
    '</EhHeader>
100     lblBrowseISAPIInterp.Enabled = True
104     lblISAPIRemove.Enabled = True
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
100     lblBrowsevHostRoot.Enabled = True
104     lblBrowsevHostLog.Enabled = True
108     lblvHostRemove.Enabled = True
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
            lblSrvStart.Enabled = True
            lblSrvStop.Enabled = False
            lblSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = WinUI.GetTranslatedText("Running")
            WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmMain.tmrStatus_Timer", "Service Status: Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            lblSrvStart.Enabled = False
            lblSrvStop.Enabled = True
            lblSrvRestart.Enabled = True
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
312         lblRegister.Enabled = False
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
