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
   Begin VB.Frame fraLogs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   2520
      TabIndex        =   36
      Top             =   840
      Width           =   6975
      Begin RichTextLib.RichTextBox rtfViewLogFiles 
         Height          =   5415
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   9551
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":0CCA
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":0D4C
         Left            =   120
         List            =   "frmMain.frx":0D4E
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Frame fraConfigvHost 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Width           =   6975
      Begin VB.ListBox lstvHosts 
         Height          =   5130
         ItemData        =   "frmMain.frx":0D50
         Left            =   120
         List            =   "frmMain.frx":0D52
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
         MouseIcon       =   "frmMain.frx":0D54
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
         MouseIcon       =   "frmMain.frx":0EA6
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
         MouseIcon       =   "frmMain.frx":0FF8
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
         MouseIcon       =   "frmMain.frx":114A
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
      Height          =   5895
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      Width           =   6975
      Begin VB.ListBox lstISAPI 
         Height          =   5130
         ItemData        =   "frmMain.frx":129C
         Left            =   120
         List            =   "frmMain.frx":12A3
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
         MouseIcon       =   "frmMain.frx":12B1
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
         MouseIcon       =   "frmMain.frx":1403
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
         MouseIcon       =   "frmMain.frx":1555
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
      Height          =   5895
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
               MouseIcon       =   "frmMain.frx":16A7
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
               MouseIcon       =   "frmMain.frx":17F9
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
               MouseIcon       =   "frmMain.frx":194B
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
               MouseIcon       =   "frmMain.frx":1A9D
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
   Begin VB.Frame fraConfigBasic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5895
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
      Height          =   5895
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
      Height          =   5895
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
   Begin VB.Shape shpMain 
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
      Top             =   6960
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
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Core - Custom
'***************************************************************************
'
' SWEBS/Core
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

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private WithEvents SysTray As cSysTray
Attribute SysTray.VB_VarHelpID = -1

Dim blnDirty As Boolean 'if true then assume that some bit of data has changed

Private Sub cmbViewLogFiles_Click()
Dim strLog As String
    
    SetStatus Translator.GetText("Loading Log File") & "...", True
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
        MsgBox Translator.GetText("File not found, it may not have been created yet."), vbExclamation + vbOKOnly + vbApplicationModal
    End If
    SetStatus "Ready..."
End Sub

Private Sub cmdBrowseNewvHostLogs_Click()
    blnDirty = True
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
    dlgMain.InitDir = Core.Path
    dlgMain.ShowSave
    txtvHostLog.Text = dlgMain.FileName
End Sub

Private Sub cmdBrowseNewvHostRoot_Click()
Dim strPath As String
    strPath = Util.BrowseForFolder(, True, Core.Server.HTTP.Config.WebRoot)
    If strPath <> "" Then
        txtNewvHostRoot.Text = strPath
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
Dim vItem As Variant
Dim i As Long

    If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
        blnDirty = True
        Core.Server.HTTP.Config.VirtHost.Add txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text, txtNewvHostName.Text
        lstvHosts.Clear
        If Core.Server.HTTP.Config.VirtHost.Count > 0 Then
            For Each vItem In Core.Server.HTTP.Config.VirtHost
                lstvHosts.AddItem vItem.HostName
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
        MsgBox Translator.GetText("Please fill all fields.")
    End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long
    
    'setup the translated strings...
    SetStatus "Loading Translated Strings..."
    
    lblHTTPExportSettings.Caption = Translator.GetText("Export Setings")
    lblHomePage.Caption = Translator.GetText("SWEBS Web Site")
    lblForum.Caption = Translator.GetText("SWEBS Forum")
    lblUpdates.Caption = Translator.GetText("Check For Update")
    lblRegister.Caption = Translator.GetText("Register")
    lblAbout.Caption = Translator.GetText("&About")
    lblOK.Caption = Translator.GetText("&OK")
    lblApply.Caption = Translator.GetText("&Apply")
    lblExit.Caption = Translator.GetText("E&xit")
    fraSrvStatus.Caption = Translator.GetText("Current Service Status:")
    lblSrvStatus.Caption = Translator.GetText("Status:")
    lblSrvStart.Caption = Translator.GetText("S&tart")
    lblSrvStop.Caption = Translator.GetText("St&op")
    lblSrvRestart.Caption = Translator.GetText("R&estart")
    fraUpdate.Caption = Translator.GetText("Update Status:")
    fraBasicStats.Caption = Translator.GetText("Basic Stats:")
    lblMaxConnect.Caption = Translator.GetText("What is the maximum number of connections that your server can handle at any one time.")
    lblAllowIndex.Caption = Translator.GetText("Display file list if no index is found?")
    lblIndexFiles.Caption = Translator.GetText("Files that will be used as indexes when a request is made to a folder. If a client requests a folder, the server will look inside that folder for a file with these names.")
    lblErrorPages.Caption = Translator.GetText("Where is the location of the folder which stores pages to be used when the server receives an error.")
    lblServerName.Caption = Translator.GetText("What is the name of your server?")
    lblPort.Caption = Translator.GetText("What port do you want to use? (Default is 80)")
    lblWebroot.Caption = Translator.GetText("This is the root directory where files are kept. Any files/folders in this folder will be publicly visible on the internet. Be careful when changing this entry.")
    lblLogFile.Caption = Translator.GetText("This is the file where all logging is written to. Any requests that DO NOT use a virtual server will be logged here.")
    lblISAPIInterp.Caption = Translator.GetText("Where is the ISAPI Plugin?")
    lblISAPIExt.Caption = Translator.GetText("What is the extension that is mapped to this interpreter.")
    lblISAPINew.Caption = Translator.GetText("Add New...")
    lblISAPIRemove.Caption = Translator.GetText("Remove...")
    lblvHostNew.Caption = Translator.GetText("Add New...")
    lblvHostRemove.Caption = Translator.GetText("Remove...")
    lblvHostName.Caption = Translator.GetText("What is the name of this Virtual Host?")
    lblvHostDomain.Caption = Translator.GetText("What is it's domain name?")
    lblvHostRoot.Caption = Translator.GetText("This is the root directory where files are kept for this Virtual Host.")
    lblvHostLog.Caption = Translator.GetText("Where do you want to keep the log file for this Virtual Host?")
    lblNewvHostTitle.Caption = Translator.GetText("Add a new Virtual Host:")
    lblNewvHostName.Caption = Translator.GetText("What is the name of this Virtual Host?")
    lblNewvHostDomain.Caption = Translator.GetText("What is the domain for this Virtual Host?")
    lblNewvHostRoot.Caption = Translator.GetText("Where is the root folder for this Virtual Host?")
    lblNewvHostLogs.Caption = Translator.GetText("Where do you want to keep the log for this Virtual Host?")
    cmdNewvHostOK.Caption = Translator.GetText("&OK")
    cmdNewvHostCancel.Caption = Translator.GetText("&Cancel")
    lblConfigAdvIPBind.Caption = Translator.GetText("What IP should the server listen to? (Default: Leave blank for all available)")
    lblConfigBasicErrorLog.Caption = Translator.GetText("Where do you want to store the server error log?")
    
    If LoadConfigData = False Then
        RetVal = MsgBox(Translator.GetText("There was an error while loading your configuration data.\r\rPress 'Abort' to give up and exit, 'Retry' to try to load the data again," & vbCrLf & "or 'Ignore' to continue."), vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                End
            Case vbRetry
                If LoadConfigData = False Then
                    MsgBox Translator.GetText("A second attempt to load your configuration data failed. Aborting.\r\rThis application will now close."), vbApplicationModal + vbCritical
                    End
                End If
            Case vbIgnore
                MsgBox Translator.GetText("NOTICE: You have chosen to proceed after a data error,\rthis application may not function properly or you may loose data."), vbInformation
        End Select
    End If
    
    Set SysTray = New cSysTray
    Set SysTray.SourceWindow = Me
    SysTray.IconInSysTray
    SysTray.ToolTip = Translator.GetText("SWEBS Web Server") & " " & Core.Version
    SysTray.Icon = Me.Icon

    fraStatus.ZOrder 0
    tmrStatus_Timer
    SetStatus Translator.GetText("Ready") & "..."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lngRetVal As Long

    If UnloadMode <> vbFormControlMenu Then
        If blnDirty = True Then
            lngRetVal = MsgBox(Translator.GetText("Do you want to save your settings before closing?"), vbYesNo + vbQuestion + vbApplicationModal)
            If lngRetVal = vbYes Then
                If Core.Server.HTTP.Config.Save(Core.Server.HTTP.Config.File) = False Then
                    MsgBox Translator.GetText("Data was not saved, no idea why...")
                    Cancel = True
                End If
            End If
        End If
    Else
        Cancel = True
        Me.WindowState = vbMinimized
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

    Me.Hide
    PostMessage Me.hWnd, 0&, 0&, 0&
    DoEvents
    SysTray.RemoveFromSysTray
    Set SysTray = Nothing
    DoEvents
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next
    Util.LoadUser32 False
    Set Core = Nothing
    SetExceptionFilter False
    End
End Sub

Private Sub lblAbout_Click()
    Me.MousePointer = 11
    DoEvents
    Load frmAbout
    Me.MousePointer = 0
    frmAbout.Show
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "Displays more information about this application and who wrote it. This also contains useful copyright and other legal information."
End Sub

Private Sub lblApply_Click()
    If Core.Server.HTTP.Config.Save(Core.Server.HTTP.Config.File) = False Then
        MsgBox Translator.GetText("Data was not saved, no idea why...")
    Else
        blnDirty = False
        MsgBox Translator.GetText("You data has been saved.\r\rYou will need to restart the SWEBS Service before these setting will take effect."), vbOKOnly + vbInformation
    End If
End Sub

Private Sub lblBrowseErrorLog_Click()
Dim strDefaultFile As String

    blnDirty = True
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
    strDefaultFile = Mid$(Core.Server.HTTP.Config.ErrorLog, (InStrRev(Core.Server.HTTP.Config.ErrorLog, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Core.Path
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtvHostLog.Text = dlgMain.FileName
    End If
End Sub

Private Sub lblBrowseErrorPages_Click()
Dim strPath As String
    blnDirty = True
    strPath = Util.BrowseForFolder(, True, Core.Server.HTTP.Config.ErrorPages)
    If strPath <> "" Then
        txtErrorPages.Text = strPath
    End If
End Sub

Private Sub lblBrowseISAPIInterp_Click()
Dim strDefaultFile As String
    blnDirty = True
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("ISAPI Plugin Files (*.dll)|*.dll|All Files (*.*)|*.*")
    strDefaultFile = Mid$(Core.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, (InStrRev(Core.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Core.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, 1, (Len(Core.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter) - InStrRev(Core.Server.HTTP.Config.ISAPI(lstISAPI.ListIndex + 1).Interpreter, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtISAPIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub lblBrowseLogFile_Click()
Dim strDefaultFile As String

    blnDirty = True
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
    strDefaultFile = Mid$(Core.Server.HTTP.Config.LogFile, (InStrRev(Core.Server.HTTP.Config.LogFile, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Core.Server.HTTP.Config.LogFile, 1, (Len(Core.Server.HTTP.Config.LogFile) - InStrRev(Core.Server.HTTP.Config.LogFile, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtLogFile.Text = dlgMain.FileName
    End If
End Sub

Private Sub lblBrowseRoot_Click()
Dim strPath As String
    blnDirty = True
    strPath = Util.BrowseForFolder(, True, Core.Server.HTTP.Config.WebRoot)
    If strPath <> "" Then
        txtWebroot.Text = strPath
    End If
End Sub

Private Sub lblBrowsevHostLog_Click()
Dim strDefaultFile As String

    blnDirty = True
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
    strDefaultFile = Mid$(Core.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, (InStrRev(Core.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid$(Core.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, 1, (Len(Core.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log) - InStrRev(Core.Server.HTTP.Config.VirtHost(lstvHosts.ListIndex + 1).Log, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtvHostLog.Text = dlgMain.FileName
    End If
End Sub

Private Sub lblBrowsevHostRoot_Click()
Dim strPath As String

    strPath = Util.BrowseForFolder(, True, Core.Server.HTTP.Config.VirtHost((lstvHosts.ListIndex + 1)).Root)
    If strPath <> "" Then
        txtvHostRoot.Text = strPath
    End If
End Sub

Private Sub lblCurrentStatus_Click()
    fraStatus.ZOrder 0
End Sub

Private Sub lblCurrentStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "This displays the most important status information about this application and your servers."
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblForum_Click()
    Core.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
End Sub

Private Sub lblHomePage_Click()
    Core.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub lblHTTPConfigAdv_Click()
    fraConfigAdv.ZOrder 0
End Sub

Private Sub lblHTTPConfigAdv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "From here you can adjust the servers advanced settings, these should options should be used with care."
End Sub

Private Sub lblHTTPConfigBasic_Click()
    fraConfigBasic.ZOrder 0
End Sub

Private Sub lblHTTPConfigBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "From here you can change your servers basic settings, such as it's port, logging information and other simple items."
End Sub

Private Sub lblHTTPConfigISAPI_Click()
    fraConfigISAPI.ZOrder 0
End Sub

Private Sub lblHTTPConfigISAPI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "From here you can adjust the setiings of your ISAPI plugins, such as PHP and other scripting languages."
End Sub

Private Sub lblHTTPConfigVirtHost_Click()
    fraConfigvHost.ZOrder 0
End Sub

Private Sub lblHTTPConfigVirtHost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "This allows you to host several domains from one server. This fature is useful for hosting many websites at the same time."
End Sub

Private Sub lblHTTPExportSettings_Click()
    'this needs some kind of error control, file checks, etc..
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgMain.ShowSave
    If dlgMain.FileName <> "" Then
        Open dlgMain.FileName For Append As 1
            Print #1, Core.Server.HTTP.Config.Report
        Close 1
    End If
End Sub

Private Sub lblISAPINew_Click()
Dim vItem As Variant
Dim i As Long

    Load frmNewISAPI
    frmNewISAPI.Show vbModal
    If Core.Server.HTTP.Config.ISAPI.Count > 0 Then
        lstISAPI.Clear
        For Each vItem In Core.Server.HTTP.Config.ISAPI
            lstISAPI.AddItem vItem.Extension
            lstISAPI.Enabled = True
        Next
    Else
        lstISAPI.Enabled = False
    End If
End Sub

Private Sub lblISAPIRemove_Click()
Dim lngRetVal As Long
Dim vItem As Variant
Dim i As Long

    If lstISAPI.ListIndex >= 0 Then
        lngRetVal = MsgBox(Translator.GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
        If lngRetVal = vbYes Then
            blnDirty = True
            Core.Server.HTTP.Config.ISAPI.Remove (lstISAPI.Text)
            lstISAPI.Clear
            If Core.Server.HTTP.Config.ISAPI.Count > 0 Then
                For Each vItem In Core.Server.HTTP.Config.ISAPI
                    lstISAPI.AddItem vItem.Extension
                    lstISAPI.Enabled = True
                Next
            Else
                lstISAPI.Enabled = False
                lblBrowseISAPIInterp.Enabled = False
                lblISAPIRemove.Enabled = False
                txtISAPIInterp.Enabled = False
                txtISAPIExt.Enabled = False
                txtISAPIInterp.Text = ""
                txtISAPIExt.Text = ""
            End If
        End If
    End If
End Sub

Private Sub lblOK_Click()
    If blnDirty <> False Then
        If Core.Server.HTTP.Config.Save(Core.Server.HTTP.Config.File) = False Then
            MsgBox Translator.GetText("Data was not saved, no idea why...")
        Else
            blnDirty = False
            Core.Server.HTTP.StopServer
            DoEvents
            Core.Server.HTTP.StartServer
            UpdateStats
            Me.Hide
        End If
    Else
        Me.WindowState = vbMinimized
        Me.Hide
    End If
End Sub

Private Sub lblRegister_Click()
    Core.Registration.Start
End Sub

Private Sub lblSrvRestart_Click()
    SetStatus Translator.GetText("Restarting Service") & "...", True
    Core.Server.HTTP.StopServer
    DoEvents
    Core.Server.HTTP.StartServer
    UpdateStats
    SetStatus Translator.GetText("Ready") & "..."
End Sub

Private Sub lblSrvStart_Click()
    SetStatus Translator.GetText("Starting Service") & "...", True
    Core.Server.HTTP.StartServer
    UpdateStats
    SetStatus "Ready..."
End Sub

Private Sub lblSrvStop_Click()
    SetStatus Translator.GetText("Stopping Service") & "...", True
    Core.Server.HTTP.StopServer
    SetStatus "Ready..."
End Sub

Private Sub lblUpdates_Click()
    SetStatus Translator.GetText("Retrieving Update Information") & "...", True
    Core.Update.Check
    If Core.Update.IsAvailable = True Then
        lblUpdateStatus.Caption = Translator.GetText("New Version Available")
        lblUpdateStatus.Font.Underline = True
        lblUpdateStatus.ForeColor = vbBlue
        lblUpdateStatus.MousePointer = vbCustom
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox Translator.GetText("You have the most current version available."), vbOKOnly + vbInformation
    End If
    SetStatus "Ready..."
End Sub

Private Sub lblUpdateStatus_Click()
    If Core.Update.IsAvailable = True Then
        Load frmUpdate
        frmUpdate.Show
    End If
End Sub

Private Sub lblvHostNew_Click()
    fraNewvHost.ZOrder 0
End Sub

Private Sub lblvHostRemove_Click()
Dim lngRetVal As Long
Dim blnMore As Boolean
Dim vItem As Variant

    If lstvHosts.ListIndex >= 0 Then
        lngRetVal = MsgBox(Translator.GetText("Are you sure you want to delete this item?\r\rThis can not be undone."), vbQuestion + vbYesNo)
        If lngRetVal = vbYes Then
            blnDirty = True
            Core.Server.HTTP.Config.VirtHost.Remove lstvHosts.Text
            txtvHostName.Text = ""
            txtvHostDomain.Text = ""
            txtvHostRoot.Text = ""
            txtvHostLog.Text = ""
            lstvHosts.Clear
            For Each vItem In Core.Server.HTTP.Config.VirtHost
                lstvHosts.AddItem vItem.HostName
                blnMore = True
            Next
            If blnMore = False Then
                lblBrowsevHostRoot.Enabled = False
                lblBrowsevHostLog.Enabled = False
                lblvHostRemove.Enabled = False
                txtvHostName.Enabled = False
                txtvHostDomain.Enabled = False
                txtvHostRoot.Enabled = False
                txtvHostLog.Enabled = False
                lstvHosts.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lblViewLogs_Click()
    fraLogs.ZOrder 0
End Sub

Private Sub lblViewLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMoreInfoData.Caption = "Displays the log files for your servers. From here you can view any events, errors and a log of all connections."
End Sub

Private Sub lstISAPI_Click()
    lblBrowseISAPIInterp.Enabled = True
    lblISAPIRemove.Enabled = True
    txtISAPIInterp.Enabled = True
    txtISAPIExt.Enabled = True
    txtISAPIInterp.Text = Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter
    txtISAPIExt.Text = Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension
End Sub

Private Sub lstvHosts_Click()
    lblBrowsevHostRoot.Enabled = True
    lblBrowsevHostLog.Enabled = True
    lblvHostRemove.Enabled = True
    txtvHostName.Enabled = True
    txtvHostDomain.Enabled = True
    txtvHostRoot.Enabled = True
    txtvHostLog.Enabled = True
    txtvHostName.Text = Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName
    txtvHostDomain.Text = Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain
    txtvHostRoot.Text = Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root
    txtvHostLog.Text = Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log
End Sub

Private Sub mnuSysTrayPopupAbout_Click()
    Me.MousePointer = 11
    DoEvents
    Load frmAbout
    Me.MousePointer = 0
    frmAbout.Show
End Sub

Private Sub mnuSysTrayPopupExit_Click()
    Unload Me
End Sub

Private Sub mnuSysTrayPopupForum_Click()
    Core.Net.LaunchURL "http://swebs.sourceforge.net/html/modules.php?op=modload&name=PNphpBB2&file=index"
End Sub

Private Sub mnuSysTrayPopupHomePage_Click()
    Core.Net.LaunchURL "http://swebs.sourceforge.net/html/index.php"
End Sub

Private Sub mnuSysTrayPopupOpenCC_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub mnuSysTrayPopupUpdate_Click()
    SetStatus Translator.GetText("Retrieving Update Information") & "...", True
    Core.Update.Check
    If Core.Update.IsAvailable = True Then
        lblUpdateStatus.Caption = Translator.GetText("New Version Available")
        lblUpdateStatus.Font.Underline = True
        lblUpdateStatus.ForeColor = vbBlue
        lblUpdateStatus.MousePointer = vbCustom
        Load frmUpdate
        frmUpdate.Show
    Else
        MsgBox Translator.GetText("You have the most current version available."), vbOKOnly + vbInformation
    End If
    SetStatus Translator.GetText("Ready") & "..."
End Sub

Private Sub SysTray_LButtonDblClk()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub SysTray_RButtonUp()
    SetForegroundWindow Me.hWnd
    PopupMenu mnuSysTrayPopup, , , , mnuSysTrayPopupOpenCC
    PostMessage Me.hWnd, 0&, 0&, 0&
End Sub

Private Sub tmrStats_Timer()
    UpdateStats
End Sub

Private Sub tmrStatus_Timer()
Dim strSrvStatusCur As String

    strSrvStatusCur = Core.Server.HTTP.Status
    lblSrvStatusCur.Font.Bold = False
    Select Case strSrvStatusCur
        Case "Stopped"
            lblSrvStatusCur.Caption = Translator.GetText("Stopped")
            Core.EventLog.AddEvent "SWEBS_Core_Main.frmMain.tmrStatus_Timer", "Service Status: Stopped"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            lblSrvStart.Enabled = True
            lblSrvStop.Enabled = False
            lblSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = Translator.GetText("Running")
            Core.EventLog.AddEvent "SWEBS_Core_Main.frmMain.tmrStatus_Timer", "Service Status: Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            lblSrvStart.Enabled = False
            lblSrvStop.Enabled = True
            lblSrvRestart.Enabled = True
    End Select
End Sub


Private Function LoadConfigData() As Boolean
Dim strTemp As String
Dim strResult As String
Dim vItem As Variant
    
    Core.EventLog.AddEvent "SWEBS_Core_Main.frmMain.LoadConfigData", "Loading Config Data"
    SetStatus Translator.GetText("Loading Configuration Data") & "...", True
    LoadConfigData = Core.Server.HTTP.Config.LoadData
    
    'Setup the form...
    txtServerName.Text = Core.Server.HTTP.Config.ServerName
    txtPort.Text = Core.Server.HTTP.Config.Port
    txtWebroot.Text = Core.Server.HTTP.Config.WebRoot
    txtMaxConnect.Text = Core.Server.HTTP.Config.MaxConnections
    txtLogFile.Text = Core.Server.HTTP.Config.LogFile
    txtConfigAdvIPBind.Text = Core.Server.HTTP.Config.ListeningAddress
    txtAllowIndex.Text = Core.Server.HTTP.Config.AllowIndex
    txtErrorPages.Text = Core.Server.HTTP.Config.ErrorPages
    txtConfigBasicErrorLog.Text = Core.Server.HTTP.Config.ErrorLog
    
    For Each vItem In Core.Server.HTTP.Config.Index
        strTemp = strTemp & vItem.FileName & " "
    Next
    txtIndexFiles.Text = Trim$(strTemp)
    
    lstISAPI.Enabled = False
    lstISAPI.Clear
    For Each vItem In Core.Server.HTTP.Config.ISAPI
        lstISAPI.AddItem vItem.Extension
        lstISAPI.Enabled = True
    Next
    
    lstvHosts.Enabled = False
    lstvHosts.Clear
    For Each vItem In Core.Server.HTTP.Config.VirtHost
        lstvHosts.AddItem vItem.HostName
        lstvHosts.Enabled = True
    Next
    
    cmbViewLogFiles.Clear
    If Dir$(Core.Server.HTTP.Config.LogFile) <> "" Then
        cmbViewLogFiles.AddItem Core.Server.HTTP.Config.LogFile
    End If
    If Dir$(Core.Server.HTTP.Config.ErrorLog) <> "" Then
        cmbViewLogFiles.AddItem Core.Server.HTTP.Config.ErrorLog
    End If
    For Each vItem In Core.Server.HTTP.Config.VirtHost
        If Dir$(vItem.Log) <> "" Then
            cmbViewLogFiles.AddItem vItem.Log
        End If
    Next
    
    'we now only check for updates every 24 hours, this could confuse some people.
    'but this should make loading faster.
    SetStatus "Checking For Updates...", True
    strResult = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "LastUpdateCheck")
    If strResult = "" Then
        strResult = CDate(1.1)
    End If
    If DateDiff("h", CDate(strResult), Now) >= 24 Then
        Core.Update.Check
        If Core.Update.IsAvailable = True Then
            lblUpdateStatus.Caption = Translator.GetText("New Version Available")
        Else
            lblUpdateStatus.Caption = Translator.GetText("No Updates Available")
            lblUpdateStatus.Font.Underline = False
            lblUpdateStatus.ForeColor = vbButtonText
            lblUpdateStatus.MousePointer = vbDefault
            Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "LastUpdateCheck", Now
        End If
    Else
        lblUpdateStatus.Caption = Translator.GetText("No Updates Available")
        lblUpdateStatus.Font.Underline = False
        lblUpdateStatus.ForeColor = vbButtonText
        lblUpdateStatus.MousePointer = vbDefault
    End If
    
    UpdateStats
        
    If Core.Registration.IsRegistered = True Then
        SetStatus "Updating Registration..."
        lblRegister.Enabled = False
        Core.Registration.Renew
    End If
    
    SetStatus "Ready..."
End Function

Private Sub txtAllowIndex_Change()
    If Core.Server.HTTP.Config.AllowIndex <> IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false") Then
        Core.Server.HTTP.Config.AllowIndex = IIf(LCase$(txtAllowIndex.Text) = "true", "true", "false")
        blnDirty = True
    End If
End Sub

Private Sub txtISAPIExt_Change()
    If lstISAPI.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension <> txtISAPIExt.Text Then
            Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Extension = txtISAPIExt.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtISAPIInterp_Change()
    If lstISAPI.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter <> txtISAPIInterp.Text Then
            Core.Server.HTTP.Config.ISAPI.Item(lstISAPI.Text).Interpreter = txtISAPIInterp.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtConfigAdvIPBind_Change()
    If Core.Server.HTTP.Config.ListeningAddress <> txtConfigAdvIPBind.Text Then
        Core.Server.HTTP.Config.ListeningAddress = txtConfigAdvIPBind.Text
        blnDirty = True
    End If
End Sub

Private Sub txtConfigBasicErrorLog_Change()
    If Core.Server.HTTP.Config.ErrorLog <> txtConfigBasicErrorLog.Text Then
        Core.Server.HTTP.Config.ErrorLog = txtConfigBasicErrorLog.Text
        blnDirty = True
    End If
End Sub

Private Sub txtErrorPages_Change()
    If Core.Server.HTTP.Config.ErrorPages <> txtErrorPages.Text Then
        Core.Server.HTTP.Config.ErrorPages = txtErrorPages.Text
        blnDirty = True
    End If
End Sub

Private Sub txtIndexFiles_Change()
Dim strTmpArray() As String
Dim lngRecCount As Long
Dim i As Long
    strTmpArray = Split(Trim$(txtIndexFiles.Text), " ")
    If Not IsEmpty(strTmpArray) Then
        Core.Server.HTTP.Config.Index.Clear
        lngRecCount = UBound(strTmpArray)
        For i = 0 To lngRecCount
            Core.Server.HTTP.Config.Index.Add strTmpArray(i)
        Next
    End If
End Sub

Private Sub txtIndexFiles_KeyPress(KeyAscii As Integer)
    blnDirty = True
End Sub

Private Sub txtIndexFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDirty = True
End Sub

Private Sub txtLogFile_Change()
    If Core.Server.HTTP.Config.LogFile <> Trim$(txtLogFile.Text) Then
        Core.Server.HTTP.Config.LogFile = Trim$(txtLogFile.Text)
        blnDirty = True
    End If
End Sub

Private Sub txtMaxConnect_Change()
    If Core.Server.HTTP.Config.MaxConnections <> Int(Val(txtMaxConnect.Text)) Then
        Core.Server.HTTP.Config.MaxConnections = Int(Val(txtMaxConnect.Text))
        blnDirty = True
    End If
End Sub

Private Sub txtPort_Change()
    If Core.Server.HTTP.Config.Port <> Int(Val(txtPort.Text)) Then
        Core.Server.HTTP.Config.Port = Int(Val(txtPort.Text))
        blnDirty = True
    End If
End Sub

Private Sub txtServerName_Change()
    If Core.Server.HTTP.Config.ServerName <> Trim$(txtServerName.Text) Then
        Core.Server.HTTP.Config.ServerName = Trim$(txtServerName.Text)
        blnDirty = True
    End If
End Sub

Private Sub txtvHostDomain_Change()
    If lstvHosts.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain <> txtvHostDomain.Text Then
            Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Domain = txtvHostDomain.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtvHostLog_Change()
    If lstvHosts.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log <> txtvHostLog.Text Then
            Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Log = txtvHostLog.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtvHostName_Change()
    If lstvHosts.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName <> txtvHostName.Text Then
            blnDirty = True
            Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).HostName = txtvHostName.Text
        End If
    End If
End Sub

Private Sub txtvHostRoot_Change()
    If lstvHosts.ListIndex <> -1 Then
        If Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root <> txtvHostRoot.Text Then
            Core.Server.HTTP.Config.VirtHost.Item(lstvHosts.Text).Root = txtvHostRoot.Text
            blnDirty = True
        End If
    End If
End Sub

Private Sub txtWebroot_Change()
    If Core.Server.HTTP.Config.WebRoot <> Trim$(txtWebroot.Text) Then
        Core.Server.HTTP.Config.WebRoot = Trim$(txtWebroot.Text)
        blnDirty = True
    End If
End Sub

Private Sub UpdateStats()
    Core.Server.HTTP.Stats.Reload
    lblStatsLastRestart.Caption = Translator.GetText("Last Restart") & ": " & Core.Server.HTTP.Stats.LastRestart
    lblStatsRequestCount.Caption = Translator.GetText("Request Count") & ": " & Core.Server.HTTP.Stats.RequestCount
    lblStatsBytesSent.Caption = Translator.GetText("Total Bytes Sent") & ": " & Format$(Core.Server.HTTP.Stats.TotalBytesSent, "###,###,###,###,##0")
    lblCurVersion.Caption = Translator.GetText("Current Version") & ": " & Core.Version
    lblUpdateVersion.Caption = Translator.GetText("Update Version") & ": " & IIf(Core.Update.Version <> "", Core.Update.Version, Core.Version)
End Sub
