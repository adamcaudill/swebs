VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewVirtHost 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Virtual Host"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewvHostName 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtNewvHostDomain 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtNewvHostRoot 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtNewvHostLogs 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   5640
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblBrowseLog 
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
      Left            =   5400
      MouseIcon       =   "frmNewVirtHost.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2880
      Width           =   660
   End
   Begin VB.Label lblBrowseRoot 
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
      Left            =   5400
      MouseIcon       =   "frmNewVirtHost.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
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
      Left            =   3319
      MouseIcon       =   "frmNewVirtHost.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3480
      Width           =   585
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2306
      MouseIcon       =   "frmNewVirtHost.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblNewvHostName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is the name of this Virtual Host?"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblNewvHostRoot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Where is the root folder for this Virtual Host?"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3105
   End
   Begin VB.Label lblNewvHostDomain 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is the domain for this Virtual Host?"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2820
   End
   Begin VB.Label lblNewvHostLogs 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Where do you want to keep the log for this Virtual Host?"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   3960
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Virtual Host"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2850
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H00804008&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6330
   End
   Begin VB.Line Line1 
      X1              =   6330
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmNewVirtHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblOK.Caption = Translator.GetText("&OK")
    lblCancel.Caption = Translator.GetText("&Cancel")
    lblTitle.Caption = Translator.GetText("Add a new Virtual Host")
    lblNewvHostName.Caption = Translator.GetText("What is the name of this Virtual Host?")
    lblNewvHostDomain.Caption = Translator.GetText("What is the domain for this Virtual Host?")
    lblNewvHostRoot.Caption = Translator.GetText("Where is the root folder for this Virtual Host?")
    lblNewvHostLogs.Caption = Translator.GetText("Where do you want to keep the log for this Virtual Host?")
End Sub

Private Sub lblBrowseLog_Click()
    dlgMain.DialogTitle = Translator.GetText("Please select a file...")
    dlgMain.Filter = Translator.GetText("Log Files (*.log)|*.log|All Files (*.*)|*.*")
    dlgMain.InitDir = Core.Path
    dlgMain.ShowSave
    txtNewvHostLogs.Text = dlgMain.FileName
End Sub

Private Sub lblBrowseRoot_Click()
Dim strPath As String

    strPath = Util.BrowseForFolder(, True, Core.Server.HTTP.Config.WebRoot)
    If strPath <> "" Then
        txtNewvHostRoot.Text = strPath
    End If
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblOK_Click()
    If txtNewvHostName.Text <> "" And txtNewvHostDomain.Text <> "" And txtNewvHostRoot.Text <> "" And txtNewvHostLogs.Text <> "" Then
        Core.Server.HTTP.Config.VirtHost.Add txtNewvHostName.Text, txtNewvHostDomain.Text, txtNewvHostRoot.Text, txtNewvHostLogs.Text, txtNewvHostName.Text
        Unload Me
    Else
        MsgBox Translator.GetText("Please fill all fields.")
    End If
End Sub
