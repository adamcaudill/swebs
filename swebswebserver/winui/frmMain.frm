VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SWEBS Web Server - Control Center"
   ClientHeight    =   4920
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   4320
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Server Status"
      TabPicture(0)   =   "frmMain.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tmrStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSrvStatus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Configuration"
      TabPicture(1)   =   "frmMain.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sstConfig"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logs"
      TabPicture(2)   =   "frmMain.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtViewLogFiles"
      Tab(2).Control(1)=   "cmbViewLogFiles"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtViewLogFiles 
         Appearance      =   0  'Flat
         Height          =   3255
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   41
         Top             =   840
         Width           =   5775
      End
      Begin VB.ComboBox cmbViewLogFiles 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":0D1E
         Left            =   -74880
         List            =   "frmMain.frx":0D20
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   480
         Width           =   5775
      End
      Begin VB.Frame fraSrvStatus 
         Caption         =   "Current Service Status:"
         Height          =   1095
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   3375
         Begin VB.CommandButton cmdSrvRestart 
            Caption         =   "Restart"
            Height          =   375
            Left            =   2040
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdSrvStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   1080
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdSrvStart 
            Caption         =   "Start"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblSrvStatusCur 
            Caption         =   "<current-status>"
            Height          =   255
            Left            =   720
            TabIndex        =   33
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblSrvStatus 
            Caption         =   "Status: "
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Timer tmrStatus 
         Interval        =   750
         Left            =   5520
         Top             =   360
      End
      Begin TabDlg.SSTab sstConfig 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6165
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
         TabPicture(0)   =   "frmMain.frx":0D22
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblServerName"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblPort"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblWebroot"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtServerName"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtPort"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtWebroot"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdBrowseRoot"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Advanced"
         TabPicture(1)   =   "frmMain.frx":0D3E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdBrowseLogFile"
         Tab(1).Control(1)=   "txtLogFile"
         Tab(1).Control(2)=   "txtAllowIndex"
         Tab(1).Control(3)=   "txtIndexFiles"
         Tab(1).Control(4)=   "txtMaxConnect"
         Tab(1).Control(5)=   "lblLogFile"
         Tab(1).Control(6)=   "lblAllowIndex"
         Tab(1).Control(7)=   "lblIndexFiles"
         Tab(1).Control(8)=   "lblMaxConnect"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "vHosts"
         TabPicture(2)   =   "frmMain.frx":0D5A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdBrowsevHostLog"
         Tab(2).Control(1)=   "cmdBrowsevHostRoot"
         Tab(2).Control(2)=   "txtvHostLog"
         Tab(2).Control(3)=   "txtvHostRoot"
         Tab(2).Control(4)=   "txtvHostDomain"
         Tab(2).Control(5)=   "txtvHostName"
         Tab(2).Control(6)=   "lstvHosts"
         Tab(2).Control(7)=   "lblvHostLog"
         Tab(2).Control(8)=   "lblvHostRoot"
         Tab(2).Control(9)=   "lblvHostDomain"
         Tab(2).Control(10)=   "lblvHostName"
         Tab(2).ControlCount=   11
         TabCaption(3)   =   "CGI Handlers"
         TabPicture(3)   =   "frmMain.frx":0D76
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdBrowseCGIInterp"
         Tab(3).Control(1)=   "txtCGIExt"
         Tab(3).Control(2)=   "txtCGIInterp"
         Tab(3).Control(3)=   "lstCGI"
         Tab(3).Control(4)=   "lblCGIExt"
         Tab(3).Control(5)=   "lblCGIInterp"
         Tab(3).ControlCount=   6
         Begin VB.CommandButton cmdBrowseCGIInterp 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69720
            TabIndex        =   44
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostLog 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -70200
            TabIndex        =   43
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseLogFile 
            Caption         =   "..."
            Height          =   255
            Left            =   -72120
            TabIndex        =   42
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowsevHostRoot 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   -70200
            TabIndex        =   39
            Top             =   1920
            Width           =   255
         End
         Begin VB.CommandButton cmdBrowseRoot 
            Caption         =   "..."
            Height          =   255
            Left            =   3960
            TabIndex        =   38
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox txtCGIExt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72480
            TabIndex        =   30
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtCGIInterp 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72480
            TabIndex        =   28
            Top             =   840
            Width           =   2655
         End
         Begin VB.ListBox lstCGI 
            Appearance      =   0  'Flat
            Height          =   2565
            ItemData        =   "frmMain.frx":0D92
            Left            =   -74880
            List            =   "frmMain.frx":0D94
            TabIndex        =   26
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtvHostLog 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72720
            TabIndex        =   25
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox txtvHostRoot 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72720
            TabIndex        =   23
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtvHostDomain 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72720
            TabIndex        =   21
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtvHostName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72720
            TabIndex        =   19
            Top             =   720
            Width           =   2415
         End
         Begin VB.ListBox lstvHosts 
            Appearance      =   0  'Flat
            Height          =   2760
            ItemData        =   "frmMain.frx":0D96
            Left            =   -74880
            List            =   "frmMain.frx":0D98
            TabIndex        =   17
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtLogFile 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   12
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtAllowIndex 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   11
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtIndexFiles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   10
            Top             =   2760
            Width           =   2655
         End
         Begin VB.TextBox txtMaxConnect 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   9
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtWebroot 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   2400
            Width           =   3735
         End
         Begin VB.TextBox txtPort 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtServerName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblCGIExt 
            Caption         =   "What is the file extension?"
            Height          =   255
            Left            =   -72600
            TabIndex        =   29
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblCGIInterp 
            Caption         =   "What Interpreter would you like to use?"
            Height          =   255
            Left            =   -72600
            TabIndex        =   27
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lblvHostLog 
            Caption         =   "Where to you want to keep the log file?"
            Height          =   255
            Left            =   -72840
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label lblvHostRoot 
            Caption         =   "Where is the root folder?"
            Height          =   255
            Left            =   -72840
            TabIndex        =   22
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblvHostDomain 
            Caption         =   "What is it's domain?"
            Height          =   255
            Left            =   -72840
            TabIndex        =   20
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lblvHostName 
            Caption         =   "What is the name of this vHost?"
            Height          =   255
            Left            =   -72840
            TabIndex        =   18
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblLogFile 
            Caption         =   "Log File"
            Height          =   255
            Left            =   -74880
            TabIndex        =   16
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label lblAllowIndex 
            Caption         =   "Allow index"
            Height          =   255
            Left            =   -74880
            TabIndex        =   15
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label lblIndexFiles 
            Caption         =   "Index files"
            Height          =   255
            Left            =   -74880
            TabIndex        =   14
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label lblMaxConnect 
            Caption         =   "How many simultaneous connections do you want to accept?"
            Height          =   495
            Left            =   -74880
            TabIndex        =   13
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblWebroot 
            Caption         =   "Where is the root folder for your server?"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2160
            Width           =   3015
         End
         Begin VB.Label lblPort 
            Caption         =   "What port do you want to use? (Default is 80)"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label lblServerName 
            Caption         =   "What is the name of your server?"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   2535
         End
      End
   End
   Begin VB.Label lblAppStatus 
      Caption         =   "Current App Status..."
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   4560
      Width           =   4095
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
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

Private Sub cmbViewLogFiles_Click()
'this will be slow as ^(*& on big files, should be fixed someday
Dim strLog As String
Dim strTemp As String
    AppStatus True, "Loading Log File..."
    If Dir(cmbViewLogFiles.Text) <> "" Then
        Open cmbViewLogFiles.Text For Random As 1
            Do Until EOF(1)
                Line Input #1, strTemp
                strLog = strLog & strTemp
            Loop
        Close 1
    Else
        DoEvents
        MsgBox "File not found, it may not have been created yet."
    End If
    AppStatus False
End Sub

Private Sub cmdBrowseCGIInterp_Click()
Dim strDefaultFile As String
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Executable Files (*.exe)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid(Config.CGI((lstCGI.ListIndex + 1), 1), (InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid(Config.CGI((lstCGI.ListIndex + 1), 1), 1, (Len(Config.CGI((lstCGI.ListIndex + 1), 1)) - InStrRev(Config.CGI((lstCGI.ListIndex + 1), 1), "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtCGIInterp.Text = dlgMain.FileName
    End If
End Sub

Private Sub cmdBrowseRoot_Click()
Dim strPath As String
    strPath = BrowseForFolder(Me, , True, Config.WebRoot)
    If strPath <> "" Then
        txtWebroot.Text = strPath
    End If
End Sub

Private Sub cmdBrowsevHostLog_Click()
Dim strDefaultFile As String
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid(Config.vHost((lstvHosts.ListIndex + 1), 4), (InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid(Config.vHost((lstvHosts.ListIndex + 1), 4), 1, (Len(Config.vHost((lstvHosts.ListIndex + 1), 4)) - InStrRev(Config.vHost((lstvHosts.ListIndex + 1), 4), "\")))
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
    dlgMain.DialogTitle = "Please select a file..."
    dlgMain.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    strDefaultFile = Mid(Config.LogFile, (InStrRev(Config.LogFile, "\") + 1))
    dlgMain.FileName = strDefaultFile
    dlgMain.InitDir = Mid(Config.LogFile, 1, (Len(Config.LogFile) - InStrRev(Config.LogFile, "\")))
    dlgMain.ShowSave
    If dlgMain.FileName <> strDefaultFile Then
        txtLogFile.Text = dlgMain.FileName
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
    tmrStatus_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lngRetVal As Long
    lngRetVal = MsgBox("Do you want to save your settings before closing?", vbYesNo + vbQuestion + vbApplicationModal)
    If lngRetVal = vbYes Then
        If SaveConfigData(strConfigFile) = False Then
            MsgBox "Data was not saved, no idea why..."
        End If
    End If
    Me.Visible = False
    DoEvents
End Sub

Private Sub lstCGI_Click()
    cmdBrowseCGIInterp.Enabled = True
    txtCGIInterp.Enabled = True
    txtCGIExt.Enabled = True
    txtCGIInterp.Text = Config.CGI((lstCGI.ListIndex + 1), 1)
    txtCGIExt.Text = Config.CGI((lstCGI.ListIndex + 1), 2)
End Sub

Private Sub lstvHosts_Click()
    cmdBrowsevHostRoot.Enabled = True
    cmdBrowsevHostLog.Enabled = True
    txtvHostName.Enabled = True
    txtvHostDomain.Enabled = True
    txtvHostRoot.Enabled = True
    txtvHostLog.Enabled = True
    txtvHostName.Text = Config.vHost((lstvHosts.ListIndex + 1), 1)
    txtvHostDomain.Text = Config.vHost((lstvHosts.ListIndex + 1), 2)
    txtvHostRoot.Text = Config.vHost((lstvHosts.ListIndex + 1), 3)
    txtvHostLog.Text = Config.vHost((lstvHosts.ListIndex + 1), 4)
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

Private Sub AppStatus(blnBusy As Boolean, Optional strMessage As String)
    If blnBusy = True Then
        Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
    Else
        Screen.MousePointer = vbDefault  '0 default
    End If
    If strMessage = "" Then
        lblAppStatus.Caption = "Ready..."
    Else
        lblAppStatus.Caption = strMessage
    End If
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
    For i = 1 To UBound(Config.Index)
        strTemp = strTemp & Config.Index(i) & " "
    Next i
    txtIndexFiles.Text = strTemp
    If Config.CGI(1, 2) <> "" Then
        lstCGI.Clear
        For i = 1 To UBound(Config.CGI)
            lstCGI.AddItem Config.CGI(i, 2)
        Next i
    Else
        lstCGI.Enabled = False
    End If
    If Config.vHost(1, 1) <> "" Then
        lstvHosts.Clear
        For i = 1 To UBound(Config.vHost)
            lstvHosts.AddItem Config.vHost(i, 1)
        Next i
    Else
        lstvHosts.Enabled = False
    End If
    cmbViewLogFiles.Clear
    cmbViewLogFiles.AddItem Config.LogFile
    For i = 1 To UBound(Config.vHost)
        cmbViewLogFiles.AddItem Config.vHost(i, 4)
    Next i
    AppStatus False
End Function

Private Sub txtAllowIndex_Change()
    Config.AllowIndex = IIf(LCase(txtAllowIndex.Text) = "true", "true", "false")
End Sub

Private Sub txtCGIExt_Change()
    Config.CGI((lstCGI.ListIndex + 1), 2) = txtCGIExt.Text
End Sub

Private Sub txtCGIInterp_Change()
    Config.CGI((lstCGI.ListIndex + 1), 1) = txtCGIInterp.Text
End Sub

Private Sub txtIndexFiles_Change()
Dim strTmpArray() As String
Dim i As Long
    strTmpArray = Split(Trim(txtIndexFiles.Text), " ")
    ReDim Config.Index(1 To (UBound(strTmpArray) + 1))
    For i = 0 To UBound(strTmpArray)
        Config.Index(i + 1) = strTmpArray(i)
    Next
End Sub

Private Sub txtLogFile_Change()
    Config.LogFile = Trim(txtLogFile.Text)
End Sub

Private Sub txtMaxConnect_Change()
    Config.MaxConnections = Int(Val(txtMaxConnect.Text))
End Sub

Private Sub txtPort_Change()
    Config.Port = Int(Val(txtPort.Text))
End Sub

Private Sub txtServerName_Change()
    Config.ServerName = Trim(txtServerName.Text)
End Sub

Private Sub txtvHostDomain_Change()
    Config.vHost((lstvHosts.ListIndex + 1), 2) = txtvHostDomain.Text
End Sub

Private Sub txtvHostLog_Change()
    Config.vHost((lstvHosts.ListIndex + 1), 4) = txtvHostLog.Text
End Sub

Private Sub txtvHostName_Change()
    Config.vHost((lstvHosts.ListIndex + 1), 1) = txtvHostName.Text
End Sub

Private Sub txtvHostRoot_Change()
    Config.vHost((lstvHosts.ListIndex + 1), 3) = txtvHostRoot.Text
End Sub

Private Sub txtWebroot_Change()
    Config.WebRoot = Trim(txtWebroot.Text)
End Sub

