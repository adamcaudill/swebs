VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logs"
      TabPicture(2)   =   "frmMain.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLogs"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraSrvStatus 
         Caption         =   "Current Service Status:"
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   3375
         Begin VB.CommandButton cmdSrvRestart 
            Caption         =   "Restart"
            Height          =   375
            Left            =   2040
            TabIndex        =   37
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdSrvStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   1080
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdSrvStart 
            Caption         =   "Start"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblSrvStatusCur 
            Caption         =   "<current-status>"
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblSrvStatus 
            Caption         =   "Status: "
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Timer tmrStatus 
         Interval        =   1000
         Left            =   4200
         Top             =   480
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
         TabPicture(0)   =   "frmMain.frx":0D1E
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
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Advanced"
         TabPicture(1)   =   "frmMain.frx":0D3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtLogFile"
         Tab(1).Control(1)=   "txtAllowIndex"
         Tab(1).Control(2)=   "txtIndexFiles"
         Tab(1).Control(3)=   "txtMaxConnect"
         Tab(1).Control(4)=   "lblLogFile"
         Tab(1).Control(5)=   "lblAllowIndex"
         Tab(1).Control(6)=   "lblIndexFiles"
         Tab(1).Control(7)=   "lblMaxConnect"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "vHosts"
         TabPicture(2)   =   "frmMain.frx":0D56
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtvHostLog"
         Tab(2).Control(1)=   "txtvHostRoot"
         Tab(2).Control(2)=   "txtvHostDomain"
         Tab(2).Control(3)=   "txtvHostName"
         Tab(2).Control(4)=   "lstvHosts"
         Tab(2).Control(5)=   "lblvHostLog"
         Tab(2).Control(6)=   "lblvHostRoot"
         Tab(2).Control(7)=   "lblvHostDomain"
         Tab(2).Control(8)=   "lblvHostName"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "CGI Handlers"
         TabPicture(3)   =   "frmMain.frx":0D72
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtCGIExt"
         Tab(3).Control(1)=   "txtCGIInterp"
         Tab(3).Control(2)=   "lstCGI"
         Tab(3).Control(3)=   "lblCGIExt"
         Tab(3).Control(4)=   "lblCGIInterp"
         Tab(3).ControlCount=   5
         Begin VB.TextBox txtCGIExt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72480
            TabIndex        =   31
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox txtCGIInterp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72480
            TabIndex        =   29
            Top             =   840
            Width           =   2895
         End
         Begin VB.ListBox lstCGI 
            Appearance      =   0  'Flat
            Height          =   2565
            ItemData        =   "frmMain.frx":0D8E
            Left            =   -74880
            List            =   "frmMain.frx":0D90
            TabIndex        =   27
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtvHostLog 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72600
            TabIndex        =   26
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox txtvHostRoot 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72600
            TabIndex        =   24
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtvHostDomain 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72600
            TabIndex        =   22
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtvHostName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72600
            TabIndex        =   20
            Top             =   720
            Width           =   2415
         End
         Begin VB.ListBox lstvHosts 
            Appearance      =   0  'Flat
            Height          =   2760
            ItemData        =   "frmMain.frx":0D92
            Left            =   -74880
            List            =   "frmMain.frx":0D94
            TabIndex        =   18
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
            Caption         =   "What is the file extention?"
            Height          =   255
            Left            =   -72600
            TabIndex        =   30
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblCGIInterp 
            Caption         =   "What Interpreter would you like to use?"
            Height          =   255
            Left            =   -72600
            TabIndex        =   28
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lblvHostLog 
            Caption         =   "Where to you want to keep the log file?"
            Height          =   255
            Left            =   -72720
            TabIndex        =   25
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label lblvHostRoot 
            Caption         =   "Where is the root folder?"
            Height          =   255
            Left            =   -72720
            TabIndex        =   23
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblvHostDomain 
            Caption         =   "What is it's domain?"
            Height          =   255
            Left            =   -72720
            TabIndex        =   21
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lblvHostName 
            Caption         =   "What is the name of this vHost?"
            Height          =   255
            Left            =   -72840
            TabIndex        =   19
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
            Caption         =   "How many simultainious connections do you want to accept?"
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
      Begin VB.Label lblLogs 
         Caption         =   $"frmMain.frx":0D96
         Height          =   1095
         Left            =   -74160
         TabIndex        =   17
         Top             =   1440
         Width           =   3975
      End
   End
   Begin VB.Label lblAppStatus 
      Caption         =   "Current App Status..."
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Data..."
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSrvRestart_Click()
    AppStatus True, "Restarting Service..."
    ServiceStop "", "SWS Web Server"
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
    If LoadConfigData(strConfigFile) = False Then
        RetVal = MsgBox("There was an error while loading your configuration data." & vbCrLf & vbCrLf & "Press 'Abort' to give up and exit, 'Retry' to try to load th data again," & vbCrLf & "or 'Ignore' to continue.", vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                End
            Case vbRetry
                If LoadConfigData(strConfigFile) = False Then
                    MsgBox "A second attempt to load your configuration data failed. Aborting." & vbCrLf & vbCrLf & "This application will now close.", vbApplicationModal + vbCritical
                    End
                End If
            Case vbIgnore
                MsgBox "NOTICE: You have chosen to proceed after a data error, this application may" & vbCrLf & "not function properly or you may loose data."
        End Select
    End If
    tmrStatus_Timer
End Sub

Private Function LoadConfigData(strCurConfigFile As String) As Boolean
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       WinUI
' Procedure  :       LoadConfigData
' Description:       This loads the data from the config XML file, returns true
'                    if the load is sucessful, otherwise returns false
' Created by :       Adam
' Date-Time  :       8/24/2003-3:01:42 PM
' Parameters :       strCurConfigFile (String)
'--------------------------------------------------------------------------------
'</CSCM>

Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
    
    AppStatus True, "Loading Configuration Data..."
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile strCurConfigFile
    
    '<ServerName>
    Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
    txtServerName.Text = Node.Content
    
    '<Port>
    Set Node = ConfigXML.SearchForTag(Nothing, "Port")
    txtPort.Text = Node.Content
    
    '<Webroot>
    Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
    txtWebroot.Text = Node.Content
    
    '<MaxConnections>
    Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
    txtMaxConnect.Text = Node.Content
    
    '<LogFile>
    Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
    txtLogFile.Text = Node.Content
    
    '<AllowIndex>
    Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
    txtAllowIndex.Text = Node.Content
    
    '<IndexFile>
    txtIndexFiles.Text = ""
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    Do While Not (Node Is Nothing)
        txtIndexFiles.Text = txtIndexFiles.Text & Node.Content & " "
        Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
    Loop
    
    '<VirtualHost>
    lstvHosts.Clear
    Set Node = ConfigXML.FindChild("VirtualHost")
    Do While Not (Node Is Nothing)
        If Node.Content <> "" Then
            lstvHosts.AddItem Node.Content
        End If
        Set Node = ConfigXML.SearchForTag(Node, "vhName")
    Loop
    
    '<CGI>
    lstCGI.Clear
    Set Node = ConfigXML.FindChild("CGI")
    Do While Not (Node Is Nothing)
        If Node.Content <> "" Then
            lstCGI.AddItem Node.Content
        End If
        Set Node = ConfigXML.SearchForTag(Node, "Extension")
    Loop
    
    Set XML = Nothing
    Set ConfigXML = Nothing
    Set Node = Nothing
    AppStatus False
    LoadConfigData = True
End Function

Private Sub lstCGI_Click()
Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
    
    AppStatus True, "Loading CGI Data..."
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile strConfigFile
    
    Set Node = ConfigXML.SearchAllForContent(Nothing, lstCGI.Text)
    Set Node = Node.GetParent
    
    Node.FindChild2 ("Interpreter")
    txtCGIInterp.Text = Node.Content
    Set Node = Node.GetParent
    
    Node.FindChild2 ("Extension")
    txtCGIExt.Text = Node.Content
    Set Node = Node.GetParent
    
    Set XML = Nothing
    Set ConfigXML = Nothing
    Set Node = Nothing
    AppStatus False
End Sub

Private Sub lstvHosts_Click()
Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
    
    AppStatus True, "Loading vHost Data..."
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile strConfigFile
    
    Set Node = ConfigXML.SearchAllForContent(Nothing, lstvHosts.Text)
    Set Node = Node.GetParent
    
    Node.FindChild2 ("vhName")
    txtvHostName.Text = Node.Content
    Set Node = Node.GetParent
    
    Node.FindChild2 ("vhHostName")
    txtvHostDomain.Text = Node.Content
    Set Node = Node.GetParent
    
    Node.FindChild2 ("vhRoot")
    txtvHostRoot.Text = Node.Content
    Set Node = Node.GetParent
    
    Node.FindChild2 ("vhLogFile")
    txtvHostLog.Text = Node.Content
    Set Node = Node.GetParent
    
    Set XML = Nothing
    Set ConfigXML = Nothing
    Set Node = Nothing
    AppStatus False
End Sub

Private Sub mnuFileReload_Click()
Dim RetVal As Long
    RetVal = MsgBox("This will reset any changes you make." & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion)
    If RetVal = vbYes Then
        If LoadConfigData(strConfigFile) = False Then
            RetVal = MsgBox("There was an error while loading your configuration data." & vbCrLf & vbCrLf & "Press 'Abort' to give up and exit, 'Retry' to try to load th data again," & vbCrLf & "or 'Ignore' to continue.", vbCritical + vbAbortRetryIgnore + vbApplicationModal)
            Select Case RetVal
                Case vbAbort
                    Unload Me
                Case vbRetry
                    If LoadConfigData(strConfigFile) = False Then
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
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "This is going to be an about box someday."
End Sub

Private Function SaveConfigData(strCurConfigFile As String) As Boolean
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       WinUI
' Procedure  :       SaveConfigData
' Description:       this is where we save the changes to the config data.
'
'                    returns true on sucess
'
'                    does nothing for now, i'll fix it later :-P
' Created by :       Adam
' Date-Time  :       8/25/2003-1:12:28 AM
' Parameters :       strCurConfigFile (String)
'--------------------------------------------------------------------------------
'</CSCM>

    MsgBox "If this were a latter version this would save the changes you've made." & vbCrLf & vbCrLf & "But... This isn't a later version and all this does is show this pretty box." & vbCrLf & vbCrLf & ":-P", vbOKOnly + vbInformation

    SaveConfigData = True
End Function

Private Sub tmrStatus_Timer()
Dim strSrvStatusCur As String
    strSrvStatusCur = ServiceStatus("", "SWS Web Server")
    cmdSrvStart.Enabled = True
    cmdSrvStop.Enabled = True
    cmdSrvRestart.Enabled = True
    lblSrvStatusCur.Font.Bold = False
    Select Case strSrvStatusCur
        Case "Stopped"
            lblSrvStatusCur.Caption = "Stopped"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Start Pending"
            lblSrvStatusCur.Caption = "Start Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Stop Pending"
            lblSrvStatusCur.Caption = "Stop Pending"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStop.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Running"
            lblSrvStatusCur.Caption = "Running"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbGreen
            cmdSrvStart.Enabled = False
        Case "Coninue Pending"
            lblSrvStatusCur.Caption = "Coninue Pending"
            lblSrvStatusCur.ForeColor = vbYellow
            cmdSrvStart.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Pause Pending"
            lblSrvStatusCur.Caption = "Pause Pending"
            lblSrvStatusCur.ForeColor = vbRed
            cmdSrvStart.Enabled = False
            cmdSrvRestart.Enabled = False
        Case "Paused"
            lblSrvStatusCur.Caption = "Paused"
            lblSrvStatusCur.Font.Bold = True
            lblSrvStatusCur.ForeColor = vbRed
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

