VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stovell Web Server - Control Center"
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIndexFiles 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox txtAllowIndex 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox txtLogFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtMaxConnect 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtWebroot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtServerName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblIndexFiles 
      Caption         =   "Index files"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label lblAllowIndex 
      Caption         =   "Allow index"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label lblLogFile 
      Caption         =   "Log File"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblMaxConnect 
      Caption         =   "How many simultainious connections do you want to accept?"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblWebroot 
      Caption         =   "Where is the root folder for your server?"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblPort 
      Caption         =   "What port do you want to use? (Default is 80)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblServerName 
      Caption         =   "What is the name of your server?"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
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
Option Explicit

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

Private Sub Form_Load()
Dim RetVal As Long
    If LoadConfigData(strConfigFile) = False Then
        RetVal = MsgBox("There was an error while loading your configuration data." & vbCrLf & vbCrLf & "Press 'Abort' to give up and exit, 'Retry' to try to load th data again," & vbCrLf & "or 'Ignore' to continue.", vbCritical + vbAbortRetryIgnore + vbApplicationModal)
        Select Case RetVal
            Case vbAbort
                Unload Me
            Case vbRetry
                If LoadConfigData(strConfigFile) = False Then
                    MsgBox "A second attempt to load your configuration data failed. Aborting." & vbCrLf & vbCrLf & "This application will now close.", vbApplicationModal + vbCritical
                End If
            Case vbIgnore
                MsgBox "NOTICE: You have chosen to proceed after a data error, this application may" & vbCrLf & "not function properly or you may loose data."
        End Select
    End If
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
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    Do While Not (Node Is Nothing)
        txtIndexFiles.Text = txtIndexFiles.Text & Node.Content & " "
        Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
    Loop
    
    LoadConfigData = True
End Function
