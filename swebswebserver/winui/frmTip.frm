VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   3720
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   120
      Picture         =   "frmTip.frx":0CCA
      ScaleHeight     =   3015
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblDidYouKnow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
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

Private Sub chkLoadTipsAtStartup_Click()
    '<EhHeader>
    On Error GoTo chkLoadTipsAtStartup_Click_Err
    '</EhHeader>
100     If chkLoadTipsAtStartup.Value = vbChecked Then
104         SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODEnable", "true"
108         WinUI.EventLog.AddEvent "WinUI.frmTip.chkLoadTipsAtStartup_Click", "TOD Enabled"
        Else
112         SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODEnable", "false"
116         WinUI.EventLog.AddEvent "WinUI.frmTip.chkLoadTipsAtStartup_Click", "TOD Disabled"
        End If
    '<EhFooter>
    Exit Sub

chkLoadTipsAtStartup_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmTip.chkLoadTipsAtStartup_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNextTip_Click()
    '<EhHeader>
    On Error GoTo cmdNextTip_Click_Err
    '</EhHeader>
100     GetTip
    '<EhFooter>
    Exit Sub

cmdNextTip_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmTip.cmdNextTip_Click", Erl, False
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
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmTip.cmdOK_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub GetTip()
    '<EhHeader>
    On Error GoTo GetTip_Err
    '</EhHeader>
    Dim strTOD As String
    Dim lngCurTip As Long

100     If Dir$(WinUI.Path & "tips.xml") <> "" Then
104         strTOD = Space$(FileLen(WinUI.Path & "tips.xml"))
108         Open WinUI.Path & "tips.xml" For Binary As 1
112             Get #1, 1, strTOD
116         Close 1
120         lngCurTip = Val(GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODCurrent"))
124         lngCurTip = lngCurTip + 1
128         If lngCurTip > GetTaggedData(strTOD, "Count") Then
132             lngCurTip = 1
            End If
136         SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODCurrent", Trim$(Str$(lngCurTip))
140         strTOD = GetTaggedData(strTOD, Trim$(Str$(lngCurTip)))
144         lblTitle = GetTaggedData(strTOD, "Title")
148         lblTipText = CUnescape(GetTaggedData(strTOD, "TipText"))
152         WinUI.EventLog.AddEvent "WinUI.frmTip.GetTip", "Loaded tip #" & lngCurTip & " (" & lblTitle.Caption & ")"
        Else
156         MsgBox WinUI.GetTranslatedText("TOD XML Data File Not Found."), vbCritical + vbApplicationModal
160         WinUI.EventLog.AddEvent "WinUI.frmTip.GetTip", "Tips.xml not found."
164         Unload Me
        End If
    '<EhFooter>
    Exit Sub

GetTip_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmTip.GetTip", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     GetTip
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmTip.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub
