VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.PictureBox picTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   120
      Picture         =   "frmTip.frx":0CCA
      ScaleHeight     =   3045
      ScaleWidth      =   3705
      TabIndex        =   0
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
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblDidYouKnow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   180
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Label lblNextTip 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Next Tip"
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
      Left            =   4335
      MouseIcon       =   "frmTip.frx":0FD4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
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
      Left            =   4545
      MouseIcon       =   "frmTip.frx":1126
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "frmTip"
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

Private Sub chkLoadTipsAtStartup_Click()
    If chkLoadTipsAtStartup.Value = vbChecked Then
        Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODEnable", "true"
        Core.EventLog.AddEvent "SWEBS_Core_Main.frmTip.chkLoadTipsAtStartup_Click", "TOD Enabled"
    Else
        Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODEnable", "false"
        Core.EventLog.AddEvent "SWEBS_Core_Main.frmTip.chkLoadTipsAtStartup_Click", "TOD Disabled"
    End If
End Sub

Private Sub lblNextTip_Click()
    GetTip
End Sub

Private Sub lblOK_Click()
    Unload Me
End Sub

Private Sub GetTip()
Dim strTOD As String
Dim lngCurTip As Long

    If Dir$(Core.Path & "tips.xml") <> "" Then
        strTOD = Space$(FileLen(Core.Path & "tips.xml"))
        Open Core.Path & "tips.xml" For Binary As 1
            Get #1, 1, strTOD
        Close 1
        lngCurTip = Val(Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODCurrent"))
        lngCurTip = lngCurTip + 1
        If lngCurTip > Util.GetTaggedData(strTOD, "Count") Then
            lngCurTip = 1
        End If
        Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "TODCurrent", Trim$(Str$(lngCurTip))
        strTOD = Util.GetTaggedData(strTOD, Trim$(Str$(lngCurTip)))
        lblTitle = Util.GetTaggedData(strTOD, "Title")
        lblTipText = Util.CUnescape(Util.GetTaggedData(strTOD, "TipText"))
        Core.EventLog.AddEvent "SWEBS_Core_Main.frmTip.GetTip", "Loaded tip #" & lngCurTip & " (" & lblTitle.Caption & ")"
    Else
        MsgBox Translator.GetText("TOD XML Data File Not Found."), vbCritical + vbApplicationModal
        Core.EventLog.AddEvent "SWEBS_Core_Main.frmTip.GetTip", "Tips.xml not found."
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    GetTip
End Sub
