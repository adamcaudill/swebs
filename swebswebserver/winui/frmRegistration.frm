VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegistration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  SWEBS Web Server - Registration"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet netReg 
      Left            =   5520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   2513
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cmbUse 
      Height          =   315
      ItemData        =   "frmRegistration.frx":0000
      Left            =   240
      List            =   "frmRegistration.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cmbExpiriance 
      Height          =   315
      ItemData        =   "frmRegistration.frx":0054
      Left            =   240
      List            =   "frmRegistration.frx":0064
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtFindUs 
      Height          =   285
      Left            =   240
      MaxLength       =   128
      TabIndex        =   7
      Top             =   2400
      Width           =   5775
   End
   Begin VB.ComboBox cmbWhere 
      Height          =   315
      ItemData        =   "frmRegistration.frx":0088
      Left            =   240
      List            =   "frmRegistration.frx":0098
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox cmbComputers 
      Height          =   315
      ItemData        =   "frmRegistration.frx":00CC
      Left            =   240
      List            =   "frmRegistration.frx":00DF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   240
      MaxLength       =   128
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblUse 
      Caption         =   "What will you use this software for?"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblExpiriance 
      Caption         =   "How much computer experience do you have?"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label lblFindUs 
      Caption         =   "How did you find out about us?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblWhere 
      Caption         =   "Where are you using this?"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblComputers 
      Caption         =   "How Many Computers Do You Own?"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblEMail 
      Caption         =   "What is your e-mail address? (We will not contact you, this is simply used to track installations)."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmRegistration"
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

Private Sub cmdSubmit_Click()
    '<EhHeader>
    On Error GoTo cmdSubmit_Click_Err
    '</EhHeader>
    Dim strResult As String
    Dim strQuery As String
 
100     If txtEmail.Text = "" Then
104         MsgBox GetText("You must provide a e-mail address."), vbInformation + vbApplicationModal + vbOKOnly
108         txtEmail.SetFocus
112         WinUI.EventLog.AddEvent "WinUI.frmRegistration.cmdSubmit_Click", "User did not enter email address."
            Exit Sub
        End If
    
116     Me.MousePointer = vbHourglass
120     cmdSubmit.Enabled = False
124     txtEmail.Enabled = False
128     cmbComputers.Enabled = False
132     cmbWhere.Enabled = False
136     txtFindUs.Enabled = False
140     cmbExpiriance.Enabled = False
144     cmbUse.Enabled = False
    
148     strQuery = "?email=" & UrlEncode(txtEmail.Text) & "&ccount=" & UrlEncode(cmbComputers.Text) & "&where=" & UrlEncode(cmbWhere.Text) & "&find=" & UrlEncode(txtFindUs.Text) & "&exp=" & UrlEncode(cmbExpiriance.Text) & "&use=" & UrlEncode(cmbUse.Text) & "&ver=" & UrlEncode(WinUI.Version)
152     strResult = netReg.OpenURL("http://swebs.sf.net/register/reginit.php" & strQuery)
    
156     Me.Hide
160     Select Case strResult
            Case "Completed"
164             Call SaveRegistryString(&H80000002, "SOFTWARE\SWS", "RegID", txtEmail.Text)
168             WinUI.EventLog.AddEvent "WinUI.frmRegistration.cmdSubmit_Click", "Registration completed."
172             WinUI.Registered = True
176         Case "Duplicate"
180             MsgBox GetText("You have already registered, you only need to register once."), vbApplicationModal + vbInformation + vbOKOnly
184             Call SaveRegistryString(&H80000002, "SOFTWARE\SWS", "RegID", txtEmail.Text)
188             WinUI.EventLog.AddEvent "WinUI.frmRegistration.cmdSubmit_Click", "Registration duplicate."
192             WinUI.Registered = True
196         Case Else
200             MsgBox GetText("There was a unknown error. Registration Failed./r/rThe Registration server returned the following information:\r") & strResult
204             WinUI.EventLog.AddEvent "WinUI.frmRegistration.cmdSubmit_Click", "Registration failed."
        End Select
208     Unload Me
    '<EhFooter>
    Exit Sub

cmdSubmit_Click_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmRegistration.cmdSubmit_Click", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
100     lblEMail.Caption = GetText("What is your e-mail address? (We will not contact you, this is simply used to track installations).")
104     lblComputers.Caption = GetText("How Many Computers Do You Own?")
108     lblWhere.Caption = GetText("Where are you using this?")
112     lblFindUs.Caption = GetText("How did you find out about us?")
116     lblExpiriance.Caption = GetText("How much computer experience do you have?")
120     lblUse.Caption = GetText("What will you use this software for?")
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmRegistration.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub
