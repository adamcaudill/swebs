VERSION 5.00
Begin VB.Form frmEventView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SWEBS Web Server - Control Center Event Viewer"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmEventView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCallStack 
      Height          =   1335
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Timer tmrEvents 
      Interval        =   500
      Left            =   120
      Top             =   3480
   End
   Begin VB.TextBox txtEvents 
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmEventView"
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

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmEventView.Form_Load")
    '</EhHeader>
100     WinUI.EventLog.Enabled = True
104     WinUI.EventLog.AddEvent "SWEBS_WinUI_Main.frmEventView.Form_Load", "Event Viewer Loaded"
108     Form_Resize
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_Load_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmEventView.Form_Load", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    txtEvents.Move 0, 0, (Me.ScaleWidth), (Me.ScaleHeight) - 1500
    txtCallStack.Move 0, (Me.ScaleHeight - 1400), Me.ScaleWidth, 1400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.frmEventView.Form_Unload")
    '</EhHeader>
100     WinUI.EventLog.Enabled = False
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Form_Unload_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.frmEventView.Form_Unload", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub tmrEvents_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
Dim strCallStack As String
Dim i As Long

    If WinUI.EventLog.Changed = True Then
        txtEvents.Text = WinUI.EventLog.Log
        txtEvents.SelStart = Len(txtEvents.Text)
    End If
    strCallStack = "Current Call Stack:" & vbCrLf
    If WinUI.Debuger.CallStack.Count >= 1 Then
        For i = 1 To WinUI.Debuger.CallStack.Count
            strCallStack = strCallStack & Chr(9) & WinUI.Debuger.CallStack.Peek(i) & vbCrLf
        Next
    Else
        strCallStack = strCallStack & Chr(9) & "(None)"
    End If
    txtCallStack.Text = strCallStack
End Sub
