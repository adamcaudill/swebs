VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: WinUI Custom
'***************************************************************************
'
' SWEBS/WinUI
'
' Copyright (c) Martin Richardson.
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

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private t As NOTIFYICONDATA
Private WithEvents pichook As PictureBox
Attribute pichook.VB_VarHelpID = -1
Private mvarToolTip As String

Public Event LButtonDblClk()
Public Event RButtonUp()

Private mvarSourceWindow As Form 'local copy

Public Property Let ToolTip(ByVal vData As String)
    ChangeToolTip vData
End Property

Public Property Let Icon(ByVal vData As Variant)
    ChangeIcon vData
End Property

Public Property Get Icon() As Variant
    Icon = t.hIcon      'pichook.Picture
End Property

Public Property Set SourceWindow(ByVal vData As Form)
    Set mvarSourceWindow = vData
    SetPicHook
End Property

Private Sub Class_Initialize()
    t.cbSize = Len(t)
    t.uID = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.uCallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = Chr$(0)
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Rec As Boolean, msg As Long, oldmsg As Long
    
    oldmsg = msg
    msg = X / Screen.TwipsPerPixelX
   
    If Rec = False Then
        Rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                RaiseEvent LButtonDblClk
            Case WM_RBUTTONUP:
                RaiseEvent RButtonUp
        End Select
        Rec = False
    End If
End Sub

Public Sub RemoveFromSysTray()
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uID = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub IconInSysTray()
    Shell_NotifyIcon NIM_ADD, t
End Sub

Private Sub SetPicHook()
On Error GoTo AlreadyAdded
    Set pichook = mvarSourceWindow.Controls.Add("VB.PictureBox", "pichook")

    pichook.Visible = False
    pichook.Picture = mvarSourceWindow.Icon
    t.hwnd = pichook.hwnd
    
    Exit Sub

AlreadyAdded:
    If Err.Number <> 727 Then  ' pichook has already been added
       MsgBox "Run-time error '" & Err.Number & "':" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
       Resume
    End If
End Sub

Public Sub ChangeIcon(toNewIcon)
    Set pichook.Picture = toNewIcon
    t.hIcon = pichook.Picture
    Shell_NotifyIcon NIM_MODIFY, t
End Sub

Public Sub ChangeToolTip(ByVal cNewTip As String)
    mvarToolTip = cNewTip
    t.szTip = cNewTip & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, t
End Sub


