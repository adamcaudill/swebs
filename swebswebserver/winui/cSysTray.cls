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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CSEH: WinUI - Custom
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
    '<EhHeader>
    On Error GoTo ToolTip_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.ToolTip")
    '</EhHeader>
100     ChangeToolTip vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ToolTip_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.ToolTip", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Icon(ByVal vData As Variant)
    '<EhHeader>
    On Error GoTo Icon_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.Icon")
    '</EhHeader>
100     ChangeIcon vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Icon_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.Icon", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Icon() As Variant
    '<EhHeader>
    On Error GoTo Icon_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.Icon")
    '</EhHeader>
100     Icon = t.hIcon      'pichook.Picture
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Icon_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.Icon", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Set SourceWindow(ByVal vData As Form)
    '<EhHeader>
    On Error GoTo SourceWindow_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.SourceWindow")
    '</EhHeader>
100     Set mvarSourceWindow = vData
104     SetPicHook
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

SourceWindow_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.SourceWindow", Erl, False
    Resume Next
    '</EhFooter>
End Property

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.Class_Initialize")
    '</EhHeader>
100     t.cbSize = Len(t)
104     t.uID = 1&
108     t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
112     t.uCallbackMessage = WM_MOUSEMOVE
116     t.hIcon = Me.Icon
120     t.szTip = Chr$(0)
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo pichook_MouseMove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.pichook_MouseMove")
    '</EhHeader>
    Static Rec As Boolean, MSG As Long, oldmsg As Long
    
100     oldmsg = MSG
104     MSG = X / Screen.TwipsPerPixelX
   
108     If Rec = False Then
112         Rec = True
116         Select Case MSG
                Case WM_LBUTTONDBLCLK:
120                 RaiseEvent LButtonDblClk
                Case WM_RBUTTONUP:
124                 RaiseEvent RButtonUp
            End Select
128         Rec = False
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

pichook_MouseMove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.pichook_MouseMove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub RemoveFromSysTray()
    '<EhHeader>
    On Error GoTo RemoveFromSysTray_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.RemoveFromSysTray")
    '</EhHeader>
100     t.cbSize = Len(t)
104     t.hwnd = pichook.hwnd
108     t.uID = 1&
112     Shell_NotifyIcon NIM_DELETE, t
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

RemoveFromSysTray_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.RemoveFromSysTray", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub IconInSysTray()
    '<EhHeader>
    On Error GoTo IconInSysTray_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.IconInSysTray")
    '</EhHeader>
100     Shell_NotifyIcon NIM_ADD, t
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

IconInSysTray_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.IconInSysTray", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub SetPicHook()
    '<EhHeader>
    On Error GoTo SetPicHook_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.SetPicHook")
    '</EhHeader>
100     Set pichook = mvarSourceWindow.Controls.Add("VB.PictureBox", "pichook")

104     pichook.Visible = False
108     pichook.Picture = mvarSourceWindow.Icon
112     t.hwnd = pichook.hwnd
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

SetPicHook_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.SetPicHook", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub ChangeIcon(toNewIcon)
    '<EhHeader>
    On Error GoTo ChangeIcon_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.ChangeIcon")
    '</EhHeader>
100     Set pichook.Picture = toNewIcon
104     t.hIcon = pichook.Picture
108     Shell_NotifyIcon NIM_MODIFY, t
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

ChangeIcon_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.ChangeIcon", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub ChangeToolTip(ByVal cNewTip As String)
    '<EhHeader>
    On Error GoTo ChangeToolTip_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cSysTray.ChangeToolTip")
    '</EhHeader>
100     mvarToolTip = cNewTip
104     t.szTip = cNewTip & Chr$(0)
108     Shell_NotifyIcon NIM_MODIFY, t
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

ChangeToolTip_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cSysTray.ChangeToolTip", Erl, False
    Resume Next
    '</EhFooter>
End Sub


