Attribute VB_Name = "basMain"
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

Public WinUI As cWinUI

'CLI Option variables
Dim blnNoSplash As Boolean
Dim blnTrayOnly As Boolean 'currently unused, to be added arter the move to .net
Dim blnNoTips As Boolean

Public Sub Main()
    GetArgs Command()
    Set WinUI = New cWinUI
    If blnNoSplash <> True Then
        Load frmSplash
        frmSplash.Show
        frmSplash.Refresh
    End If
    If App.PrevInstance = True Then
        If WinUI.Util.SetFocusByCaption(WinUI.GetTranslatedText("SWEBS Web Server - Control Center")) = False Then
            MsgBox "There is already a instance of this application running.", vbApplicationModal + vbCritical
            End
         End If
        End
     End If
    App.Title = WinUI.GetTranslatedText("SWEBS Web Server - Control Center")
    If Dir$(WinUI.ConfigFile) = "" Then
        MsgBox "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", vbApplicationModal + vbCritical
        End
    End If
    SetStatus "Checking For Registration Data..."
    If WinUI.Net.IsOnline = True Then
        If WinUI.Registration.IsRegistered = False Then
            SetStatus "Starting Registration..."
            WinUI.Registration.Start
         End If
     End If
    Load frmMain
    If blnNoSplash <> True Then
        Unload frmSplash
        DoEvents
    End If
    frmMain.Show
    If blnNoTips <> True Then
        If LCase$(WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
            Load frmTip
            frmTip.Show vbModal
        End If
    End If
End Sub

Public Sub UnloadApp()
Dim i As Long

    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next
    WinUI.Util.LoadUser32 False
    Set WinUI = Nothing
    End
End Sub

Public Sub SetStatus(strStatus As String, Optional blnBusy As Boolean = False)
    If IsLoaded("SWEBS-Splash") = True Then
        frmSplash.lblStatus.Caption = strStatus
        frmSplash.Refresh
    ElseIf IsLoaded("SWEBS Web Server - Control Center") = True Then
        If blnBusy = True Then
            Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
        Else
            Screen.MousePointer = vbNormal  '0 default
        End If
        frmMain.lblAppStatus.Caption = strStatus
        frmMain.Refresh
    End If
    WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cDialog.SetStatus", "App Status Message: " & strStatus
    DoEvents
End Sub

Private Sub GetArgs(strCommand As String)
Dim strArgs() As String
Dim i As Long

    strArgs = Split(strCommand, " ")
    For i = 0 To UBound(strArgs)
        Select Case strArgs(i)
            Case "--nosplash"
                blnNoSplash = True
            Case "--debuglang"
                WinUI.Debuger.DebugLang = True
            Case "--tray"
                blnTrayOnly = True
            Case "--notips"
                blnNoTips = True
            Case Else
                MsgBox "Unknown Argument: " & strArgs(i), vbApplicationModal + vbCritical
                UnloadApp
        End Select
    Next
End Sub

Public Function IsLoaded(strCaption As String) As Boolean
Dim i As Long

    If Forms(i).Caption = strCaption Then
        IsLoaded = True
    End If
End Function
