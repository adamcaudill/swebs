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

'<GlobalVars>
Public WinUI As cWinUI
'</GlobalVars>

'<LocalVars>
Dim strLang As String
'</LocalVars>

Public Sub Main()
    '<EhHeader>
    On Error GoTo Main_Err
    '</EhHeader>
100     SetExceptionFilter True
104     LoadUser32 True
108     InitCommonControlsVB
112     Set WinUI = New cWinUI
116     Load frmSplash
120     frmSplash.Show
124     frmSplash.Refresh
132     If App.PrevInstance = True Then
136         If SetFocusByCaption(WinUI.GetTranslatedText("SWEBS Web Server - Control Center")) = False Then
140             DisplayErrMsg "There is already a instance of this application running.", "basMain", , True
             End If
144         End
         End If
148     App.Title = WinUI.GetTranslatedText("SWEBS Web Server - Control Center")
152     If Dir$(WinUI.ConfigFile) = "" Then
156         DisplayErrMsg "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", "basMain.Main", , True
         End If
160     WinUI.Status "Checking For Registration Data..."
164     WinUI.DynDNS.Reload
168     If WinUI.Net.IsOnline = True Then
172         If WinUI.Registration.IsRegistered = False Then
176             WinUI.Status "Starting Registration..."
180             WinUI.Registration.Start
             End If
         End If
184     Load frmMain
188     frmSplash.Hide
192     DoEvents
196     frmMain.Show
200     Unload frmSplash
204     If LCase$(GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
208         Load frmTip
212         frmTip.Show vbModal
         End If
    '<EhFooter>
    Exit Sub

Main_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.Main", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub DisplayErrMsg(strMessage As String, strLocation As String, Optional strLine As String = "(Unknown)", Optional blnFatal As Boolean = False)
    '<EhHeader>
    On Error GoTo DisplayErrMsg_Err
    '</EhHeader>
    Dim strErrMsg As String

100     If strMessage = "" Then
104         strMessage = "There was an unknown error."
        End If
108     strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "")
112     MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
116     WinUI.EventLog.AddEvent "WinUI.basMain.DisplayErrMsg", "An error message was raised. The message was: " & strMessage
120     If blnFatal = True Then
124         End
        End If
    '<EhFooter>
    Exit Sub

DisplayErrMsg_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.DisplayErrMsg", Erl, False
    Resume Next
    '</EhFooter>
End Sub
