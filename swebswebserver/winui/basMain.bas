Attribute VB_Name = "basMain"
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

Public WinUI As cWinUI

Public g_oMenuHook As cHookingThunk
Public g_oMenuHookImpl As cMenuHook
Public g_oCurrentMenu As ctxHookMenu

'CLI Option variables
Dim blnNoSplash As Boolean
Dim blnTrayOnly As Boolean
Dim blnNoTips As Boolean
Dim blnDebugLang As Boolean
Dim blnNoUpdate As Boolean
Dim blnKillUpdate As Boolean
Dim blnDebugMode As Boolean

'CSEH: WinUI - Custom(No Stack)
Public Sub Main()
    '<EhHeader>
    On Error GoTo Main_Err
    '</EhHeader>
100     SetExceptionFilter True
104     GetArgs Command$()
108     Set WinUI = New cWinUI
112     WinUI.Setup
116     If blnDebugLang = True Then WinUI.Debuger.DebugLang = True
120     If blnNoSplash = True Then WinUI.Debuger.DisableSplash = True
124     If blnNoTips = True Then WinUI.Debuger.DisableTips = True
128     If blnNoUpdate = True Then WinUI.Debuger.DisableUpdate = True
132     If blnKillUpdate = True Then WinUI.Debuger.KillUpdate
136     If blnDebugMode = True Then WinUI.Debuger.DebugMode = True
    
140     If WinUI.Debuger.DisableSplash <> True Then
144         Load frmSplash
148         WinUI.Util.FormFade frmSplash, False
        End If
152     If App.PrevInstance = True Then
156         If WinUI.Util.SetFocusByCaption(WinUI.GetTranslatedText("SWEBS Web Server - Control Center")) = False Then
160             MsgBox "There is already a instance of this application running.", vbApplicationModal + vbCritical
164             End
             End If
168         End
         End If
172     App.Title = WinUI.GetTranslatedText("SWEBS Web Server - Control Center")
176     If Dir$(WinUI.Server.HTTP.Config.File) = "" Then
180         MsgBox "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", vbApplicationModal + vbCritical
184         End
        End If
188     SetStatus "Checking For Registration Data..."
192     If WinUI.Net.IsOnline = True Then
196         If WinUI.Registration.IsRegistered = False Then
200             SetStatus "Starting Registration..."
204             WinUI.Registration.Start
             End If
         End If
208     Load frmMain
212     If WinUI.Debuger.DisableSplash <> True Then
216         WinUI.Util.FormFade frmSplash, True
220         Unload frmSplash
224         DoEvents
        End If
228     If blnTrayOnly <> True Then
232         frmMain.Show
        End If
236     If WinUI.Debuger.DisableTips <> True Then
240         If LCase$(WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
244             Load frmTip
248             frmTip.Show vbModal
            End If
        End If
    '<EhFooter>
    Exit Sub

Main_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.Main", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub SetStatus(strStatus As String, Optional blnBusy As Boolean = False)
    '<EhHeader>
    On Error GoTo SetStatus_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.basMain.SetStatus")
    '</EhHeader>
100     If IsLoaded("SWEBS-Splash") = True Then
104         If frmSplash.lblStatus.Caption <> strStatus Then
108             frmSplash.lblStatus.Caption = strStatus
112             frmSplash.Refresh
            End If
116     ElseIf IsLoaded("SWEBS Web Server - Control Center") = True Then
120         If frmMain.lblAppStatus.Caption <> strStatus Then
124             If blnBusy = True Then
128                 Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
                Else
132                 Screen.MousePointer = vbNormal  '0 default
                End If
136             frmMain.lblAppStatus.Caption = strStatus
140             frmMain.Refresh
            End If
        End If
144     WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cDialog.SetStatus", "App Status Message: " & strStatus
148     DoEvents
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

SetStatus_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.SetStatus", Erl, False
    Resume Next
    '</EhFooter>
End Sub

'CSEH: WinUI - Custom(No Stack)
Private Sub GetArgs(strCommand As String)
    '<EhHeader>
    On Error GoTo GetArgs_Err
    '</EhHeader>
    Dim strArgs() As String
    Dim i As Long

100     strArgs = Split(strCommand, " ")
104     For i = 0 To UBound(strArgs)
108         Select Case strArgs(i)
                Case "--nosplash"
112                 blnNoSplash = True
116             Case "--debuglang"
120                 blnDebugLang = True
124             Case "--tray"
128                 blnTrayOnly = True
132             Case "--notips"
136                 blnNoTips = True
140             Case "--noupdate"
144                 blnNoUpdate = True
148             Case "--killupdate"
152                 blnKillUpdate = True
156             Case "--debug"
160                 blnDebugMode = True
164             Case Else
168                 MsgBox "Unknown Argument: " & strArgs(i) & vbCrLf & vbCrLf & "Valid arguments are:" & vbCrLf & "--nosplash" & vbCrLf & "--debuglang" & vbCrLf & "--tray" & vbCrLf & "--notips" & vbCrLf & "--noupdate" & vbCrLf & "--killupdate" & vbCrLf & "--debug", vbApplicationModal + vbCritical
172                 End
            End Select
        Next
    '<EhFooter>
    Exit Sub

GetArgs_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetArgs", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function IsLoaded(strCaption As String) As Boolean
    '<EhHeader>
    On Error GoTo IsLoaded_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.basMain.IsLoaded")
    '</EhHeader>
    Dim i As Long

100     For i = 0 To Forms.Count - 1
104         If Forms(i).Caption = strCaption Then
108             IsLoaded = True
            End If
        Next
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

IsLoaded_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.IsLoaded", Erl, False
    Resume Next
    '</EhFooter>
End Function

'CSEH: ErrResumeNext
Public Sub DisplayErrMsg(strMessage As String, strLocation As String, Optional strLine As String = "(Unknown)", Optional blnFatal As Boolean = False)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
Dim strErrMsg As String
Dim strErrReport As String
Dim lngRetVal As Long

    If strMessage = "" Then
        strMessage = "There was an unknown error."
    End If
    If WinUI Is Nothing Then
        strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "")
        MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
    Else
        If WinUI.Debuger.DebugMode = True Then
            strErrReport = WinUI.Debuger.ErrorReport(strMessage, strLine, strLocation)
            strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "") & vbCrLf & vbCrLf & "An error log has been written to:" & vbCrLf & strErrReport
            MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
        Else
            strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "") & vbCrLf & vbCrLf & "Would you like to create an error log?"
            lngRetVal = MsgBox(strErrMsg, vbApplicationModal + vbCritical + vbYesNo, "SWEBS System Error")
            If lngRetVal = vbYes Then
                strErrReport = WinUI.Debuger.ErrorReport(strMessage, strLine, strLocation)
                MsgBox "An error log has been written to:" & vbCrLf & strErrReport, vbInformation + vbApplicationModal
            End If
        End If
    End If
    If blnFatal = True Then
        End
    End If
End Sub
