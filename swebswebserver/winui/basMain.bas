Attribute VB_Name = "basMain"
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

'FadeForm
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'FadeForm
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Public Core As cCore
Public Util As cUtil
Public Translator As cTranslate

'CLI Option variables
Dim blnNoSplash As Boolean
Dim blnTrayOnly As Boolean
Dim blnNoTips As Boolean
Dim blnDebugLang As Boolean
Dim blnNoUpdate As Boolean
Dim blnKillUpdate As Boolean
Dim blnDebugMode As Boolean
Dim blnPerfMon As Boolean

'CSEH: Core - Custom(No Stack)
Public Sub Main()
Dim UIInt As cUIInterface

    SetExceptionFilter True
    
    'Create a Core from the default instance
    Set UIInt = New cUIInterface
    Set Core = UIInt.DefInstance
    Set UIInt = Nothing
    
    'create a instance of the util class
    Set Util = New cUtil
    
    'create a new instance of the translator
    Set Translator = New cTranslate
    
    GetArgs Command$()
    If blnDebugLang = True Then Core.Debuger.DebugLang = True
    If blnNoUpdate = True Then Core.Debuger.DisableUpdate = True
    If blnKillUpdate = True Then Core.Debuger.KillUpdate
    If blnDebugMode = True Then Core.Debuger.DebugMode = True
    If blnPerfMon = True Then Core.Debuger.PerfMon.Enabled = True
    
    If blnNoSplash <> True Then
        Load frmSplash
        FormFade frmSplash, False
    End If
    If App.PrevInstance = True Then
        If Util.SetFocusByCaption(Translator.GetText("SWEBS Web Server - Control Center")) = False Then
            MsgBox Translator.GetText("There is already a instance of this application running."), vbApplicationModal + vbCritical
            End
         End If
        End
     End If
    App.Title = Translator.GetText("SWEBS Web Server - Control Center")
    If Dir$(Core.Server.HTTP.Config.File) = "" Then
        MsgBox Translator.GetText("Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file."), vbApplicationModal + vbCritical
        End
    End If
    SetStatus Translator.GetText("Checking For Registration Data") & "..."
    If Core.Net.IsOnline = True Then
        If Core.Registration.IsRegistered = False Then
            SetStatus Translator.GetText("Starting Registration") & "..."
            Core.Registration.Start '<- This is not UI split
         End If
     End If
    Load frmMain
    If blnNoSplash <> True Then
        FormFade frmSplash, True
        Unload frmSplash
        DoEvents
    End If
    If blnTrayOnly <> True Then
        frmMain.Show
    End If
    If blnNoTips <> True Then
        If LCase$(Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
            Load frmTip
            frmTip.Show vbModal
        End If
    End If
End Sub

Public Sub SetStatus(strStatus As String, Optional blnBusy As Boolean = False)
    If IsLoaded("SWEBS-Splash") = True Then
        If frmSplash.lblStatus.Caption <> strStatus Then
            frmSplash.lblStatus.Caption = strStatus
            frmSplash.Refresh
        End If
    ElseIf IsLoaded(Translator.GetText("SWEBS Web Server - Control Center")) = True Then
        If frmMain.lblAppStatus.Caption <> strStatus Then
            If blnBusy = True Then
                Screen.MousePointer = vbArrowHourglass '13 arrow + hourglass
            Else
                Screen.MousePointer = vbNormal  '0 default
            End If
            frmMain.lblAppStatus.Caption = strStatus
            frmMain.Refresh
        End If
    End If
    Core.EventLog.AddEvent "SWEBS_Core_DLL.cDialog.SetStatus", "App Status Message: " & strStatus
    DoEvents
End Sub

'CSEH: Core - Custom(No Stack)
Private Sub GetArgs(strCommand As String)
Dim strArgs() As String
Dim i As Long

    strArgs = Split(strCommand, " ")
    For i = 0 To UBound(strArgs)
        Select Case strArgs(i)
            Case "--nosplash"               'Disables the splash screen
                blnNoSplash = True
            Case "--debuglang"              'should set a flag to show a message box is a string is not found
                blnDebugLang = True
            Case "--tray"                   'shows tray icon only, does not show frmMain
                blnTrayOnly = True
            Case "--notips"                 'doesn't show the TOD window, also disables the feature for the future.
                blnNoTips = True
            Case "--noupdate"               'doesn't run the updated check on start up.
                blnNoUpdate = True
            Case "--killupdate"             'permenently disables the update feature
                blnKillUpdate = True
            Case "--debug"                  '???
                blnDebugMode = True
            Case "--perfmon"                'enables a speed log for all functions
                blnPerfMon = True
            Case Else
                MsgBox "Unknown Argument: " & strArgs(i) & vbCrLf & vbCrLf & "Valid arguments are:" & vbCrLf & "--nosplash" & vbCrLf & "--debuglang" & vbCrLf & "--tray" & vbCrLf & "--notips" & vbCrLf & "--noupdate" & vbCrLf & "--killupdate" & vbCrLf & "--debug" & vbCrLf & "--perfmon", vbApplicationModal + vbCritical
                End
        End Select
    Next
End Sub

Public Function IsLoaded(strCaption As String) As Boolean
Dim i As Long

    For i = 0 To Forms.Count - 1
        If Forms(i).Caption = strCaption Then
            IsLoaded = True
        End If
    Next
End Function

'CSEH: ErrResumeNext
Public Sub DisplayErrMsg(strMessage As String, strLocation As String, Optional strLine As String = "(Unknown)", Optional blnFatal As Boolean = False)
Dim strErrMsg As String
Dim strErrReport As String
Dim lngRetVal As Long

    If strMessage = "" Then
        strMessage = Translator.GetText("There was an unknown error.")
    End If
    If Core Is Nothing Then
        strErrMsg = "This application has encountered a error:" & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "")
        MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
    Else
        If Core.Debuger.DebugMode = True Then
            strErrReport = Core.Debuger.ErrorReport(strMessage, strLine, strLocation)
            strErrMsg = Translator.GetText("This application has encountered a error:\n\nError:") & " '" & strMessage & "'" & vbCrLf & Translator.GetText("Location:") & " " & strLocation & " " & Translator.GetText("at line:") & " " & strLine & vbCrLf & vbCrLf & Translator.GetText("Contact ADAM@IMSPIRE.COM to report this error.") & IIf(blnFatal = True, vbCrLf & vbCrLf & Translator.GetText("This error is fatal, this program will now close."), "") & vbCrLf & vbCrLf & Translator.GetText("An error log has been written to:") & vbCrLf & strErrReport
            MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, Translator.GetText("SWEBS System Error")
        Else
            strErrMsg = Translator.GetText("This application has encountered a error:\n\nError:") & " '" & strMessage & "'" & vbCrLf & Translator.GetText("Location:") & " " & strLocation & " " & Translator.GetText("at line:") & " " & strLine & vbCrLf & vbCrLf & Translator.GetText("Contact ADAM@IMSPIRE.COM to report this error.") & IIf(blnFatal = True, vbCrLf & vbCrLf & Translator.GetText("This error is fatal, this program will now close."), "") & vbCrLf & vbCrLf & Translator.GetText("Would you like to create an error log?")
            lngRetVal = MsgBox(strErrMsg, vbApplicationModal + vbCritical + vbYesNo, Translator.GetText("SWEBS System Error"))
            If lngRetVal = vbYes Then
                strErrReport = Core.Debuger.ErrorReport(strMessage, strLine, strLocation)
                MsgBox Translator.GetText("An error log has been written to:") & vbCrLf & strErrReport, vbInformation + vbApplicationModal
            End If
        End If
    End If
    If blnFatal = True Then
        End
    End If
End Sub

Public Function FormFade(ByRef frmForm As Form, blnHide As Boolean) As Long
Dim MSG As Long
Dim i As Long

    If blnHide = True Then
        For i = 255 To 0 Step -5
            'Set window style to layered
            MSG = GetWindowLong(frmForm.hwnd, GWL_EXSTYLE)
            MSG = MSG Or WS_EX_LAYERED
            SetWindowLong frmForm.hwnd, GWL_EXSTYLE, MSG
            'Set the opacity of the layer according the the parameters
            SetLayeredWindowAttributes frmForm.hwnd, 0, i, LWA_ALPHA
            frmForm.Refresh
        Next
    Else
        frmForm.Show
        For i = 0 To 255 Step 5
            'Set window style to layered
            MSG = GetWindowLong(frmForm.hwnd, GWL_EXSTYLE)
            MSG = MSG Or WS_EX_LAYERED
            SetWindowLong frmForm.hwnd, GWL_EXSTYLE, MSG
            'Set the opacity of the layer according the the parameters
            SetLayeredWindowAttributes frmForm.hwnd, 0, i, LWA_ALPHA
            frmForm.Refresh
        Next
    End If
End Function
