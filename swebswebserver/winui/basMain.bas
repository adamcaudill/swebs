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
128     LoadLang
132     If App.PrevInstance = True Then
136         If SetFocusByCaption(GetText("SWEBS Web Server - Control Center")) = False Then
140             DisplayErrMsg "There is already a instance of this application running.", "basMain", , True
             End If
144         End
         End If
148     App.Title = GetText("SWEBS Web Server - Control Center")
152     If Dir$(WinUI.ConfigFile) = "" Then
156         DisplayErrMsg "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", "basMain.Main", , True
         End If
160     SplashStatus "Checking For Registration Data..."
164     WinUI.DynDNS.Reload
168     If WinUI.Net.IsOnline = True Then
172         If WinUI.Registration.IsRegistered = False Then
176             SplashStatus "Starting Registration..."
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

Public Function SaveConfigData(strCurConfigFile As String) As Boolean
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       WinUI
    ' Procedure  :       SaveConfigData
    ' Description:       this is where we save the changes to the config data.
    '
    '                    returns true on sucess
    ' Created by :       Adam
    ' Date-Time  :       8/25/2003-1:12:28 AM
    ' Parameters :       strCurConfigFile (String)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo SaveConfigData_Err
    '</EhHeader>
    Dim XML As CHILKATXMLLib.XmlFactory
    Dim ConfigXML As CHILKATXMLLib.IChilkatXml
    Dim ConfigXML2 As CHILKATXMLLib.IChilkatXml
    Dim vItem As Variant
    Dim i As Long

100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     Set ConfigXML2 = XML.NewXml
    
112     Set ConfigXML = ConfigXML.NewChild("sws", "")
116     ConfigXML.NewChild2 "ServerName", WinUI.Config.ServerName
120     ConfigXML.NewChild2 "Port", WinUI.Config.Port
124     ConfigXML.NewChild2 "Webroot", IIf(Right$(WinUI.Config.WebRoot, 1) = "\", Left$(WinUI.Config.WebRoot, (Len(WinUI.Config.WebRoot) - 1)), WinUI.Config.WebRoot)
128     ConfigXML.NewChild2 "ErrorPages", IIf(Right$(WinUI.Config.ErrorPages, 1) = "\", Left$(WinUI.Config.ErrorPages, (Len(WinUI.Config.ErrorPages) - 1)), WinUI.Config.ErrorPages)
132     ConfigXML.NewChild2 "MaxConnections", WinUI.Config.MaxConnections
136     ConfigXML.NewChild2 "LogFile", WinUI.Config.LogFile
140     ConfigXML.NewChild2 "ErrorLog", WinUI.Config.ErrorLog
144     If WinUI.Config.ListeningAddress <> "" Then
148         ConfigXML.NewChild2 "ListeningAddress", WinUI.Config.ListeningAddress
        End If
152     ConfigXML.NewChild2 "AllowIndex", WinUI.Config.AllowIndex
    
156     For Each vItem In WinUI.Config.Index
160         ConfigXML.NewChild2 "IndexFile", vItem.FileName
        Next
    
164     For Each vItem In WinUI.Config.CGI
168         Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
172         ConfigXML2.NewChild2 "Interpreter", vItem.Interpreter
176         ConfigXML2.NewChild2 "Extension", vItem.Extension
180         ConfigXML.AddChildTree ConfigXML2
        Next
    
184     For Each vItem In WinUI.Config.vHost
188         Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
192         ConfigXML2.NewChild2 "vhName", vItem.HostName
196         ConfigXML2.NewChild2 "vhHostName", vItem.Domain
200         ConfigXML2.NewChild2 "vhRoot", vItem.Root
204         ConfigXML2.NewChild2 "vhLogFile", vItem.Log
208         ConfigXML.AddChildTree ConfigXML2
        Next
    
        'ConfigXML.SaveXml strUIPath & "test.xml"
212     WinUI.EventLog.AddEvent "WinUI.basMain.SaveConfigData", "Saving XML Config File To: " & strCurConfigFile
216     ConfigXML.SaveXml strCurConfigFile

        'save dns config
220     WinUI.DynDNS.Save WinUI.DynDNS.HostName, WinUI.DynDNS.LastIP, WinUI.DynDNS.LastResult, WinUI.DynDNS.LastUpdate, WinUI.DynDNS.Password, WinUI.DynDNS.UserName
    
224     SaveConfigData = True
    '<EhFooter>
    Exit Function

SaveConfigData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.SaveConfigData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function GetConfigReport() As String
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       GetConfigReport
    ' Description:       This returns a nice pretty formatted report about the
    '                       project settings.
    ' Created by :       Adam
    ' Date-Time  :       9/30/2003-1:29:14 AM
    '
    ' Parameters :
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo GetConfigReport_Err
    '</EhHeader>
    Dim strReport As String
    Dim strTemp As String
    Dim vItem As Variant
    Dim i As Long

100     strReport = "SWEBS Configuration Report"
104     strReport = strReport & vbCrLf & GetText("Date") & ": " & Now
108     strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
112     strReport = strReport & GetText("Server Name") & ": " & WinUI.Config.ServerName & vbCrLf
116     strReport = strReport & GetText("Port") & ": & WinUI.Config.Port & vbCrLf"
120     strReport = strReport & GetText("Web Root") & ": " & WinUI.Config.WebRoot & vbCrLf
124     strReport = strReport & GetText("Error Pages") & ": " & WinUI.Config.ErrorPages & vbCrLf
128     strReport = strReport & GetText("Max Connections") & ": " & WinUI.Config.MaxConnections & vbCrLf
132     strReport = strReport & GetText("Primary Log File") & ": " & WinUI.Config.LogFile & vbCrLf
136     strReport = strReport & GetText("Allow Index") & ": " & WinUI.Config.AllowIndex & vbCrLf
    
140     For Each vItem In WinUI.Config.Index
144         strTemp = strTemp & vItem.FileName & " "
        Next
    
148     strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
152     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    
156     For Each vItem In WinUI.Config.CGI
160         strReport = strReport & GetText("CGI: Extension") & ": " & vItem.Extension & " " & GetText("Interpreter") & ": " & vItem.Interpreter & vbCrLf
        Next
    
164     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    
168     For Each vItem In WinUI.Config.vHost
172         strReport = strReport & GetText("vHost: Name") & ": " & vItem.HostName & " " & GetText("Host Name") & ": " & vItem.Domain & " " & GetText("Root Directory") & ": " & vItem.Root & " " & GetText("Log File") & ": " & vItem.Log & vbCrLf
        Next
    
176     GetConfigReport = strReport
    '<EhFooter>
    Exit Function

GetConfigReport_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetConfigReport", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function GetText(strString As String) As String
    '<EhHeader>
    On Error GoTo GetText_Err
    '</EhHeader>
    Dim strResult As String

100     strResult = GetTaggedData(strLang, strString)
104     strResult = CUnescape(strResult)
108     If strResult <> "" Then
112         GetText = strResult
        Else
116         GetText = CUnescape(strString)
        End If
    '<EhFooter>
    Exit Function

GetText_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetText", Erl, False
    Resume Next
    '</EhFooter>
End Function

Private Sub LoadLang()
    '<EhHeader>
    On Error GoTo LoadLang_Err
    '</EhHeader>
    Dim strLangTemp As String

100     If Dir$(WinUI.Path & "lang.xml") <> "" Then
104         strLangTemp = Space$(FileLen(WinUI.Path & "lang.xml"))
108         Open WinUI.Path & "lang.xml" For Binary As 1
112             Get #1, 1, strLangTemp
116         Close 1
120         strLang = GetTaggedData(strLangTemp, "1033")
124         If strLang <> "" Then
128             WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Loaded lang: 1033"
            Else
132             WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Failed to load lang: 1033"
            End If
136         strLang = Trim$(strLang)
140         strLang = Replace(strLang, vbCrLf, "")
144         strLang = Replace(strLang, Chr$(9), "")
        Else
148         WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Lang.xml file is missing."
        End If
    '<EhFooter>
    Exit Sub

LoadLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.LoadLang", Erl, False
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

Public Sub SplashStatus(strStatus As String)
    '<EhHeader>
    On Error GoTo SplashStatus_Err
    '</EhHeader>
    Dim i As Long

100     For i = 0 To Forms.Count - 1
104         If Forms(i).Caption = "SWEBS-Splash" Then
108             frmSplash.lblStatus.Caption = strStatus
112             DoEvents
            End If
        Next
    '<EhFooter>
    Exit Sub

SplashStatus_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.SplashStatus", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub UnloadApp()
    '<EhHeader>
    On Error GoTo UnloadApp_Err
    '</EhHeader>
    Dim i As Long

100     For i = Forms.Count - 1 To 0 Step -1
104         Unload Forms(i)
        Next
108     LoadUser32 False
112     SetExceptionFilter False
116     End
    '<EhFooter>
    Exit Sub

UnloadApp_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.UnloadApp", Erl, False
    Resume Next
    '</EhFooter>
End Sub
