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
Public strConfigFile As String
Public strStatsFile As String
Public strUIPath As String
Public strAppPath As String
Public strInstalledVer As String
Public Config As tConfig
Public Update As tUpdate
Public Stats As tStats
Public DynDNS As tDynDNS
Public blnRegistered As Boolean
Public blnUseDynDNS As Boolean
'</GlobalVars>

'<LocalVars>
Dim strLang As String
'</LocalVars>

'<LocalTypes>
Private Type tvHost
    Name As String
    Domain As String
    Root As String
    Log As String
End Type

Private Type tConfig
    ServerName As String
    Port As Integer
    WebRoot As String
    MaxConnections As Long
    LogFile As String
    Index() As String
    AllowIndex As String
    CGI() As String
    vHost() As tvHost
    ErrorPages As String
    ListeningAddress As String
    ErrorLog As String
End Type

Private Type tUpdate
    Available As Boolean
    Version As String
    Date As String
    InfoURL As String
    DownloadURL As String
    Description As String
    UpdateLevel As String
    FileSize As Long
End Type

Private Type tStats
    LastRestart As Date
    RequestCount As Long
    TotalBytesSent As Double
End Type

Private Type tDynDNS
    Enabled As Boolean
    CurrentIP As String
    Hostname As String
    UserName As String
    Password As String
    LastUpdate As String
    LastResult As String
    LastIP As String
End Type
'</LocalTypes>

Public Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>
100     LoadUser32 True
104     InitCommonControlsVB
108     Load frmSplash
112     frmSplash.Show
116     DoEvents
120     strUIPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
124     LoadLang
128     If App.PrevInstance = True Then
132         If SetFocusByCaption(GetText("SWEBS Web Server - Control Center")) = False Then
136             DisplayErrMsg "There is already a instance of this application running.", "basMain", , True
            End If
140         End
        End If
144     App.Title = GetText("SWEBS Web Server - Control Center")
148     If GetSWSInstalled = False Then
152         DisplayErrMsg "SWEBS Not detected. You must install SWEBS Web Server to use this application.", "basMain.Main", , True
        End If
156     GetConfigLocation
160     If Dir$(strConfigFile) = "" Then
164         DisplayErrMsg "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", "basMain.Main", , True
        End If
168     blnRegistered = GetRegistered
172     LoadDynDNSData
176     If GetNetStatus = True Then
180         If blnRegistered = False Then
184             StartRegistration
            End If
        End If
188     DoEvents
192     Load frmMain
196     DoEvents
200     frmSplash.Hide
204     DoEvents
208     frmMain.Show
212     Unload frmSplash
216     If LCase(GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
220         Load frmTip
224         frmTip.Show
        End If
        '<EhFooter>
        Exit Sub

Main_Err:
228     DisplayErrMsg Err.Description, "WinUI.basMain.Main", Erl, False
232     Resume Next
        '</EhFooter>
End Sub

Public Sub GetConfigLocation()
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       WinUI
    ' Procedure  :       GetConfigLocation
    ' Description:       Retrives the location of the config. XML file from the registry
    '
    '                    Location of file info:
    '                    HKEY_LOCAL_MACHINE\SOFTWARE\SWS\ConfigFile
    ' Created by :       Adam
    ' Date-Time  :       8/24/2003-1:59:20 PM
    ' Parameters :       none
    '--------------------------------------------------------------------------------
    '</CSCM>
        '<EhHeader>
        On Error GoTo GetConfigLocation_Err
        '</EhHeader>
100     strConfigFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
104     strStatsFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
        '<EhFooter>
        Exit Sub

GetConfigLocation_Err:
108     DisplayErrMsg Err.Description, "WinUI.basMain.GetConfigLocation", Erl, False
112     Resume Next
        '</EhFooter>
End Sub

Public Function GetSWSInstalled() As Boolean
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       WinUI
    ' Procedure  :       GetSWSInstalled
    ' Description:       This will check 2 things, first is to see it SWS is even installed,
    '                    then it will see if the service is installed. If it's not installed
    '                    then it will offer a link to the SWS home page, if the service isnt
    '                    installed, it'll try to install it.
    '
    '                    returns true for a useable installation, false for unusable.
    '
    '                    for now returns true if app path actually exists
    '                    i'll finish this someday, not really a high priority.
    ' Created by :       Adam
    ' Date-Time  :       8/24/2003-2:09:24 PM
    ' Parameters :       none.
    '--------------------------------------------------------------------------------
    '</CSCM>
        '<EhHeader>
        On Error GoTo GetSWSInstalled_Err
        '</EhHeader>

100     strInstalledVer = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
104     strAppPath = GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
108     strAppPath = IIf(Right$(strAppPath, 1) = "\", strAppPath, strAppPath & "\")
112     If Dir$(strAppPath) <> "" Then
116         GetSWSInstalled = True
        Else
120         GetSWSInstalled = False
        End If

        '<EhFooter>
        Exit Function

GetSWSInstalled_Err:
124     DisplayErrMsg Err.Description, "WinUI.basMain.GetSWSInstalled", Erl, False
128     Resume Next
        '</EhFooter>
End Function

Public Function GetConfigData(strCurConfigFile As String) As Boolean
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       WinUI
    ' Procedure  :       GetConfigData
    ' Description:       This loads the data from the config XML file, returns true
    '                    if the load is sucessful, otherwise returns false
    ' Created by :       Adam
    ' Date-Time  :       8/24/2003-3:01:42 PM
    ' Parameters :       strCurConfigFile (String)
    '--------------------------------------------------------------------------------
    '</CSCM>
        '<EhHeader>
        On Error GoTo GetConfigData_Err
        '</EhHeader>

    Dim XML As CHILKATXMLLib.XmlFactory
    Dim ConfigXML As CHILKATXMLLib.IChilkatXml
    Dim Node As CHILKATXMLLib.IChilkatXml
    Dim strTemp As String
    Dim strTemp1() As String
    Dim strTemp2() As String
    Dim strTemp3() As String
    Dim strTemp4() As String
    Dim i As Long
    
100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     ConfigXML.LoadXmlFile strCurConfigFile
    
        '<ServerName>
112     Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
116     If Node Is Nothing Then
120         Config.ServerName = "SWEBS Server"
        Else
124         Config.ServerName = Trim$(Node.Content)
        End If
    
        '<Port>
128     Set Node = ConfigXML.SearchForTag(Nothing, "Port")
132     If Node Is Nothing Then
136         Config.Port = 80
        Else
140         Config.Port = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
        End If
    
        '<Webroot>
144     Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
148     If Node Is Nothing Then
152         strTemp = strAppPath & "Webroot"
        Else
156         strTemp = Trim$(Node.Content)
        End If
160     Config.WebRoot = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<MaxConnections>
164     Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
168     If Node Is Nothing Then
172         Config.MaxConnections = 20
        Else
176         Config.MaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
        End If
    
        '<LogFile>
180     Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
184     If Node Is Nothing Then
188         Config.LogFile = strAppPath & "SWS.log"
        Else
192         Config.LogFile = Trim$(Node.Content)
        End If
    
        '<AllowIndex>
196     Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
200     If Node Is Nothing Then
204         Config.AllowIndex = "false"
        Else
208         Config.AllowIndex = IIf(LCase$(Node.Content) = "true", "true", "false")
        End If
    
        '<ErrorPages>
212     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorPages")
216     If Node Is Nothing Then
220         strTemp = strAppPath & "Errors"
        Else
224         strTemp = Trim$(Node.Content)
        End If
228     Config.ErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<ErrorLog>
232     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorLog")
236     If Node Is Nothing Then
240         Config.ErrorLog = strAppPath & "ErrorLog.log"
        Else
244         Config.ErrorLog = Trim$(Node.Content)
        End If
    
        '<IndexFile>
248     ReDim Config.Index(1 To 1) As String
252     Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
256     If Node Is Nothing Then
260         ReDim Config.Index(1 To 1)
264         Config.Index(1) = "index.html"
        Else
268         Do While Not (Node Is Nothing)
272             If Trim$(Node.Content) <> "" Then
276                 Config.Index(UBound(Config.Index)) = Trim$(Node.Content)
280                 ReDim Preserve Config.Index(1 To (UBound(Config.Index) + 1))
                End If
284             Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
            Loop
288         ReDim Preserve Config.Index(1 To (IIf(UBound(Config.Index) > 1, UBound(Config.Index) - 1, 1)))
        End If
    
        '<VirtualHost>
292     Set Node = ConfigXML.FindChild("VirtualHost")
296     If Not (Node Is Nothing) Then
300         ReDim Config.vHost(1 To 1) As tvHost
304         Do While Not (Node Is Nothing)
308             If Node.GetChildContent("vhName") <> "" Then
312                 Config.vHost(UBound(Config.vHost())).Name = Trim$(Node.GetChildContent("vhName"))
316                 Config.vHost(UBound(Config.vHost())).Domain = Trim$(Node.GetChildContent("vhHostName"))
320                 Config.vHost(UBound(Config.vHost())).Root = Trim$(Node.GetChildContent("vhRoot"))
324                 Config.vHost(UBound(Config.vHost())).Log = Trim$(Node.GetChildContent("vhLogFile"))
                End If
328             Set Node = ConfigXML.SearchForTag(Node, "VirtualHost")
332             If Not (Node Is Nothing) Then
336                 ReDim Preserve Config.vHost(1 To UBound(Config.vHost()) + 1) As tvHost
                End If
            Loop
        Else
340         ReDim Config.vHost(1 To 1)
        End If

        '<CGI>
344     ReDim strTemp1(1 To 1)
348     ReDim strTemp2(1 To 1)
352     Set Node = ConfigXML.FindChild("CGI")
356     If Not (Node Is Nothing) Then
360         Do While Not (Node Is Nothing)
364             If Node.GetChildContent("Interpreter") <> "" Then
368                 strTemp1(UBound(strTemp1)) = Trim$(Node.GetChildContent("Interpreter"))
372                 strTemp2(UBound(strTemp2)) = Trim$(Node.GetChildContent("Extension"))
376                 ReDim Preserve strTemp1(1 To (UBound(strTemp1) + 1))
380                 ReDim Preserve strTemp2(1 To (UBound(strTemp2) + 1))
                End If
384             Set Node = ConfigXML.SearchForTag(Node, "CGI")
            Loop
388         ReDim Config.CGI(1 To (IIf(UBound(strTemp1) > 1, UBound(strTemp1) - 1, 1)), 2) As String
392         For i = 1 To UBound(Config.CGI)
396             Config.CGI(i, 1) = strTemp1(i)
400             Config.CGI(i, 2) = strTemp2(i)
            Next
        Else
404         ReDim Config.CGI(1 To 1, 1 To 2)
        End If
    
        '<ListeningAddress>
408     Set Node = ConfigXML.SearchForTag(Nothing, "ListeningAddress")
412     If Node Is Nothing Then
416         Config.ListeningAddress = ""
        Else
420         Config.ListeningAddress = Node.Content
        End If
    
        'clean up
424     Set XML = Nothing
428     Set ConfigXML = Nothing
432     Set Node = Nothing
436     GetConfigData = True
        '<EhFooter>
        Exit Function

GetConfigData_Err:
440     DisplayErrMsg Err.Description, "WinUI.basMain.GetConfigData", Erl, False
444     Resume Next
        '</EhFooter>
End Function

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
    Dim i As Long

100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     Set ConfigXML2 = XML.NewXml
    
112     Set ConfigXML = ConfigXML.NewChild("sws", "")
116     ConfigXML.NewChild2 "ServerName", Config.ServerName
120     ConfigXML.NewChild2 "Port", Config.Port
124     ConfigXML.NewChild2 "Webroot", IIf(Right$(Config.WebRoot, 1) = "\", Left$(Config.WebRoot, (Len(Config.WebRoot) - 1)), Config.WebRoot)
128     ConfigXML.NewChild2 "ErrorPages", IIf(Right$(Config.ErrorPages, 1) = "\", Left$(Config.ErrorPages, (Len(Config.ErrorPages) - 1)), Config.ErrorPages)
132     ConfigXML.NewChild2 "MaxConnections", Config.MaxConnections
136     ConfigXML.NewChild2 "LogFile", Config.LogFile
140     ConfigXML.NewChild2 "ErrorLog", Config.ErrorLog
144     If Config.ListeningAddress <> "" Then
148         ConfigXML.NewChild2 "ListeningAddress", Config.ListeningAddress
        End If
152     ConfigXML.NewChild2 "AllowIndex", Config.AllowIndex
156     For i = 1 To UBound(Config.Index)
160         ConfigXML.NewChild2 "IndexFile", Config.Index(i)
        Next
164     If Config.CGI(1, 1) <> "" Then
168         For i = 1 To UBound(Config.CGI)
172             Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
176             ConfigXML2.NewChild2 "Interpreter", Config.CGI(i, 1)
180             ConfigXML2.NewChild2 "Extension", Config.CGI(i, 2)
184             ConfigXML.AddChildTree ConfigXML2
            Next
        End If
188     If Config.vHost(1).Name <> "" Then
192         For i = 1 To UBound(Config.vHost)
196             Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
200             ConfigXML2.NewChild2 "vhName", Config.vHost(i).Name
204             ConfigXML2.NewChild2 "vhHostName", Config.vHost(i).Domain
208             ConfigXML2.NewChild2 "vhRoot", Config.vHost(i).Root
212             ConfigXML2.NewChild2 "vhLogFile", Config.vHost(i).Log
216             ConfigXML.AddChildTree ConfigXML2
            Next
        End If
    
        'ConfigXML.SaveXml strUIPath & "test.xml"
220     ConfigXML.SaveXml strCurConfigFile

        'save dns config
224     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", DynDNS.Hostname
228     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", DynDNS.LastIP
232     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", DynDNS.LastResult
236     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", DynDNS.LastUpdate
240     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", DynDNS.Password
244     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", DynDNS.UserName
248     If DynDNS.Enabled = True Then
252         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
        Else
256         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
        End If
    
260     SaveConfigData = True
        '<EhFooter>
        Exit Function

SaveConfigData_Err:
264     DisplayErrMsg Err.Description, "WinUI.basMain.SaveConfigData", Erl, False
268     Resume Next
        '</EhFooter>
End Function

Public Function GetConfigReport() As String
        '<EhHeader>
        On Error GoTo GetConfigReport_Err
        '</EhHeader>
    Dim strReport As String
    Dim strTemp As String
    Dim i As Long

100     strReport = "SWEBS Configuration Report"
104     strReport = strReport & vbCrLf & GetText("Date") & ": " & Now
108     strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
112     strReport = strReport & GetText("Server Name") & ": " & Config.ServerName & vbCrLf
116     strReport = strReport & GetText("Port") & ": & Config.Port & vbCrLf"
120     strReport = strReport & GetText("Web Root") & ": " & Config.WebRoot & vbCrLf
124     strReport = strReport & GetText("Error Pages") & ": " & Config.ErrorPages & vbCrLf
128     strReport = strReport & GetText("Max Connections") & ": " & Config.MaxConnections & vbCrLf
132     strReport = strReport & GetText("Primary Log File") & ": " & Config.LogFile & vbCrLf
136     strReport = strReport & GetText("Allow Index") & ": " & Config.AllowIndex & vbCrLf
140     For i = 1 To UBound(Config.Index)
144         strTemp = strTemp & Config.Index(i) & " "
        Next
148     strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
152     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
156     For i = 1 To UBound(Config.CGI)
160         strReport = strReport & GetText("CGI: Extension") & ": " & Config.CGI(i, 2) & " " & GetText("Interpreter") & ": " & Config.CGI(i, 1) & vbCrLf
        Next
164     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
168     For i = 1 To UBound(Config.vHost)
172         strReport = strReport & GetText("vHost: Name") & ": " & Config.vHost(i).Name & " " & GetText("Host Name") & ": " & Config.vHost(i).Domain & " " & GetText("Root Directory") & ": " & Config.vHost(i).Root & " " & GetText("Log File") & ": " & Config.vHost(i).Log & vbCrLf
        Next
176     GetConfigReport = strReport
        '<EhFooter>
        Exit Function

GetConfigReport_Err:
180     DisplayErrMsg Err.Description, "WinUI.basMain.GetConfigReport", Erl, False
184     Resume Next
        '</EhFooter>
End Function

Public Sub AddNewCGI(strExt As String, strInterp As String)
        '<EhHeader>
        On Error GoTo AddNewCGI_Err
        '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(Config.CGI)), 1 To 2)
104     For i = 1 To UBound(Config.CGI)
108         strTemp1(i, 1) = Config.CGI(i, 1)
112         strTemp1(i, 2) = Config.CGI(i, 2)
        Next
116     ReDim Config.CGI(1 To (UBound(Config.CGI) + 1), 1 To 2)
120     For i = 1 To (UBound(Config.CGI) - 1)
124         Config.CGI(i, 1) = strTemp1(i, 1)
128         Config.CGI(i, 2) = strTemp1(i, 2)
        Next
132     Config.CGI(UBound(Config.CGI), 1) = strInterp
136     Config.CGI(UBound(Config.CGI), 2) = strExt
        '<EhFooter>
        Exit Sub

AddNewCGI_Err:
140     DisplayErrMsg Err.Description, "WinUI.basMain.AddNewCGI", Erl, False
144     Resume Next
        '</EhFooter>
End Sub

Public Sub AddNewvHost(strName As String, strDomain As String, strRoot As String, strLog As String)
        '<EhHeader>
        On Error GoTo AddNewvHost_Err
        '</EhHeader>
100     ReDim Preserve Config.vHost(1 To UBound(Config.vHost()) + 1)
104     Config.vHost(UBound(Config.vHost)).Name = strName
108     Config.vHost(UBound(Config.vHost)).Domain = strDomain
112     Config.vHost(UBound(Config.vHost)).Root = strRoot
116     Config.vHost(UBound(Config.vHost)).Log = strLog
        '<EhFooter>
        Exit Sub

AddNewvHost_Err:
120     DisplayErrMsg Err.Description, "WinUI.basMain.AddNewvHost", Erl, False
124     Resume Next
        '</EhFooter>
End Sub

Public Sub RemoveCGI(lngItem As Long)
        '<EhHeader>
        On Error GoTo RemoveCGI_Err
        '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(Config.CGI)), 1 To 2)
104     For i = 1 To UBound(Config.CGI)
108         strTemp1(i, 1) = Config.CGI(i, 1)
112         strTemp1(i, 2) = Config.CGI(i, 2)
        Next
116     ReDim Config.CGI(1 To (IIf(UBound(Config.CGI) = 1, 1, UBound(Config.CGI) - 1)), 1 To 2)
120     For i = 1 To (lngItem - 1)
124         Config.CGI(i, 1) = strTemp1(i, 1)
128         Config.CGI(i, 2) = strTemp1(i, 2)
        Next
132     For i = (lngItem + 1) To (UBound(strTemp1))
136         Config.CGI(i - 1, 1) = strTemp1(i, 1)
140         Config.CGI(i - 1, 2) = strTemp1(i, 2)
        Next
        '<EhFooter>
        Exit Sub

RemoveCGI_Err:
144     DisplayErrMsg Err.Description, "WinUI.basMain.RemoveCGI", Erl, False
148     Resume Next
        '</EhFooter>
End Sub

Public Sub RemovevHost(lngItem As Long)
        '<EhHeader>
        On Error GoTo RemovevHost_Err
        '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(Config.vHost)), 1 To 4)
104     For i = 1 To UBound(Config.vHost)
108         strTemp1(i, 1) = Config.vHost(i).Name
112         strTemp1(i, 2) = Config.vHost(i).Domain
116         strTemp1(i, 3) = Config.vHost(i).Root
120         strTemp1(i, 4) = Config.vHost(i).Log
        Next
124     ReDim Config.vHost(1 To (IIf(UBound(Config.vHost) = 1, 1, UBound(Config.vHost) - 1)))
128     For i = 1 To (lngItem - 1)
132         Config.vHost(i).Name = strTemp1(i, 1)
136         Config.vHost(i).Domain = strTemp1(i, 2)
140         Config.vHost(i).Root = strTemp1(i, 3)
144         Config.vHost(i).Log = strTemp1(i, 4)
        Next
148     For i = lngItem + 1 To (UBound(strTemp1))
152         Config.vHost(i - 1).Name = strTemp1(i, 1)
156         Config.vHost(i - 1).Domain = strTemp1(i, 2)
160         Config.vHost(i - 1).Root = strTemp1(i, 3)
164         Config.vHost(i - 1).Log = strTemp1(i, 4)
        Next
        '<EhFooter>
        Exit Sub

RemovevHost_Err:
168     DisplayErrMsg Err.Description, "WinUI.basMain.RemovevHost", Erl, False
172     Resume Next
        '</EhFooter>
End Sub

Public Sub GetUpdateStatus(strdata As String)
        '<EhHeader>
        On Error GoTo GetUpdateStatus_Err
        '</EhHeader>
    Dim strNewVer() As String
    Dim strCurVer() As String
    Dim i As Long

100     If InStr(1, strdata, "Server at swebs.sourceforge.net Port 80") = 0 And strdata <> "" Then
104         Update.Date = GetTaggedData(strdata, "Date")
108         Update.Description = GetTaggedData(strdata, "Description")
112         Update.DownloadURL = GetTaggedData(strdata, "DownloadURL")
116         Update.InfoURL = GetTaggedData(strdata, "InfoURL")
120         Update.Version = GetTaggedData(strdata, "Version")
124         Update.UpdateLevel = GetTaggedData(strdata, "UpgradeLevel")
128         Update.FileSize = Val(GetTaggedData(strdata, "FileSize"))
        
            'check to see if this is newer
132         strNewVer() = Split(Update.Version, ".")
136         strCurVer() = Split(strInstalledVer, ".")
140         For i = 0 To UBound(strNewVer)
144             If Val(strNewVer(i)) > Val(strCurVer(i)) Then
148                 Update.Available = True
                End If
            Next
152     ElseIf Update.Version <> "" Then
156         Update.Available = True
        Else
160         Update.Available = False
        End If
        '<EhFooter>
        Exit Sub

GetUpdateStatus_Err:
164     DisplayErrMsg Err.Description, "WinUI.basMain.GetUpdateStatus", Erl, False
168     Resume Next
        '</EhFooter>
End Sub

Public Sub GetStatsData()
        '<EhHeader>
        On Error GoTo GetStatsData_Err
        '</EhHeader>
    Dim XML As CHILKATXMLLib.XmlFactory
    Dim StatsXML As CHILKATXMLLib.IChilkatXml
    Dim Node As CHILKATXMLLib.IChilkatXml
    
100     Set XML = New XmlFactory
104     Set StatsXML = XML.NewXml
108     If Dir$(strStatsFile) <> "" And strStatsFile <> "" Then
112         StatsXML.LoadXmlFile strStatsFile
        End If
    
        '<TotalBytesSent>
116     Set Node = StatsXML.SearchForTag(Nothing, "TotalBytesSent")
120     If Node Is Nothing Then
124         Stats.TotalBytesSent = 0
        Else
128         Stats.TotalBytesSent = Node.Content
        End If
    
        '<LastRestart>
132     Set Node = StatsXML.SearchForTag(Nothing, "LastRestart")
136     If Node Is Nothing Then
140         Stats.LastRestart = CDate(Now)
        Else
144         Stats.LastRestart = CDate(Node.Content)
        End If
    
        '<RequestCount>
148     Set Node = StatsXML.SearchForTag(Nothing, "RequestCount")
152     If Node Is Nothing Then
156         Stats.RequestCount = 0
        Else
160         Stats.RequestCount = Val(Node.Content)
        End If
    
        'clean up
164     Set XML = Nothing
168     Set StatsXML = Nothing
172     Set Node = Nothing
        '<EhFooter>
        Exit Sub

GetStatsData_Err:
176     DisplayErrMsg Err.Description, "WinUI.basMain.GetStatsData", Erl, False
180     Resume Next
        '</EhFooter>
End Sub

Public Function GetRegistered() As Boolean
    'Dim strResult As String
    '    strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")
    '    If strResult <> "" Then
    '        GetRegistered = True
    '    Else
    '        GetRegistered = False
    '    End If
        '<EhHeader>
        On Error GoTo GetRegistered_Err
        '</EhHeader>
    
100     GetRegistered = True
        '<EhFooter>
        Exit Function

GetRegistered_Err:
104     DisplayErrMsg Err.Description, "WinUI.basMain.GetRegistered", Erl, False
108     Resume Next
        '</EhFooter>
End Function

Public Sub StartRegistration()
        '<EhHeader>
        On Error GoTo StartRegistration_Err
        '</EhHeader>
    Dim lngResult As Long
100     lngResult = MsgBox(GetText("Would you like to register your software? It's fast and Free!\r\rProduct registration is used to provide the best possible service, products, and support for our users.\rWe will not contact you nor will we sell or give away any of your information.\r\rWould you like to register now?"), vbQuestion + vbYesNo + vbApplicationModal)
104     If lngResult = vbYes Then
108         Load frmRegistration
112         frmRegistration.Show vbModal
        End If
        '<EhFooter>
        Exit Sub

StartRegistration_Err:
116     DisplayErrMsg Err.Description, "WinUI.basMain.StartRegistration", Erl, False
120     Resume Next
        '</EhFooter>
End Sub

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
120     DisplayErrMsg Err.Description, "WinUI.basMain.GetText", Erl, False
124     Resume Next
        '</EhFooter>
End Function

Private Sub LoadLang()
        '<EhHeader>
        On Error GoTo LoadLang_Err
        '</EhHeader>
    Dim strLangTemp As String
    Dim lngLen As String

100     If Dir$(strUIPath & "lang.xml") <> "" Then
104         lngLen = FileLen(strUIPath & "lang.xml")
108         strLangTemp = Space$(lngLen)
112         Open strUIPath & "lang.xml" For Binary As 1 Len = lngLen
116             Get #1, 1, strLangTemp
120         Close 1
124         strLang = GetTaggedData(strLangTemp, "1033")
128         strLang = Trim$(strLang)
132         strLang = Replace(strLang, vbCrLf, "")
136         strLang = Replace(strLang, Chr$(9), "")
        End If
        '<EhFooter>
        Exit Sub

LoadLang_Err:
140     DisplayErrMsg Err.Description, "WinUI.basMain.LoadLang", Erl, False
144     Resume Next
        '</EhFooter>
End Sub

Public Sub LoadDynDNSData()
        '<EhHeader>
        On Error GoTo LoadDynDNSData_Err
        '</EhHeader>
    Dim strResult As String
    
100     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSEnable")
104     If LCase(strResult) = "true" Then
108         DynDNS.Enabled = True
        Else
112         DynDNS.Enabled = False
        End If
116     DynDNS.Hostname = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSHostname")
120     DynDNS.LastIP = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastIP")
124     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastResult")
128     If strResult = "" Then
132         DynDNS.LastResult = "(None)"
        Else
136         DynDNS.LastResult = strResult
        End If
140     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastUpdate")
144     If strResult = "" Then
148         DynDNS.LastUpdate = CDate(2.00001)
        Else
152         DynDNS.LastUpdate = CDate(strResult)
        End If
156     DynDNS.Password = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSPassword")
160     DynDNS.UserName = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSUsername")
        '<EhFooter>
        Exit Sub

LoadDynDNSData_Err:
164     DisplayErrMsg Err.Description, "WinUI.basMain.LoadDynDNSData", Erl, False
168     Resume Next
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
116     If blnFatal = True Then
120         End
        End If
        '<EhFooter>
        Exit Sub

DisplayErrMsg_Err:
124     DisplayErrMsg Err.Description, "WinUI.basMain.DisplayErrMsg", Erl, False
128     Resume Next
        '</EhFooter>
End Sub
