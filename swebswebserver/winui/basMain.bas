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
Public WinUI As tWinUI
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
    CGI() As String 'this needs to be converted to a UDT
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

Private Type tWinUI
    ConfigFile As String
    StatsFile As String
    Path As String
    Version As String
    Config As tConfig
    Update As tUpdate
    DynDNS As tDynDNS
    Stats As cStats
    EventLog As cEventLog
    Registration As cRegistration
End Type
'</LocalTypes>

Public Sub Main()
    '<EhHeader>
    On Error GoTo Main_Err
    '</EhHeader>
100     SetExceptionFilter True
104     LoadUser32 True
108     InitCommonControlsVB
112     Set WinUI.EventLog = New cEventLog
116     Set WinUI.Stats = New cStats
117     Set WinUI.Registration = New cRegistration
120     Load frmSplash
124     frmSplash.Show
128     frmSplash.Refresh
132     LoadLang
136     If App.PrevInstance = True Then
140         If SetFocusByCaption(GetText("SWEBS Web Server - Control Center")) = False Then
144             DisplayErrMsg "There is already a instance of this application running.", "basMain", , True
             End If
148         End
         End If
152     App.Title = GetText("SWEBS Web Server - Control Center")
156     If GetSWSInstalled = False Then
160         DisplayErrMsg "SWEBS Not detected. You must install SWEBS Web Server to use this application.", "basMain.Main", , True
         End If
164     GetConfigLocation
168     If Dir$(WinUI.ConfigFile) = "" Then
172         DisplayErrMsg "Your configuration file could not be found. Please re-install the SWEBS Web Server to replace your configuration file.", "basMain.Main", , True
         End If
176     SplashStatus "Checking For Registration Data..."
184     LoadDynDNSData
188     If GetNetStatus = True Then
192         If WinUI.Registration.IsRegistered = False Then
196             SplashStatus "Starting Registration..."
200             WinUI.Registration.Start
             End If
         End If
204     Load frmMain
208     frmSplash.Hide
212     DoEvents
216     frmMain.Show
220     Unload frmSplash
224     If LCase(GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
228         Load frmTip
232         frmTip.Show vbModal
         End If
    '<EhFooter>
    Exit Sub

Main_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.Main", Erl, False
    Resume Next
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
100     WinUI.ConfigFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
104     WinUI.StatsFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
    '<EhFooter>
    Exit Sub

GetConfigLocation_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetConfigLocation", Erl, False
    Resume Next
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

100     WinUI.Version = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
104     WinUI.Path = GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
108     WinUI.Path = IIf(Right$(WinUI.Path, 1) = "\", WinUI.Path, WinUI.Path & "\")
112     If Dir$(WinUI.Path) <> "" Then
116         GetSWSInstalled = True
        Else
120         GetSWSInstalled = False
        End If

    '<EhFooter>
    Exit Function

GetSWSInstalled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetSWSInstalled", Erl, False
    Resume Next
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
    
100     WinUI.EventLog.AddEvent "WinUI.basMain.GetConfigData", "Loading Config Data."
104     Set XML = New XmlFactory
108     Set ConfigXML = XML.NewXml
112     WinUI.EventLog.AddEvent "basMain.GetConfigData", "Loading file: " & strCurConfigFile
116     ConfigXML.LoadXmlFile strCurConfigFile
    
        '<ServerName>
120     Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
124     If Node Is Nothing Then
128         WinUI.Config.ServerName = "SWEBS Server"
        Else
132         WinUI.Config.ServerName = Trim$(Node.Content)
        End If
    
        '<Port>
136     Set Node = ConfigXML.SearchForTag(Nothing, "Port")
140     If Node Is Nothing Then
144         WinUI.Config.Port = 80
        Else
148         WinUI.Config.Port = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
        End If
    
        '<Webroot>
152     Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
156     If Node Is Nothing Then
160         strTemp = WinUI.Path & "Webroot"
        Else
164         strTemp = Trim$(Node.Content)
        End If
168     WinUI.Config.WebRoot = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<MaxConnections>
172     Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
176     If Node Is Nothing Then
180         WinUI.Config.MaxConnections = 20
        Else
184         WinUI.Config.MaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
        End If
    
        '<LogFile>
188     Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
192     If Node Is Nothing Then
196         WinUI.Config.LogFile = WinUI.Path & "SWS.log"
        Else
200         WinUI.Config.LogFile = Trim$(Node.Content)
        End If
    
        '<AllowIndex>
204     Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
208     If Node Is Nothing Then
212         WinUI.Config.AllowIndex = "false"
        Else
216         WinUI.Config.AllowIndex = IIf(LCase$(Node.Content) = "true", "true", "false")
        End If
    
        '<ErrorPages>
220     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorPages")
224     If Node Is Nothing Then
228         strTemp = WinUI.Path & "Errors"
        Else
232         strTemp = Trim$(Node.Content)
        End If
236     WinUI.Config.ErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<ErrorLog>
240     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorLog")
244     If Node Is Nothing Then
248         WinUI.Config.ErrorLog = WinUI.Path & "ErrorLog.log"
        Else
252         WinUI.Config.ErrorLog = Trim$(Node.Content)
        End If
    
        '<IndexFile>
256     ReDim WinUI.Config.Index(1 To 1) As String
260     Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
264     If Node Is Nothing Then
268         ReDim WinUI.Config.Index(1 To 1)
272         WinUI.Config.Index(1) = "index.html"
        Else
276         Do While Not (Node Is Nothing)
280             If Trim$(Node.Content) <> "" Then
284                 WinUI.Config.Index(UBound(WinUI.Config.Index)) = Trim$(Node.Content)
288                 ReDim Preserve WinUI.Config.Index(1 To (UBound(WinUI.Config.Index) + 1))
                End If
292             Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
            Loop
296         ReDim Preserve WinUI.Config.Index(1 To (IIf(UBound(WinUI.Config.Index) > 1, UBound(WinUI.Config.Index) - 1, 1)))
        End If
    
        '<VirtualHost>
300     Set Node = ConfigXML.FindChild("VirtualHost")
304     If Not (Node Is Nothing) Then
308         ReDim WinUI.Config.vHost(1 To 1) As tvHost
312         Do While Not (Node Is Nothing)
316             If Node.GetChildContent("vhName") <> "" Then
320                 WinUI.Config.vHost(UBound(WinUI.Config.vHost())).Name = Trim$(Node.GetChildContent("vhName"))
324                 WinUI.Config.vHost(UBound(WinUI.Config.vHost())).Domain = Trim$(Node.GetChildContent("vhHostName"))
328                 WinUI.Config.vHost(UBound(WinUI.Config.vHost())).Root = Trim$(Node.GetChildContent("vhRoot"))
332                 WinUI.Config.vHost(UBound(WinUI.Config.vHost())).Log = Trim$(Node.GetChildContent("vhLogFile"))
                End If
336             Set Node = ConfigXML.SearchForTag(Node, "VirtualHost")
340             If Not (Node Is Nothing) Then
344                 ReDim Preserve WinUI.Config.vHost(1 To UBound(WinUI.Config.vHost()) + 1) As tvHost
                End If
            Loop
        Else
348         ReDim WinUI.Config.vHost(1 To 1)
        End If

        '<CGI>
352     ReDim strTemp1(1 To 1)
356     ReDim strTemp2(1 To 1)
360     Set Node = ConfigXML.FindChild("CGI")
364     If Not (Node Is Nothing) Then
368         Do While Not (Node Is Nothing)
372             If Node.GetChildContent("Interpreter") <> "" Then
376                 strTemp1(UBound(strTemp1)) = Trim$(Node.GetChildContent("Interpreter"))
380                 strTemp2(UBound(strTemp2)) = Trim$(Node.GetChildContent("Extension"))
384                 ReDim Preserve strTemp1(1 To (UBound(strTemp1) + 1))
388                 ReDim Preserve strTemp2(1 To (UBound(strTemp2) + 1))
                End If
392             Set Node = ConfigXML.SearchForTag(Node, "CGI")
            Loop
396         ReDim WinUI.Config.CGI(1 To (IIf(UBound(strTemp1) > 1, UBound(strTemp1) - 1, 1)), 2) As String
400         For i = 1 To UBound(WinUI.Config.CGI)
404             WinUI.Config.CGI(i, 1) = strTemp1(i)
408             WinUI.Config.CGI(i, 2) = strTemp2(i)
            Next
        Else
412         ReDim WinUI.Config.CGI(1 To 1, 1 To 2)
        End If
    
        '<ListeningAddress>
416     Set Node = ConfigXML.SearchForTag(Nothing, "ListeningAddress")
420     If Node Is Nothing Then
424         WinUI.Config.ListeningAddress = ""
        Else
428         WinUI.Config.ListeningAddress = Node.Content
        End If
    
        'clean up
432     Set XML = Nothing
436     Set ConfigXML = Nothing
440     Set Node = Nothing
444     GetConfigData = True
    '<EhFooter>
    Exit Function

GetConfigData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetConfigData", Erl, False
    Resume Next
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
156     For i = 1 To UBound(WinUI.Config.Index)
160         ConfigXML.NewChild2 "IndexFile", WinUI.Config.Index(i)
        Next
164     If WinUI.Config.CGI(1, 1) <> "" Then
168         For i = 1 To UBound(WinUI.Config.CGI)
172             Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
176             ConfigXML2.NewChild2 "Interpreter", WinUI.Config.CGI(i, 1)
180             ConfigXML2.NewChild2 "Extension", WinUI.Config.CGI(i, 2)
184             ConfigXML.AddChildTree ConfigXML2
            Next
        End If
188     If WinUI.Config.vHost(1).Name <> "" Then
192         For i = 1 To UBound(WinUI.Config.vHost)
196             Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
200             ConfigXML2.NewChild2 "vhName", WinUI.Config.vHost(i).Name
204             ConfigXML2.NewChild2 "vhHostName", WinUI.Config.vHost(i).Domain
208             ConfigXML2.NewChild2 "vhRoot", WinUI.Config.vHost(i).Root
212             ConfigXML2.NewChild2 "vhLogFile", WinUI.Config.vHost(i).Log
216             ConfigXML.AddChildTree ConfigXML2
            Next
        End If
    
        'ConfigXML.SaveXml strUIPath & "test.xml"
220     WinUI.EventLog.AddEvent "WinUI.basMain.SaveConfigData", "Saving XML Config File To: " & strCurConfigFile
224     ConfigXML.SaveXml strCurConfigFile

        'save dns config
228     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", WinUI.DynDNS.Hostname
232     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", WinUI.DynDNS.LastIP
236     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", WinUI.DynDNS.LastResult
240     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", WinUI.DynDNS.LastUpdate
244     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", WinUI.DynDNS.Password
248     SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", WinUI.DynDNS.UserName
252     If WinUI.DynDNS.Enabled = True Then
256         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
        Else
260         SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
        End If
    
264     SaveConfigData = True
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
140     For i = 1 To UBound(WinUI.Config.Index)
144         strTemp = strTemp & WinUI.Config.Index(i) & " "
        Next
148     strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
152     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
156     For i = 1 To UBound(WinUI.Config.CGI)
160         strReport = strReport & GetText("CGI: Extension") & ": " & WinUI.Config.CGI(i, 2) & " " & GetText("Interpreter") & ": " & WinUI.Config.CGI(i, 1) & vbCrLf
        Next
164     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
168     For i = 1 To UBound(WinUI.Config.vHost)
172         strReport = strReport & GetText("vHost: Name") & ": " & WinUI.Config.vHost(i).Name & " " & GetText("Host Name") & ": " & WinUI.Config.vHost(i).Domain & " " & GetText("Root Directory") & ": " & WinUI.Config.vHost(i).Root & " " & GetText("Log File") & ": " & WinUI.Config.vHost(i).Log & vbCrLf
        Next
176     GetConfigReport = strReport
    '<EhFooter>
    Exit Function

GetConfigReport_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetConfigReport", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Sub AddNewCGI(strExt As String, strInterp As String)
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       AddNewCGI
    ' Description:       adds a new item to the WinUI.Config.CGI array
    '
    '                    this needs to be updated to the WinUI.Config.vHost style
    '                    array, since thats a much prettier way of doing things
    ' Created by :       Adam
    ' Date-Time  :       9/30/2003-1:30:48 AM
    '
    ' Parameters :       strExt (String)
    '                    strInterp (String)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo AddNewCGI_Err
    '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(WinUI.Config.CGI)), 1 To 2)
104     For i = 1 To UBound(WinUI.Config.CGI)
108         strTemp1(i, 1) = WinUI.Config.CGI(i, 1)
112         strTemp1(i, 2) = WinUI.Config.CGI(i, 2)
        Next
116     ReDim WinUI.Config.CGI(1 To (UBound(WinUI.Config.CGI) + 1), 1 To 2)
120     For i = 1 To (UBound(WinUI.Config.CGI) - 1)
124         WinUI.Config.CGI(i, 1) = strTemp1(i, 1)
128         WinUI.Config.CGI(i, 2) = strTemp1(i, 2)
        Next
132     WinUI.Config.CGI(UBound(WinUI.Config.CGI), 1) = strInterp
136     WinUI.Config.CGI(UBound(WinUI.Config.CGI), 2) = strExt
    '<EhFooter>
    Exit Sub

AddNewCGI_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.AddNewCGI", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub AddNewvHost(strName As String, strDomain As String, strRoot As String, strLog As String)
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       AddNewvHost
    ' Description:       adds a new item to the WinUI.Config.vHost array
    ' Created by :       Adam
    ' Date-Time  :       9/30/2003-1:33:45 AM
    ' Parameters :       strName (String)
    '                    strDomain (String)
    '                    strRoot (String)
    '                    strLog (String)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo AddNewvHost_Err
    '</EhHeader>
100     ReDim Preserve WinUI.Config.vHost(1 To UBound(WinUI.Config.vHost()) + 1)
104     WinUI.Config.vHost(UBound(WinUI.Config.vHost)).Name = strName
108     WinUI.Config.vHost(UBound(WinUI.Config.vHost)).Domain = strDomain
112     WinUI.Config.vHost(UBound(WinUI.Config.vHost)).Root = strRoot
116     WinUI.Config.vHost(UBound(WinUI.Config.vHost)).Log = strLog
    '<EhFooter>
    Exit Sub

AddNewvHost_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.AddNewvHost", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub RemoveCGI(lngItem As Long)
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       RemoveCGI
    ' Description:       removes an item from the WinUI.Config.CGI aray
    '
    '                    this needs to be updated..
    ' Created by :       Adam
    ' Date-Time  :       9/30/2003-1:34:29 AM
    '
    ' Parameters :       lngItem (Long)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo RemoveCGI_Err
    '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(WinUI.Config.CGI)), 1 To 2)
104     For i = 1 To UBound(WinUI.Config.CGI)
108         strTemp1(i, 1) = WinUI.Config.CGI(i, 1)
112         strTemp1(i, 2) = WinUI.Config.CGI(i, 2)
        Next
116     ReDim WinUI.Config.CGI(1 To (IIf(UBound(WinUI.Config.CGI) = 1, 1, UBound(WinUI.Config.CGI) - 1)), 1 To 2)
120     For i = 1 To (lngItem - 1)
124         WinUI.Config.CGI(i, 1) = strTemp1(i, 1)
128         WinUI.Config.CGI(i, 2) = strTemp1(i, 2)
        Next
132     For i = (lngItem + 1) To (UBound(strTemp1))
136         WinUI.Config.CGI(i - 1, 1) = strTemp1(i, 1)
140         WinUI.Config.CGI(i - 1, 2) = strTemp1(i, 2)
        Next
    '<EhFooter>
    Exit Sub

RemoveCGI_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.RemoveCGI", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub RemovevHost(lngItem As Long)
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       RemovevHost
    ' Description:       remove a item from the WinUI.Config.vHost array
    ' Created by :       Adam
    ' Date-Time  :       9/30/2003-1:35:31 AM
    '
    ' Parameters :       lngItem (Long)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo RemovevHost_Err
    '</EhHeader>
    Dim strTemp1() As String
    Dim i As Long

100     ReDim strTemp1(1 To (UBound(WinUI.Config.vHost)), 1 To 4)
104     For i = 1 To UBound(WinUI.Config.vHost)
108         strTemp1(i, 1) = WinUI.Config.vHost(i).Name
112         strTemp1(i, 2) = WinUI.Config.vHost(i).Domain
116         strTemp1(i, 3) = WinUI.Config.vHost(i).Root
120         strTemp1(i, 4) = WinUI.Config.vHost(i).Log
        Next
124     ReDim WinUI.Config.vHost(1 To (IIf(UBound(WinUI.Config.vHost) = 1, 1, UBound(WinUI.Config.vHost) - 1)))
128     For i = 1 To (lngItem - 1)
132         WinUI.Config.vHost(i).Name = strTemp1(i, 1)
136         WinUI.Config.vHost(i).Domain = strTemp1(i, 2)
140         WinUI.Config.vHost(i).Root = strTemp1(i, 3)
144         WinUI.Config.vHost(i).Log = strTemp1(i, 4)
        Next
148     For i = lngItem + 1 To (UBound(strTemp1))
152         WinUI.Config.vHost(i - 1).Name = strTemp1(i, 1)
156         WinUI.Config.vHost(i - 1).Domain = strTemp1(i, 2)
160         WinUI.Config.vHost(i - 1).Root = strTemp1(i, 3)
164         WinUI.Config.vHost(i - 1).Log = strTemp1(i, 4)
        Next
    '<EhFooter>
    Exit Sub

RemovevHost_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.RemovevHost", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub GetUpdateStatus(strData As String)
    '<CSCM>
    '--------------------------------------------------------------------------------
    ' Project    :       SWEBS_WinUI
    ' Procedure  :       GetUpdateStatus
    ' Description:       this checks to see if there is an updata available..
    '
    '                       data is parsed here but pulled from frmMain this could
    '                       be done better via some API, a call to something in
    '                       basUtil would be much better.
    '
    '                       fills WinUI.Update with the retreived data
    '
    '                       Split()s the version strings, and uses the length of the
    '                       update stringto determine the number of times thru the
    '                       loop, thats wrong, needs to be re-written, again.
    ' Created by :       Adam
    ' Machine    :       Adams_Box
    ' Date-Time  :       9/30/2003-1:36:28 AM
    '
    ' Parameters :       strData (String)
    '--------------------------------------------------------------------------------
    '</CSCM>
    '<EhHeader>
    On Error GoTo GetUpdateStatus_Err
    '</EhHeader>
    Dim strNewVer() As String
    Dim strCurVer() As String
    Dim i As Long

100     If InStr(1, strData, "Server at swebs.sourceforge.net Port 80") = 0 And strData <> "" Then
104         WinUI.EventLog.AddEvent "basMain.GetUpdateStatus", "Update Data Found, Processing."
108         WinUI.Update.Date = GetTaggedData(strData, "Date")
112         WinUI.Update.Description = GetTaggedData(strData, "Description")
116         WinUI.Update.DownloadURL = GetTaggedData(strData, "DownloadURL")
120         WinUI.Update.InfoURL = GetTaggedData(strData, "InfoURL")
124         WinUI.Update.Version = GetTaggedData(strData, "Version")
128         WinUI.Update.UpdateLevel = GetTaggedData(strData, "UpgradeLevel")
132         WinUI.Update.FileSize = Val(GetTaggedData(strData, "FileSize"))
        
            'check to see if this is newer
136         strNewVer() = Split(WinUI.Update.Version, ".")
140         strCurVer() = Split(WinUI.Version, ".")
144         For i = 0 To UBound(strNewVer)
148             If Val(strNewVer(i)) > Val(strCurVer(i)) Then
152                 WinUI.Update.Available = True
156                 WinUI.EventLog.AddEvent "WinUI.basMain.GetUpdateStatus", "Update Available. Old Version: " & WinUI.Version & "; New Version: " & WinUI.Update.Version
                End If
            Next
160     ElseIf WinUI.Update.Available = True Then
164         WinUI.EventLog.AddEvent "WinUI.basMain.GetUpdateStatus", "Update status already true."
        Else
168         WinUI.Update.Available = False
172         WinUI.EventLog.AddEvent "WinUI.basMain.GetUpdateStatus", "No update data or update file not found."
        End If
    '<EhFooter>
    Exit Sub

GetUpdateStatus_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.GetUpdateStatus", Erl, False
    Resume Next
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
108         strLangTemp = Space$(FileLen(WinUI.Path & "lang.xml"))
112         Open WinUI.Path & "lang.xml" For Binary As 1
116             Get #1, 1, strLangTemp
120         Close 1
124         strLang = GetTaggedData(strLangTemp, "1033")
128         If strLang <> "" Then
132             WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Loaded lang: 1033"
            Else
136             WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Failed to load lang: 1033"
            End If
140         strLang = Trim$(strLang)
144         strLang = Replace(strLang, vbCrLf, "")
148         strLang = Replace(strLang, Chr$(9), "")
        Else
152         WinUI.EventLog.AddEvent "WinUI.basMain.LoadLang", "Lang.xml file is missing."
        End If
    '<EhFooter>
    Exit Sub

LoadLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.LoadLang", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub LoadDynDNSData()
    '<EhHeader>
    On Error GoTo LoadDynDNSData_Err
    '</EhHeader>
    Dim strResult As String
    
100     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSEnable")
104     If LCase(strResult) = "true" Then
108         WinUI.DynDNS.Enabled = True
        Else
112         WinUI.DynDNS.Enabled = False
        End If
116     WinUI.DynDNS.Hostname = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSHostname")
120     WinUI.DynDNS.LastIP = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastIP")
124     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastResult")
128     If strResult = "" Then
132         WinUI.DynDNS.LastResult = "(None)"
        Else
136         WinUI.DynDNS.LastResult = strResult
        End If
140     strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastUpdate")
144     If strResult = "" Then
148         WinUI.DynDNS.LastUpdate = CDate(2.00001)
        Else
152         WinUI.DynDNS.LastUpdate = CDate(strResult)
        End If
156     WinUI.DynDNS.Password = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSPassword")
160     WinUI.DynDNS.UserName = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSUsername")
    '<EhFooter>
    Exit Sub

LoadDynDNSData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basMain.LoadDynDNSData", Erl, False
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
