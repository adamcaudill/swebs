Attribute VB_Name = "basMain"
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
Private Type tConfig
    ServerName As String
    Port As Integer
    WebRoot As String
    MaxConnections As Long
    LogFile As String
    Index() As String
    AllowIndex As String
    CGI() As String
    vHost() As String
    ErrorPages As String
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
    LoadUser32 True
    InitCommonControlsVB
    Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    strUIPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
    LoadLang
    If App.PrevInstance = True Then
        If SetFocusByCaption(GetText("SWEBS Web Server - Control Center")) = False Then
            MsgBox GetText("There is already a instance of this application running.\r\rThis application will now close."), vbOKOnly + vbInformation
        End If
        End
    End If
    App.Title = GetText("SWEBS Web Server - Control Center")
    If GetSWSInstalled = False Then
        MsgBox GetText("SWEBS Not detected. You must install SWEBS Web Server to use this application.\r\rThis application will now exit."), vbCritical + vbOKOnly + vbApplicationModal
        End
    End If
    GetConfigLocation
    If Dir$(strConfigFile) = "" Then
        MsgBox GetText("Your configuration file could not be found.\r\rPlease re-install the SWEBS Web Server to replace your configuration file."), vbCritical + vbOKOnly + vbApplicationModal
        End
    End If
    blnRegistered = GetRegistered
    LoadDynDNSData
    If GetNetStatus = True Then
        If blnRegistered = False Then
            StartRegistration
        End If
    End If
    frmSplash.Refresh
    DoEvents
    Load frmMain
    DoEvents
    frmSplash.Hide
    frmMain.Show
    Unload frmSplash
    If LCase(GetRegistryString(&H80000002, "SOFTWARE\SWS", "TODEnable")) <> "false" Then
        Load frmTip
        frmTip.Show
    End If
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
    strConfigFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
    strStatsFile = GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
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

    strInstalledVer = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
    strAppPath = GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
    strAppPath = IIf(Right$(strAppPath, 1) = "\", strAppPath, strAppPath & "\")
    If Dir$(strAppPath) <> "" Then
        GetSWSInstalled = True
    Else
        GetSWSInstalled = False
    End If

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

Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
Dim strTemp As String
Dim strTemp1() As String
Dim strTemp2() As String
Dim strTemp3() As String
Dim strTemp4() As String
Dim i As Long
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile strCurConfigFile
    
    '<ServerName>
    Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
    If Node Is Nothing Then
        Config.ServerName = "SWEBS Server"
    Else
        Config.ServerName = Trim$(Node.Content)
    End If
    
    '<Port>
    Set Node = ConfigXML.SearchForTag(Nothing, "Port")
    If Node Is Nothing Then
        Config.Port = 80
    Else
        Config.Port = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
    End If
    
    '<Webroot>
    Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
    If Node Is Nothing Then
        strTemp = strAppPath & "Webroot"
    Else
        strTemp = Trim$(Node.Content)
    End If
    Config.WebRoot = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
    '<MaxConnections>
    Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
    If Node Is Nothing Then
        Config.MaxConnections = 20
    Else
        Config.MaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
    End If
    
    '<LogFile>
    Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
    If Node Is Nothing Then
        Config.LogFile = strAppPath & "SWS.log"
    Else
        Config.LogFile = Trim$(Node.Content)
    End If
    
    '<AllowIndex>
    Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
    If Node Is Nothing Then
        Config.AllowIndex = "false"
    Else
        Config.AllowIndex = IIf(LCase$(Node.Content) = "true", "true", "false")
    End If
    
    '<ErrorPages>
    Set Node = ConfigXML.SearchForTag(Nothing, "ErrorPages")
    If Node Is Nothing Then
        strTemp = strAppPath & "Errors"
    Else
        strTemp = Trim$(Node.Content)
    End If
    Config.ErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
    '<IndexFile>
    ReDim Config.Index(1 To 1) As String
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    If Node Is Nothing Then
        ReDim Config.Index(1 To 1)
        Config.Index(1) = "index.html"
    Else
        Do While Not (Node Is Nothing)
            If Trim$(Node.Content) <> "" Then
                Config.Index(UBound(Config.Index)) = Trim$(Node.Content)
                ReDim Preserve Config.Index(1 To (UBound(Config.Index) + 1))
            End If
            Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
        Loop
        ReDim Preserve Config.Index(1 To (IIf(UBound(Config.Index) > 1, UBound(Config.Index) - 1, 1)))
    End If

    
    '<VirtualHost>
    ReDim strTemp1(1 To 1)
    ReDim strTemp2(1 To 1)
    ReDim strTemp3(1 To 1)
    ReDim strTemp4(1 To 1)
    Set Node = ConfigXML.FindChild("VirtualHost")
    If Not (Node Is Nothing) Then
        Do While Not (Node Is Nothing)
            If Node.GetChildContent("vhName") <> "" Then
                strTemp1(UBound(strTemp1)) = Trim$(Node.GetChildContent("vhName"))
                strTemp2(UBound(strTemp2)) = Trim$(Node.GetChildContent("vhHostName"))
                strTemp3(UBound(strTemp3)) = Trim$(Node.GetChildContent("vhRoot"))
                strTemp4(UBound(strTemp4)) = Trim$(Node.GetChildContent("vhLogFile"))
                ReDim Preserve strTemp1(1 To (UBound(strTemp1) + 1))
                ReDim Preserve strTemp2(1 To (UBound(strTemp2) + 1))
                ReDim Preserve strTemp3(1 To (UBound(strTemp3) + 1))
                ReDim Preserve strTemp4(1 To (UBound(strTemp4) + 1))
            End If
            Set Node = ConfigXML.SearchForTag(Node, "VirtualHost")
        Loop
        ReDim Config.vHost(1 To (IIf(UBound(strTemp1) > 1, UBound(strTemp1) - 1, 1)), 1 To 4) As String
        For i = 1 To UBound(Config.vHost)
            Config.vHost(i, 1) = strTemp1(i)
            Config.vHost(i, 2) = strTemp2(i)
            Config.vHost(i, 3) = IIf(Right$(strTemp3(i), 1) = "\", Left$(strTemp3(i), (Len(strTemp3(i)) - 1)), strTemp3(i))
            Config.vHost(i, 4) = strTemp4(i)
        Next
    Else
        ReDim Config.vHost(1 To 1, 1 To 4)
    End If

    '<CGI>
    ReDim strTemp1(1 To 1)
    ReDim strTemp2(1 To 1)
    Set Node = ConfigXML.FindChild("CGI")
    If Not (Node Is Nothing) Then
        Do While Not (Node Is Nothing)
            If Node.GetChildContent("Interpreter") <> "" Then
                strTemp1(UBound(strTemp1)) = Trim$(Node.GetChildContent("Interpreter"))
                strTemp2(UBound(strTemp2)) = Trim$(Node.GetChildContent("Extension"))
                ReDim Preserve strTemp1(1 To (UBound(strTemp1) + 1))
                ReDim Preserve strTemp2(1 To (UBound(strTemp2) + 1))
            End If
            Set Node = ConfigXML.SearchForTag(Node, "CGI")
        Loop
        ReDim Config.CGI(1 To (IIf(UBound(strTemp1) > 1, UBound(strTemp1) - 1, 1)), 2) As String
        For i = 1 To UBound(Config.CGI)
            Config.CGI(i, 1) = strTemp1(i)
            Config.CGI(i, 2) = strTemp2(i)
        Next
    Else
        ReDim Config.CGI(1 To 1, 1 To 2)
    End If
    
    'clean up
    Set XML = Nothing
    Set ConfigXML = Nothing
    Set Node = Nothing
    GetConfigData = True
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
Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim ConfigXML2 As CHILKATXMLLib.IChilkatXml
Dim i As Long

    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    Set ConfigXML2 = XML.NewXml
    
    Set ConfigXML = ConfigXML.NewChild("sws", "")
    ConfigXML.NewChild2 "ServerName", Config.ServerName
    ConfigXML.NewChild2 "Port", Config.Port
    ConfigXML.NewChild2 "Webroot", IIf(Right$(Config.WebRoot, 1) = "\", Left$(Config.WebRoot, (Len(Config.WebRoot) - 1)), Config.WebRoot)
    ConfigXML.NewChild2 "ErrorPages", IIf(Right$(Config.ErrorPages, 1) = "\", Left$(Config.ErrorPages, (Len(Config.ErrorPages) - 1)), Config.ErrorPages)
    ConfigXML.NewChild2 "MaxConnections", Config.MaxConnections
    ConfigXML.NewChild2 "LogFile", Config.LogFile
    ConfigXML.NewChild2 "AllowIndex", Config.AllowIndex
    For i = 1 To UBound(Config.Index)
        ConfigXML.NewChild2 "IndexFile", Config.Index(i)
    Next
    If Config.CGI(1, 1) <> "" Then
        For i = 1 To UBound(Config.CGI)
            Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
            ConfigXML2.NewChild2 "Interpreter", Config.CGI(i, 1)
            ConfigXML2.NewChild2 "Extension", Config.CGI(i, 2)
            ConfigXML.AddChildTree ConfigXML2
        Next
    End If
    If Config.vHost(1, 1) <> "" Then
        For i = 1 To UBound(Config.vHost)
            Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
            ConfigXML2.NewChild2 "vhName", Config.vHost(i, 1)
            ConfigXML2.NewChild2 "vhHostName", IIf(Right$(Config.vHost(i, 2), 1) = "\", Left$(Config.vHost(i, 2), (Len(Config.vHost(i, 2)) - 1)), Config.vHost(i, 2))
            ConfigXML2.NewChild2 "vhRoot", Config.vHost(i, 3)
            ConfigXML2.NewChild2 "vhLogFile", Config.vHost(i, 4)
            ConfigXML.AddChildTree ConfigXML2
        Next
    End If
    
    'ConfigXML.SaveXml strUIPath & "test.xml"
    ConfigXML.SaveXml strCurConfigFile

    'save dns config
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSHostname", DynDNS.Hostname
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastIP", DynDNS.LastIP
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastResult", DynDNS.LastResult
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSLastUpdate", DynDNS.LastUpdate
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSPassword", DynDNS.Password
    SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSUsername", DynDNS.UserName
    If DynDNS.Enabled = True Then
        SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "true"
    Else
        SaveRegistryString &H80000002, "SOFTWARE\SWS", "DNSEnable", "false"
    End If
    
    SaveConfigData = True
End Function

Public Function GetConfigReport() As String
Dim strReport As String
Dim strTemp As String
Dim i As Long

    strReport = "SWEBS Configuration Report"
    strReport = strReport & vbCrLf & GetText("Date") & ": " & Now
    strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
    strReport = strReport & GetText("Server Name") & ": " & Config.ServerName & vbCrLf
    strReport = strReport & GetText("Port") & ": & Config.Port & vbCrLf"
    strReport = strReport & GetText("Web Root") & ": " & Config.WebRoot & vbCrLf
    strReport = strReport & GetText("Error Pages") & ": " & Config.ErrorPages & vbCrLf
    strReport = strReport & GetText("Max Connections") & ": " & Config.MaxConnections & vbCrLf
    strReport = strReport & GetText("Primary Log File") & ": " & Config.LogFile & vbCrLf
    strReport = strReport & GetText("Allow Index") & ": " & Config.AllowIndex & vbCrLf
    For i = 1 To UBound(Config.Index)
        strTemp = strTemp & Config.Index(i) & " "
    Next
    strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    For i = 1 To UBound(Config.CGI)
        strReport = strReport & GetText("CGI: Extension") & ": " & Config.CGI(i, 2) & " " & GetText("Interpreter") & ": " & Config.CGI(i, 1) & vbCrLf
    Next
    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    For i = 1 To UBound(Config.vHost)
        strReport = strReport & GetText("vHost: Name") & ": " & Config.vHost(i, 1) & " " & GetText("Host Name") & ": " & Config.vHost(i, 2) & " " & GetText("Root Directory") & ": " & Config.vHost(i, 3) & " " & GetText("Log File") & ": " & Config.vHost(i, 4) & vbCrLf
    Next
    GetConfigReport = strReport
End Function

Public Sub AddNewCGI(strExt As String, strInterp As String)
Dim strTemp1() As String
Dim i As Long

    ReDim strTemp1(1 To (UBound(Config.CGI)), 1 To 2)
    For i = 1 To UBound(Config.CGI)
        strTemp1(i, 1) = Config.CGI(i, 1)
        strTemp1(i, 2) = Config.CGI(i, 2)
    Next
    ReDim Config.CGI(1 To (UBound(Config.CGI) + 1), 1 To 2)
    For i = 1 To (UBound(Config.CGI) - 1)
        Config.CGI(i, 1) = strTemp1(i, 1)
        Config.CGI(i, 2) = strTemp1(i, 2)
    Next
    Config.CGI(UBound(Config.CGI), 1) = strInterp
    Config.CGI(UBound(Config.CGI), 2) = strExt
End Sub

Public Sub AddNewvHost(strName As String, strDomain As String, strRoot As String, strLog As String)
Dim strTemp1() As String
Dim i As Long

    ReDim strTemp1(1 To (UBound(Config.vHost)), 1 To 4)
    For i = 1 To UBound(Config.vHost)
        strTemp1(i, 1) = Config.vHost(i, 1)
        strTemp1(i, 2) = Config.vHost(i, 2)
        strTemp1(i, 3) = Config.vHost(i, 3)
        strTemp1(i, 4) = Config.vHost(i, 4)
    Next
    ReDim Config.vHost(1 To IIf(Config.vHost(1, 1) = "", 1, (UBound(Config.vHost) + 1)), 1 To 4)
    For i = 1 To (UBound(Config.vHost) - 1)
        Config.vHost(i, 1) = strTemp1(i, 1)
        Config.vHost(i, 2) = strTemp1(i, 2)
        Config.vHost(i, 3) = strTemp1(i, 3)
        Config.vHost(i, 4) = strTemp1(i, 4)
    Next
    Config.vHost(UBound(Config.vHost), 1) = strName
    Config.vHost(UBound(Config.vHost), 2) = strDomain
    Config.vHost(UBound(Config.vHost), 3) = strRoot
    Config.vHost(UBound(Config.vHost), 4) = strLog
End Sub

Public Sub RemoveCGI(lngItem As Long)
Dim strTemp1() As String
Dim i As Long

    ReDim strTemp1(1 To (UBound(Config.CGI)), 1 To 2)
    For i = 1 To UBound(Config.CGI)
        strTemp1(i, 1) = Config.CGI(i, 1)
        strTemp1(i, 2) = Config.CGI(i, 2)
    Next
    ReDim Config.CGI(1 To (IIf(UBound(Config.CGI) = 1, 1, UBound(Config.CGI) - 1)), 1 To 2)
    For i = 1 To (lngItem - 1)
        Config.CGI(i, 1) = strTemp1(i, 1)
        Config.CGI(i, 2) = strTemp1(i, 2)
    Next
    For i = (lngItem + 1) To (UBound(strTemp1))
        Config.CGI(i - 1, 1) = strTemp1(i, 1)
        Config.CGI(i - 1, 2) = strTemp1(i, 2)
    Next
End Sub

Public Sub RemovevHost(lngItem As Long)
Dim strTemp1() As String
Dim i As Long

    ReDim strTemp1(1 To (UBound(Config.vHost)), 1 To 4)
    For i = 1 To UBound(Config.vHost)
        strTemp1(i, 1) = Config.vHost(i, 1)
        strTemp1(i, 2) = Config.vHost(i, 2)
        strTemp1(i, 3) = Config.vHost(i, 3)
        strTemp1(i, 4) = Config.vHost(i, 4)
    Next
    ReDim Config.vHost(1 To (IIf(UBound(Config.vHost) = 1, 1, UBound(Config.vHost) - 1)), 1 To 4)
    For i = 1 To (lngItem - 1)
        Config.vHost(i, 1) = strTemp1(i, 1)
        Config.vHost(i, 2) = strTemp1(i, 2)
        Config.vHost(i, 3) = strTemp1(i, 3)
        Config.vHost(i, 4) = strTemp1(i, 4)
    Next
    For i = lngItem + 1 To (UBound(strTemp1))
        Config.vHost(i - 1, 1) = strTemp1(i, 1)
        Config.vHost(i - 1, 2) = strTemp1(i, 2)
        Config.vHost(i - 1, 3) = strTemp1(i, 3)
        Config.vHost(i - 1, 4) = strTemp1(i, 4)
    Next
End Sub

Public Sub GetUpdateStatus(strdata As String)
    If InStr(1, strdata, "Server at swebs.sourceforge.net Port 80") = 0 And strdata <> "" Then
        Update.Date = GetTaggedData(strdata, "Date")
        Update.Description = GetTaggedData(strdata, "Description")
        Update.DownloadURL = GetTaggedData(strdata, "DownloadURL")
        Update.InfoURL = GetTaggedData(strdata, "InfoURL")
        Update.Version = GetTaggedData(strdata, "Version")
        Update.UpdateLevel = GetTaggedData(strdata, "UpgradeLevel")
        Update.FileSize = Val(GetTaggedData(strdata, "FileSize"))
        
        'check to see if this is newer
        If strInstalledVer < Update.Version Then
            Update.Available = True
        End If
    ElseIf Update.Version <> "" Then
        Update.Available = True
    Else
        Update.Available = False
    End If
End Sub

Public Sub GetStatsData()
Dim XML As CHILKATXMLLib.XmlFactory
Dim StatsXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
    
    Set XML = New XmlFactory
    Set StatsXML = XML.NewXml
    If Dir$(strStatsFile) <> "" And strStatsFile <> "" Then
        StatsXML.LoadXmlFile strStatsFile
    End If
    
    '<TotalBytesSent>
    Set Node = StatsXML.SearchForTag(Nothing, "TotalBytesSent")
    If Node Is Nothing Then
        Stats.TotalBytesSent = 0
    Else
        Stats.TotalBytesSent = Node.Content
    End If
    
    '<LastRestart>
    Set Node = StatsXML.SearchForTag(Nothing, "LastRestart")
    If Node Is Nothing Then
        Stats.LastRestart = CDate(Now)
    Else
        Stats.LastRestart = CDate(Node.Content)
    End If
    
    '<RequestCount>
    Set Node = StatsXML.SearchForTag(Nothing, "RequestCount")
    If Node Is Nothing Then
        Stats.RequestCount = 0
    Else
        Stats.RequestCount = Val(Node.Content)
    End If
    
    'clean up
    Set XML = Nothing
    Set StatsXML = Nothing
    Set Node = Nothing
End Sub

Public Function GetRegistered() As Boolean
'Dim strResult As String
'    strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "RegID")
'    If strResult <> "" Then
'        GetRegistered = True
'    Else
'        GetRegistered = False
'    End If
    
    GetRegistered = True
End Function

Public Sub StartRegistration()
Dim lngResult As Long
    lngResult = MsgBox(GetText("Would you like to register your software? It's fast and Free!\r\rProduct registration is used to provide the best possible service, products, and support for our users.\rWe will not contact you nor will we sell or give away any of your information.\r\rWould you like to register now?"), vbQuestion + vbYesNo + vbApplicationModal)
    If lngResult = vbYes Then
        Load frmRegistration
        frmRegistration.Show vbModal
    End If
End Sub

Public Function GetText(strString As String) As String
Dim strResult As String

    strResult = GetTaggedData(strLang, strString)
    strResult = CUnescape(strResult)
    If strResult <> "" Then
        GetText = strResult
    Else
        GetText = CUnescape(strString)
    End If
End Function

Private Sub LoadLang()
Dim strLangTemp As String
Dim lngLen As String

    If Dir$(strUIPath & "lang.xml") <> "" Then
        lngLen = FileLen(strUIPath & "lang.xml")
        strLangTemp = Space$(lngLen)
        Open strUIPath & "lang.xml" For Binary As 1 Len = lngLen
            Get #1, 1, strLangTemp
        Close 1
        strLang = GetTaggedData(strLangTemp, "1033")
        strLang = Trim$(strLang)
        strLang = Replace(strLang, vbCrLf, "")
        strLang = Replace(strLang, Chr$(9), "")
    End If
End Sub

Public Sub LoadDynDNSData()
Dim strResult As String
    
    strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSEnable")
    If LCase(strResult) = "true" Then
        DynDNS.Enabled = True
    Else
        DynDNS.Enabled = False
    End If
    DynDNS.Hostname = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSHostname")
    DynDNS.LastIP = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastIP")
    strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastResult")
    If strResult = "" Then
        DynDNS.LastResult = "(None)"
    Else
        DynDNS.LastResult = strResult
    End If
    strResult = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSLastUpdate")
    If strResult = "" Then
        DynDNS.LastUpdate = CDate(2.00001)
    Else
        DynDNS.LastUpdate = CDate(strResult)
    End If
    DynDNS.Password = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSPassword")
    DynDNS.UserName = GetRegistryString(&H80000002, "SOFTWARE\SWS", "DNSUsername")
End Sub
