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
Public strUIPath As String
Public strInstalledVer As String
Public Config As tConfig
Public Update As tUpdate
'</GlobalVars>

'<GlobalTypes>
Public Type tConfig
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

Public Type tUpdate
    Available As Boolean
    Version As String
    Date As String
    InfoURL As String
    DownloadURL As String
    Description As String
    UpdateLevel As String
    FileSize As Long
End Type
'</GlobalTypes>

Public Sub Main()
    Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    strUIPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
    If GetSWSInstalled = False Then
        MsgBox "SWEBS Not detected. You must install SWEBS Web Server to use this application." & vbCrLf & vbCrLf & "This application will now exit.", vbCritical + vbOKOnly + vbApplicationModal
        End
    End If
    strConfigFile = GetConfigLocation
    If Dir$(strConfigFile) = "" Then
        MsgBox "Your configuration file could not be found." & vbCrLf & vbCrLf & "Please re-install the SWEBS Web Server to replace your configuration file."
        End
    End If
    Load frmMain
    DoEvents
    frmSplash.Hide
    frmMain.Show
    Unload frmSplash
End Sub

Public Function GetConfigLocation() As String
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
    GetConfigLocation = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")

End Function

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
'                    for now returns true if 'Version' is anything but null
'                    i'll finish this someday, not really a high priority.
' Created by :       Adam
' Date-Time  :       8/24/2003-2:09:24 PM
' Parameters :       none.
'--------------------------------------------------------------------------------
'</CSCM>

    strInstalledVer = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
    If strInstalledVer <> "" Then
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
        strTemp = "C:\SWS\Webroot"
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
        Config.LogFile = "C:\SWS\SWS.log"
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
        strTemp = "C:\SWS\Errors"
    Else
        strTemp = Trim$(Node.Content)
    End If
    Config.ErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
    '<IndexFile>
    ReDim Config.Index(1 To 1) As String
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    If Node Is Nothing Then
        ReDim Config.Index(1 To 1)
        Config.Index(1) = "index.htm"
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

    SaveConfigData = True
End Function

Public Function GetConfigReport() As String
Dim strReport As String
Dim strTemp As String
Dim i As Long

    strReport = "SWEBS Configuration Report"
    strReport = strReport & vbCrLf & "Date: " & Now
    strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
    strReport = strReport & "Server Name: " & Config.ServerName & vbCrLf
    strReport = strReport & "Port: " & Config.Port & vbCrLf
    strReport = strReport & "Web Root: " & Config.WebRoot & vbCrLf
    strReport = strReport & "Error Pages: " & Config.ErrorPages & vbCrLf
    strReport = strReport & "Max Connections: " & Config.MaxConnections & vbCrLf
    strReport = strReport & "Primary Log File: " & Config.LogFile & vbCrLf
    strReport = strReport & "Allow Index: " & Config.AllowIndex & vbCrLf
    For i = 1 To UBound(Config.Index)
        strTemp = strTemp & Config.Index(i) & " "
    Next
    strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    For i = 1 To UBound(Config.CGI)
        strReport = strReport & "CGI: " & "Extension: " & Config.CGI(i, 2) & " Interpreter: " & Config.CGI(i, 1) & vbCrLf
    Next
    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    For i = 1 To UBound(Config.vHost)
        strReport = strReport & "vHost: Name: " & Config.vHost(i, 1) & " Host Name: " & Config.vHost(i, 2) & " Root Directory: " & Config.vHost(i, 3) & " Log File: " & Config.vHost(i, 4) & vbCrLf
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

Public Sub GetUpdateStatus(strData As String)
    If InStr(1, strData, "404") = 0 And strData <> "" Then
        Update.Date = GetTaggedData(strData, "Date")
        Update.Description = GetTaggedData(strData, "Description")
        Update.DownloadURL = GetTaggedData(strData, "DownloadURL")
        Update.InfoURL = GetTaggedData(strData, "InfoURL")
        Update.Version = GetTaggedData(strData, "Version")
        Update.UpdateLevel = GetTaggedData(strData, "UpgradeLevel")
        Update.FileSize = Val(GetTaggedData(strData, "FileSize"))
        
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
