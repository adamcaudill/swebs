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
End Type
'</GlobalTypes>


Public Sub Main()
    strUIPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    If GetSWSInstalled = False Then
        MsgBox "SWEBS Not detected. You must install SWEBS Web Server to use this application." & vbCrLf & vbCrLf & "This application will now exit.", vbCritical + vbOKOnly + vbApplicationModal
        End
    End If
    strConfigFile = GetConfigLocation
    Load frmMain
    DoEvents
    frmMain.Show
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
Dim strResult As String
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
'                    to finish this.
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
Dim i As Long
Dim strTemp1() As String
Dim strTemp2() As String
Dim strTemp3() As String
Dim strTemp4() As String
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile strCurConfigFile
    
    '<ServerName>
    Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
    Config.ServerName = IIf(Trim(Node.Content) = "", "SWEBS Server", Trim(Node.Content))
    
    '<Port>
    Set Node = ConfigXML.SearchForTag(Nothing, "Port")
    Config.Port = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
    
    '<Webroot>
    Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
    Config.WebRoot = IIf(Right(Config.WebRoot, 1) = "\", Left(IIf(Trim(Node.Content) = "", "C:\SWS\Webroot", Trim(Node.Content)), (Len(IIf(Trim(Node.Content) = "", "C:\SWS\Webroot", Trim(Node.Content))) - 1)), Trim(IIf(Trim(Node.Content) = "", "C:\SWS\Webroot", Trim(Node.Content))))
    
    '<MaxConnections>
    Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
    Config.MaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
    
    '<LogFile>
    Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
    Config.LogFile = IIf(Trim(Node.Content) = "", "C:\SWS\SWS.log", Trim(Node.Content))
    
    '<AllowIndex>
    Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
    Config.AllowIndex = IIf(LCase(Node.Content) = "true", "true", "false")
    
    '<IndexFile>
    ReDim Config.Index(1 To 1) As String
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    Do While Not (Node Is Nothing)
        If Trim(Node.Content) <> "" Then
            Config.Index(UBound(Config.Index)) = Trim(Node.Content)
            ReDim Preserve Config.Index(1 To (UBound(Config.Index) + 1))
        End If
        Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
    Loop
    ReDim Preserve Config.Index(1 To (IIf(UBound(Config.Index) > 1, UBound(Config.Index) - 1, 1)))
    If Config.Index(1) = "" Then
        Config.Index(1) = "index.htm"
    End If
    
    '<VirtualHost>
    ReDim strTemp1(1 To 1)
    ReDim strTemp2(1 To 1)
    ReDim strTemp3(1 To 1)
    ReDim strTemp4(1 To 1)
    Set Node = ConfigXML.FindChild("VirtualHost")
    Do While Not (Node Is Nothing)
        If Node.GetChildContent("vhName") <> "" Then
            strTemp1(UBound(strTemp1)) = Trim(Node.GetChildContent("vhName"))
            strTemp2(UBound(strTemp2)) = Trim(Node.GetChildContent("vhHostName"))
            strTemp3(UBound(strTemp3)) = Trim(Node.GetChildContent("vhRoot"))
            strTemp4(UBound(strTemp4)) = Trim(Node.GetChildContent("vhLogFile"))
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
        Config.vHost(i, 2) = IIf(Right(strTemp2(i), 1) = "\", Left(strTemp2(i), (Len(strTemp2(i)) - 1)), strTemp2(i))
        Config.vHost(i, 3) = strTemp3(i)
        Config.vHost(i, 4) = strTemp4(i)
    Next i

    '<CGI>
    ReDim strTemp1(1 To 1)
    ReDim strTemp2(1 To 1)
    Set Node = ConfigXML.FindChild("CGI")
    Do While Not (Node Is Nothing)
        If Node.GetChildContent("Interpreter") <> "" Then
            strTemp1(UBound(strTemp1)) = Trim(Node.GetChildContent("Interpreter"))
            strTemp2(UBound(strTemp2)) = Trim(Node.GetChildContent("Extension"))
            ReDim Preserve strTemp1(1 To (UBound(strTemp1) + 1))
            ReDim Preserve strTemp2(1 To (UBound(strTemp2) + 1))
        End If
        Set Node = ConfigXML.SearchForTag(Node, "CGI")
    Loop
    ReDim Config.CGI(1 To (IIf(UBound(strTemp1) > 1, UBound(strTemp1) - 1, 1)), 2) As String
    For i = 1 To UBound(Config.CGI)
        Config.CGI(i, 1) = strTemp1(i)
        Config.CGI(i, 2) = strTemp2(i)
    Next i
    
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
    ConfigXML.NewChild2 "Webroot", IIf(Right(Config.WebRoot, 1) = "\", Left(Config.WebRoot, (Len(Config.WebRoot) - 1)), Config.WebRoot)
    ConfigXML.NewChild2 "MaxConnections", Config.MaxConnections
    ConfigXML.NewChild2 "LogFile", Config.LogFile
    ConfigXML.NewChild2 "AllowIndex", Config.AllowIndex
    For i = 1 To UBound(Config.Index)
        ConfigXML.NewChild2 "IndexFile", Config.Index(i)
    Next i
    For i = 1 To UBound(Config.CGI)
        Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
        ConfigXML2.NewChild2 "Interpreter", Config.CGI(i, 1)
        ConfigXML2.NewChild2 "Extension", Config.CGI(i, 2)
        ConfigXML.AddChildTree ConfigXML2
    Next i
    For i = 1 To UBound(Config.CGI)
        Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
        ConfigXML2.NewChild2 "vhName", Config.vHost(i, 1)
        ConfigXML2.NewChild2 "vhHostName", IIf(Right(Config.vHost(i, 2), 1) = "\", Left(Config.vHost(i, 2), (Len(Config.vHost(i, 2)) - 1)), Config.vHost(i, 2))
        ConfigXML2.NewChild2 "vhRoot", Config.vHost(i, 3)
        ConfigXML2.NewChild2 "vhLogFile", Config.vHost(i, 4)
        ConfigXML.AddChildTree ConfigXML2
    Next i
    
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
    strReport = strReport & vbCrLf & vbCrLf & String(30, "-") & vbCrLf & vbCrLf
    strReport = strReport & "Server Name: " & Config.ServerName & vbCrLf
    strReport = strReport & "Port: " & Config.Port & vbCrLf
    strReport = strReport & "Web Root: " & Config.WebRoot & vbCrLf
    strReport = strReport & "Max Connections: " & Config.MaxConnections & vbCrLf
    strReport = strReport & "Primary Log File: " & Config.LogFile & vbCrLf
    strReport = strReport & "Allow Indes: " & Config.AllowIndex & vbCrLf
    For i = 1 To UBound(Config.Index)
        strTemp = strTemp & Config.Index(i) & " "
    Next i
    strReport = strReport & "Index Files: " & Trim(strTemp) & vbCrLf
    strReport = strReport & vbCrLf & String(30, "-") & vbCrLf
    For i = 1 To UBound(Config.CGI)
        strReport = strReport & "CGI: " & "Extention: " & Config.CGI(i, 2) & " Interpreter: " & Config.CGI(i, 1) & vbCrLf
    Next i
    strReport = strReport & vbCrLf & String(30, "-") & vbCrLf
    For i = 1 To UBound(Config.vHost)
        strReport = strReport & "vHost: Name: " & Config.vHost(i, 1) & " Host Name: " & Config.vHost(i, 2) & " Root Directory: " & Config.vHost(i, 3) & " Log File: " & Config.vHost(i, 4) & vbCrLf
    Next i
    GetConfigReport = strReport
End Function
