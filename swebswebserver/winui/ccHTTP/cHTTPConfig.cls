VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colvHost"
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

Public VirtHost As cHTTPVirtHostCol
Public ISAPI As cHTTPISAPICol
Public Index As cHTTPIndexCol

'local variable(s) to hold property value(s)
Private mvarServerName As String
Private mvarPort As Integer
Private mvarWebRoot As String
Private mvarMaxConnections As Long
Private mvarLogFile As String
Private mvarAllowIndex As String
Private mvarErrorPages As String
Private mvarListeningAddress As String
Private mvarErrorLog As String
Private mvarFile As String

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    Set VirtHost = New cHTTPVirtHostCol
    Set ISAPI = New cHTTPISAPICol
    Set Index = New cHTTPIndexCol
End Sub

Private Sub Class_Terminate()
    Set VirtHost = Nothing
    Set ISAPI = Nothing
    Set Index = Nothing
End Sub

Public Sub Setup()
    mvarFile = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
End Sub

Public Function LoadData() As Boolean
Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
Dim strTemp As String
    
    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    ConfigXML.LoadXmlFile mvarFile
    
    '<ServerName>
    Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
    If Node Is Nothing Then
        mvarServerName = "SWEBS Server"
    Else
        mvarServerName = Trim$(Node.Content)
    End If
    
    '<Port>
    Set Node = ConfigXML.SearchForTag(Nothing, "Port")
    If Node Is Nothing Then
        mvarPort = 80
    Else
        mvarPort = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
    End If
    
    '<Webroot>
    Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
    If Node Is Nothing Then
        strTemp = "C:\SWS\Webroot" 'TODO: Fix Me!
    Else
        strTemp = Trim$(Node.Content)
    End If
    mvarWebRoot = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
    '<MaxConnections>
    Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
    If Node Is Nothing Then
        mvarMaxConnections = 20
    Else
        mvarMaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
    End If
    
    '<LogFile>
    Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
    If Node Is Nothing Then
        mvarLogFile = "C:\SWS\SWS.log" 'TODO: Fix Me!
    Else
        mvarLogFile = Trim$(Node.Content)
    End If
    
    '<AllowIndex>
    Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
    If Node Is Nothing Then
        mvarAllowIndex = "false"
    Else
        mvarAllowIndex = IIf(LCase$(Node.Content) = "true", "true", "false")
    End If
    
    '<ErrorPages>
    Set Node = ConfigXML.SearchForTag(Nothing, "ErrorPages")
    If Node Is Nothing Then
        strTemp = "C:\SWS\Errors" 'TODO: Fix Me!
    Else
        strTemp = Trim$(Node.Content)
    End If
    mvarErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
    '<ErrorLog>
    Set Node = ConfigXML.SearchForTag(Nothing, "ErrorLog")
    If Node Is Nothing Then
        mvarErrorLog = "C:\SWS\ErrorLog.log" 'TODO: Fix Me!
    Else
        mvarErrorLog = Trim$(Node.Content)
    End If
    
    '<IndexFile>
    Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
    If Node Is Nothing Then
        Index.Add "index.html"
    Else
        Do While Not (Node Is Nothing)
            If Trim$(Node.Content) <> "" Then
                Index.Add Trim$(Node.Content)
            End If
            Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
        Loop
    End If
    
    '<VirtualHost>
    Set Node = ConfigXML.FindChild("VirtualHost")
    If Not (Node Is Nothing) Then
        Do While Not (Node Is Nothing)
            If Node.GetChildContent("vhName") <> "" Then
                VirtHost.Add Trim$(Node.GetChildContent("vhName")), Trim$(Node.GetChildContent("vhHostName")), Trim$(Node.GetChildContent("vhRoot")), Trim$(Node.GetChildContent("vhLogFile")), Trim$(Node.GetChildContent("vhName"))
            End If
            Set Node = ConfigXML.SearchForTag(Node, "VirtualHost")
        Loop
    End If

    '<CGI>
    Set Node = ConfigXML.FindChild("ISAPI")
    If Not (Node Is Nothing) Then
        Do While Not (Node Is Nothing)
            If Node.GetChildContent("Interpreter") <> "" Then
                ISAPI.Add Trim$(Node.GetChildContent("Interpreter")), Trim$(Node.GetChildContent("Extension")), Trim$(Node.GetChildContent("Extension"))
            End If
            Set Node = ConfigXML.SearchForTag(Node, "ISAPI")
        Loop
    End If
    
    '<ListeningAddress>
    Set Node = ConfigXML.SearchForTag(Nothing, "ListeningAddress")
    If Node Is Nothing Then
        mvarListeningAddress = ""
    Else
        mvarListeningAddress = Node.Content
    End If
    
    'clean up
    Set XML = Nothing
    Set ConfigXML = Nothing
    Set Node = Nothing
    
    LoadData = True
End Function

Public Property Get File() As String
    File = mvarFile
End Property

Public Property Let ErrorLog(ByVal vData As String)
    mvarErrorLog = vData
End Property

Public Property Get ErrorLog() As String
    ErrorLog = mvarErrorLog
End Property

Public Property Let ListeningAddress(ByVal vData As String)
    mvarListeningAddress = vData
End Property

Public Property Get ListeningAddress() As String
    ListeningAddress = mvarListeningAddress
End Property

Public Property Let ErrorPages(ByVal vData As String)
    mvarErrorPages = vData
End Property

Public Property Get ErrorPages() As String
    ErrorPages = mvarErrorPages
End Property

Public Property Let AllowIndex(ByVal vData As String)
    mvarAllowIndex = vData
End Property

Public Property Get AllowIndex() As String
    AllowIndex = mvarAllowIndex
End Property

Public Property Let LogFile(ByVal vData As String)
    mvarLogFile = vData
End Property

Public Property Get LogFile() As String
    LogFile = mvarLogFile
End Property

Public Property Let MaxConnections(ByVal vData As Long)
    mvarMaxConnections = vData
End Property

Public Property Get MaxConnections() As Long
    MaxConnections = mvarMaxConnections
End Property

Public Property Let WebRoot(ByVal vData As String)
    mvarWebRoot = vData
End Property

Public Property Get WebRoot() As String
    WebRoot = mvarWebRoot
End Property

Public Property Let Port(ByVal vData As Integer)
    mvarPort = vData
End Property

Public Property Get Port() As Integer
    Port = mvarPort
End Property

Public Property Let ServerName(ByVal vData As String)
    mvarServerName = vData
End Property

Public Property Get ServerName() As String
    ServerName = mvarServerName
End Property

Public Function Report() As String
Dim strReport As String
Dim strTemp As String
Dim vItem As Variant

    strReport = "SWEBS Configuration Report"
    strReport = strReport & vbCrLf & Translator.GetText("Date") & ": " & Now
    strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
    strReport = strReport & Translator.GetText("Server Name") & ": " & mvarServerName & vbCrLf
    strReport = strReport & Translator.GetText("Port") & ": & mvarPort & vbCrLf"
    strReport = strReport & Translator.GetText("Web Root") & ": " & mvarWebRoot & vbCrLf
    strReport = strReport & Translator.GetText("Error Pages") & ": " & mvarErrorPages & vbCrLf
    strReport = strReport & Translator.GetText("Max Connections") & ": " & mvarMaxConnections & vbCrLf
    strReport = strReport & Translator.GetText("Primary Log File") & ": " & mvarLogFile & vbCrLf
    strReport = strReport & Translator.GetText("Allow Index") & ": " & mvarAllowIndex & vbCrLf
    strReport = strReport & "Listening Address" & ": " & mvarListeningAddress & vbCrLf
    strReport = strReport & "Error Log" & ": " & mvarErrorLog & vbCrLf

    For Each vItem In Index
        strTemp = strTemp & vItem.FileName & " "
    Next

    strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf

    For Each vItem In ISAPI
        strReport = strReport & Translator.GetText("ISAPI: Extension") & ": " & vItem.Extension & " " & Translator.GetText("Interpreter") & ": " & vItem.Interpreter & vbCrLf
    Next

    strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf

    For Each vItem In VirtHost
        strReport = strReport & Translator.GetText("vHost: Name") & ": " & vItem.HostName & " " & Translator.GetText("Host Name") & ": " & vItem.Domain & " " & Translator.GetText("Root Directory") & ": " & vItem.Root & " " & Translator.GetText("Log File") & ": " & vItem.Log & vbCrLf
    Next

    Report = strReport
End Function

Public Function Save(strCurConfigFile As String) As Boolean
'<CSCM>
'--------------------------------------------------------------------------------
'Project:            WinUI
'Procedure:          SaveConfigData
' Description:       this is where we save the changes to the config data.
'
'                    returns true on sucess
' Created By:        Adam
' Date-Time  :       8/25/2003-1:12:28 AM
' Parameters :       strCurConfigFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Dim XML As CHILKATXMLLib.XmlFactory
Dim ConfigXML As CHILKATXMLLib.IChilkatXml
Dim ConfigXML2 As CHILKATXMLLib.IChilkatXml
Dim vItem As Variant

    Set XML = New XmlFactory
    Set ConfigXML = XML.NewXml
    Set ConfigXML2 = XML.NewXml

    Set ConfigXML = ConfigXML.NewChild("sws", "")
    ConfigXML.NewChild2 "ServerName", mvarServerName
    ConfigXML.NewChild2 "Port", mvarPort
    ConfigXML.NewChild2 "Webroot", IIf(Right$(mvarWebRoot, 1) = "\", Left$(mvarWebRoot, (Len(mvarWebRoot) - 1)), mvarWebRoot)
    ConfigXML.NewChild2 "ErrorPages", IIf(Right$(mvarErrorPages, 1) = "\", Left$(mvarErrorPages, (Len(mvarErrorPages) - 1)), mvarErrorPages)
    ConfigXML.NewChild2 "MaxConnections", mvarMaxConnections
    ConfigXML.NewChild2 "LogFile", mvarLogFile
    ConfigXML.NewChild2 "ErrorLog", mvarErrorLog
    If mvarListeningAddress <> "" Then
        ConfigXML.NewChild2 "ListeningAddress", mvarListeningAddress
    End If
    ConfigXML.NewChild2 "AllowIndex", mvarAllowIndex

    For Each vItem In Index
        ConfigXML.NewChild2 "IndexFile", vItem.FileName
    Next

    For Each vItem In ISAPI
        Set ConfigXML2 = ConfigXML2.NewChild("ISAPI", "")
        ConfigXML2.NewChild2 "Interpreter", vItem.Interpreter
        ConfigXML2.NewChild2 "Extension", vItem.Extension
        ConfigXML.AddChildTree ConfigXML2
    Next

    For Each vItem In VirtHost
        Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
        ConfigXML2.NewChild2 "vhName", vItem.HostName
        ConfigXML2.NewChild2 "vhHostName", vItem.Domain
        ConfigXML2.NewChild2 "vhRoot", vItem.Root
        ConfigXML2.NewChild2 "vhLogFile", vItem.Log
        ConfigXML.AddChildTree ConfigXML2
    Next

    ConfigXML.SaveXml strCurConfigFile

    Save = True
End Function
