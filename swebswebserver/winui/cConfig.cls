VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colvHost"
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

Public vHost As colvHost
Public CGI As colCGI
Public Index As colIndex

'local variable(s) to hold property value(s)
Private mvarServerName As String 'local copy
Private mvarPort As Integer 'local copy
Private mvarWebRoot As String 'local copy
Private mvarMaxConnections As Long 'local copy
Private mvarLogFile As String 'local copy
Private mvarAllowIndex As String 'local copy
Private mvarErrorPages As String 'local copy
Private mvarListeningAddress As String 'local copy
Private mvarErrorLog As String 'local copy

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set vHost = New colvHost
104     Set CGI = New colCGI
108     Set Index = New colIndex
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set vHost = Nothing
    Set CGI = Nothing
    Set Index = Nothing
End Sub

Public Function LoadData() As Boolean
    '<EhHeader>
    On Error GoTo LoadData_Err
    '</EhHeader>
    Dim XML As CHILKATXMLLib.XmlFactory
    Dim ConfigXML As CHILKATXMLLib.IChilkatXml
    Dim Node As CHILKATXMLLib.IChilkatXml
    Dim strTemp As String
    Dim i As Long
    
100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     ConfigXML.LoadXmlFile WinUI.ConfigFile
    
        '<ServerName>
112     Set Node = ConfigXML.SearchForTag(Nothing, "ServerName")
116     If Node Is Nothing Then
120         mvarServerName = "SWEBS Server"
        Else
124         mvarServerName = Trim$(Node.Content)
        End If
    
        '<Port>
128     Set Node = ConfigXML.SearchForTag(Nothing, "Port")
132     If Node Is Nothing Then
136         mvarPort = 80
        Else
140         mvarPort = IIf(Int(Val(Node.Content)) <= 0, 80, Int(Val(Node.Content)))
        End If
    
        '<Webroot>
144     Set Node = ConfigXML.SearchForTag(Nothing, "Webroot")
148     If Node Is Nothing Then
152         strTemp = WinUI.Path & "Webroot"
        Else
156         strTemp = Trim$(Node.Content)
        End If
160     mvarWebRoot = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<MaxConnections>
164     Set Node = ConfigXML.SearchForTag(Nothing, "MaxConnections")
168     If Node Is Nothing Then
172         mvarMaxConnections = 20
        Else
176         mvarMaxConnections = IIf(Int(Val(Node.Content)) <= 0, 20, Int(Val(Node.Content)))
        End If
    
        '<LogFile>
180     Set Node = ConfigXML.SearchForTag(Nothing, "LogFile")
184     If Node Is Nothing Then
188         mvarLogFile = WinUI.Path & "SWS.log"
        Else
192         mvarLogFile = Trim$(Node.Content)
        End If
    
        '<AllowIndex>
196     Set Node = ConfigXML.SearchForTag(Nothing, "AllowIndex")
200     If Node Is Nothing Then
204         mvarAllowIndex = "false"
        Else
208         mvarAllowIndex = IIf(LCase$(Node.Content) = "true", "true", "false")
        End If
    
        '<ErrorPages>
212     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorPages")
216     If Node Is Nothing Then
220         strTemp = WinUI.Path & "Errors"
        Else
224         strTemp = Trim$(Node.Content)
        End If
228     mvarErrorPages = IIf(Right$(strTemp, 1) = "\", Left$(strTemp, (Len(strTemp) - 1)), strTemp)
    
        '<ErrorLog>
232     Set Node = ConfigXML.SearchForTag(Nothing, "ErrorLog")
236     If Node Is Nothing Then
240         mvarErrorLog = WinUI.Path & "ErrorLog.log"
        Else
244         mvarErrorLog = Trim$(Node.Content)
        End If
    
        '<IndexFile>
248     Set Node = ConfigXML.SearchForTag(Nothing, "IndexFile")
252     If Node Is Nothing Then
256         Index.Add "index.html", "index.html"
        Else
260         Do While Not (Node Is Nothing)
264             If Trim$(Node.Content) <> "" Then
268                 Index.Add Trim$(Node.Content), Trim$(Node.Content)
                End If
272             Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
            Loop
        End If
    
        '<VirtualHost>
276     Set Node = ConfigXML.FindChild("VirtualHost")
280     If Not (Node Is Nothing) Then
284         Do While Not (Node Is Nothing)
288             If Node.GetChildContent("vhName") <> "" Then
292                 vHost.Add Trim$(Node.GetChildContent("vhName")), Trim$(Node.GetChildContent("vhHostName")), Trim$(Node.GetChildContent("vhRoot")), Trim$(Node.GetChildContent("vhLogFile")), Trim$(Node.GetChildContent("vhName"))
                End If
296             Set Node = ConfigXML.SearchForTag(Node, "VirtualHost")
            Loop
        End If

        '<CGI>
300     Set Node = ConfigXML.FindChild("CGI")
304     If Not (Node Is Nothing) Then
308         Do While Not (Node Is Nothing)
312             If Node.GetChildContent("Interpreter") <> "" Then
316                 CGI.Add Trim$(Node.GetChildContent("Interpreter")), Trim$(Node.GetChildContent("Extension")), Trim$(Node.GetChildContent("Extension"))
                End If
320             Set Node = ConfigXML.SearchForTag(Node, "CGI")
            Loop
        End If
    
        '<ListeningAddress>
324     Set Node = ConfigXML.SearchForTag(Nothing, "ListeningAddress")
328     If Node Is Nothing Then
332         mvarListeningAddress = ""
        Else
336         mvarListeningAddress = Node.Content
        End If
    
        'clean up
340     Set XML = Nothing
344     Set ConfigXML = Nothing
348     Set Node = Nothing
    
352     LoadData = True
    '<EhFooter>
    Exit Function

LoadData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.LoadData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Let ErrorLog(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ErrorLog_Err
    '</EhHeader>
100     mvarErrorLog = vData
    '<EhFooter>
    Exit Property

ErrorLog_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ErrorLog", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ErrorLog() As String
    '<EhHeader>
    On Error GoTo ErrorLog_Err
    '</EhHeader>
100     ErrorLog = mvarErrorLog
    '<EhFooter>
    Exit Property

ErrorLog_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ErrorLog", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ListeningAddress(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ListeningAddress_Err
    '</EhHeader>
100     mvarListeningAddress = vData
    '<EhFooter>
    Exit Property

ListeningAddress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ListeningAddress", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ListeningAddress() As String
    '<EhHeader>
    On Error GoTo ListeningAddress_Err
    '</EhHeader>
100     ListeningAddress = mvarListeningAddress
    '<EhFooter>
    Exit Property

ListeningAddress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ListeningAddress", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ErrorPages(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ErrorPages_Err
    '</EhHeader>
100     mvarErrorPages = vData
    '<EhFooter>
    Exit Property

ErrorPages_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ErrorPages", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ErrorPages() As String
    '<EhHeader>
    On Error GoTo ErrorPages_Err
    '</EhHeader>
100     ErrorPages = mvarErrorPages
    '<EhFooter>
    Exit Property

ErrorPages_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ErrorPages", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let AllowIndex(ByVal vData As String)
    '<EhHeader>
    On Error GoTo AllowIndex_Err
    '</EhHeader>
100     mvarAllowIndex = vData
    '<EhFooter>
    Exit Property

AllowIndex_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.AllowIndex", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get AllowIndex() As String
    '<EhHeader>
    On Error GoTo AllowIndex_Err
    '</EhHeader>
100     AllowIndex = mvarAllowIndex
    '<EhFooter>
    Exit Property

AllowIndex_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.AllowIndex", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let LogFile(ByVal vData As String)
    '<EhHeader>
    On Error GoTo LogFile_Err
    '</EhHeader>
100     mvarLogFile = vData
    '<EhFooter>
    Exit Property

LogFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.LogFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get LogFile() As String
    '<EhHeader>
    On Error GoTo LogFile_Err
    '</EhHeader>
100     LogFile = mvarLogFile
    '<EhFooter>
    Exit Property

LogFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.LogFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let MaxConnections(ByVal vData As Long)
    '<EhHeader>
    On Error GoTo MaxConnections_Err
    '</EhHeader>
100     mvarMaxConnections = vData
    '<EhFooter>
    Exit Property

MaxConnections_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.MaxConnections", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get MaxConnections() As Long
    '<EhHeader>
    On Error GoTo MaxConnections_Err
    '</EhHeader>
100     MaxConnections = mvarMaxConnections
    '<EhFooter>
    Exit Property

MaxConnections_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.MaxConnections", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let WebRoot(ByVal vData As String)
    '<EhHeader>
    On Error GoTo WebRoot_Err
    '</EhHeader>
100     mvarWebRoot = vData
    '<EhFooter>
    Exit Property

WebRoot_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.WebRoot", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get WebRoot() As String
    '<EhHeader>
    On Error GoTo WebRoot_Err
    '</EhHeader>
100     WebRoot = mvarWebRoot
    '<EhFooter>
    Exit Property

WebRoot_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.WebRoot", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Port(ByVal vData As Integer)
    '<EhHeader>
    On Error GoTo Port_Err
    '</EhHeader>
100     mvarPort = vData
    '<EhFooter>
    Exit Property

Port_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.Port", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Port() As Integer
    '<EhHeader>
    On Error GoTo Port_Err
    '</EhHeader>
100     Port = mvarPort
    '<EhFooter>
    Exit Property

Port_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.Port", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ServerName(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ServerName_Err
    '</EhHeader>
100     mvarServerName = vData
    '<EhFooter>
    Exit Property

ServerName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ServerName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ServerName() As String
    '<EhHeader>
    On Error GoTo ServerName_Err
    '</EhHeader>
100     ServerName = mvarServerName
    '<EhFooter>
    Exit Property

ServerName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cConfig.ServerName", Erl, False
    Resume Next
    '</EhFooter>
End Property
