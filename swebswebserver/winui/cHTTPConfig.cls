VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPConfig"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public CGI As cHTTPCGICol
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
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set VirtHost = New cHTTPVirtHostCol
104     Set CGI = New cHTTPCGICol
108     Set Index = New cHTTPIndexCol
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set VirtHost = Nothing
    Set CGI = Nothing
    Set Index = Nothing
End Sub

Public Sub Setup()
    '<EhHeader>
    On Error GoTo Setup_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.Setup")
    '</EhHeader>
100     mvarFile = WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Setup_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Setup", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function LoadData() As Boolean
    '<EhHeader>
    On Error GoTo LoadData_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.LoadData")
    '</EhHeader>
    Dim XML As CHILKATXMLLib.XmlFactory
    Dim ConfigXML As CHILKATXMLLib.IChilkatXml
    Dim Node As CHILKATXMLLib.IChilkatXml
    Dim strTemp As String
    
100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     ConfigXML.LoadXmlFile mvarFile
    
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
256         Index.Add "index.html"
        Else
260         Do While Not (Node Is Nothing)
264             If Trim$(Node.Content) <> "" Then
268                 Index.Add Trim$(Node.Content)
                End If
272             Set Node = ConfigXML.SearchForTag(Node, "IndexFile")
            Loop
        End If
    
        '<VirtualHost>
276     Set Node = ConfigXML.FindChild("VirtualHost")
280     If Not (Node Is Nothing) Then
284         Do While Not (Node Is Nothing)
288             If Node.GetChildContent("vhName") <> "" Then
292                 VirtHost.Add Trim$(Node.GetChildContent("vhName")), Trim$(Node.GetChildContent("vhHostName")), Trim$(Node.GetChildContent("vhRoot")), Trim$(Node.GetChildContent("vhLogFile")), Trim$(Node.GetChildContent("vhName"))
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
    WinUI.Debuger.CallStack.Pop
    Exit Function

LoadData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.LoadData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Get File() As String
    '<EhHeader>
    On Error GoTo File_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.File")
    '</EhHeader>
100     File = mvarFile
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

File_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.File", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ErrorLog(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ErrorLog_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ErrorLog")
    '</EhHeader>
100     mvarErrorLog = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ErrorLog_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ErrorLog", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ErrorLog() As String
    '<EhHeader>
    On Error GoTo ErrorLog_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ErrorLog")
    '</EhHeader>
100     ErrorLog = mvarErrorLog
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ErrorLog_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ErrorLog", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ListeningAddress(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ListeningAddress_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ListeningAddress")
    '</EhHeader>
100     mvarListeningAddress = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ListeningAddress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ListeningAddress", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ListeningAddress() As String
    '<EhHeader>
    On Error GoTo ListeningAddress_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ListeningAddress")
    '</EhHeader>
100     ListeningAddress = mvarListeningAddress
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ListeningAddress_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ListeningAddress", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ErrorPages(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ErrorPages_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ErrorPages")
    '</EhHeader>
100     mvarErrorPages = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ErrorPages_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ErrorPages", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ErrorPages() As String
    '<EhHeader>
    On Error GoTo ErrorPages_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ErrorPages")
    '</EhHeader>
100     ErrorPages = mvarErrorPages
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ErrorPages_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ErrorPages", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let AllowIndex(ByVal vData As String)
    '<EhHeader>
    On Error GoTo AllowIndex_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.AllowIndex")
    '</EhHeader>
100     mvarAllowIndex = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

AllowIndex_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.AllowIndex", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get AllowIndex() As String
    '<EhHeader>
    On Error GoTo AllowIndex_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.AllowIndex")
    '</EhHeader>
100     AllowIndex = mvarAllowIndex
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

AllowIndex_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.AllowIndex", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let LogFile(ByVal vData As String)
    '<EhHeader>
    On Error GoTo LogFile_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.LogFile")
    '</EhHeader>
100     mvarLogFile = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

LogFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.LogFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get LogFile() As String
    '<EhHeader>
    On Error GoTo LogFile_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.LogFile")
    '</EhHeader>
100     LogFile = mvarLogFile
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

LogFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.LogFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let MaxConnections(ByVal vData As Long)
    '<EhHeader>
    On Error GoTo MaxConnections_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.MaxConnections")
    '</EhHeader>
100     mvarMaxConnections = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

MaxConnections_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.MaxConnections", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get MaxConnections() As Long
    '<EhHeader>
    On Error GoTo MaxConnections_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.MaxConnections")
    '</EhHeader>
100     MaxConnections = mvarMaxConnections
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

MaxConnections_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.MaxConnections", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let WebRoot(ByVal vData As String)
    '<EhHeader>
    On Error GoTo WebRoot_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.WebRoot")
    '</EhHeader>
100     mvarWebRoot = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

WebRoot_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.WebRoot", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get WebRoot() As String
    '<EhHeader>
    On Error GoTo WebRoot_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.WebRoot")
    '</EhHeader>
100     WebRoot = mvarWebRoot
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

WebRoot_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.WebRoot", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Port(ByVal vData As Integer)
    '<EhHeader>
    On Error GoTo Port_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.Port")
    '</EhHeader>
100     mvarPort = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Port_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Port", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Port() As Integer
    '<EhHeader>
    On Error GoTo Port_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.Port")
    '</EhHeader>
100     Port = mvarPort
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Port_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Port", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let ServerName(ByVal vData As String)
    '<EhHeader>
    On Error GoTo ServerName_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ServerName")
    '</EhHeader>
100     mvarServerName = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ServerName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ServerName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ServerName() As String
    '<EhHeader>
    On Error GoTo ServerName_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.ServerName")
    '</EhHeader>
100     ServerName = mvarServerName
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ServerName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.ServerName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Function Report() As String
    '<EhHeader>
    On Error GoTo Report_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.Report")
    '</EhHeader>
    Dim strReport As String
    Dim strTemp As String
    Dim vItem As Variant

100     strReport = "SWEBS Configuration Report"
104     strReport = strReport & vbCrLf & WinUI.GetTranslatedText("Date") & ": " & Now
108     strReport = strReport & vbCrLf & vbCrLf & String$(30, "-") & vbCrLf & vbCrLf
112     strReport = strReport & WinUI.GetTranslatedText("Server Name") & ": " & mvarServerName & vbCrLf
116     strReport = strReport & WinUI.GetTranslatedText("Port") & ": & mvarPort & vbCrLf"
120     strReport = strReport & WinUI.GetTranslatedText("Web Root") & ": " & mvarWebRoot & vbCrLf
124     strReport = strReport & WinUI.GetTranslatedText("Error Pages") & ": " & mvarErrorPages & vbCrLf
128     strReport = strReport & WinUI.GetTranslatedText("Max Connections") & ": " & mvarMaxConnections & vbCrLf
132     strReport = strReport & WinUI.GetTranslatedText("Primary Log File") & ": " & mvarLogFile & vbCrLf
136     strReport = strReport & WinUI.GetTranslatedText("Allow Index") & ": " & mvarAllowIndex & vbCrLf
140     strReport = strReport & "Listening Address" & ": " & mvarListeningAddress & vbCrLf
144     strReport = strReport & "Error Log" & ": " & mvarErrorLog & vbCrLf
    
148     For Each vItem In Index
152         strTemp = strTemp & vItem.FileName & " "
        Next
    
156     strReport = strReport & "Index Files: " & Trim$(strTemp) & vbCrLf
160     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    
164     For Each vItem In CGI
168         strReport = strReport & WinUI.GetTranslatedText("CGI: Extension") & ": " & vItem.Extension & " " & WinUI.GetTranslatedText("Interpreter") & ": " & vItem.Interpreter & vbCrLf
        Next
    
172     strReport = strReport & vbCrLf & String$(30, "-") & vbCrLf
    
176     For Each vItem In VirtHost
180         strReport = strReport & WinUI.GetTranslatedText("vHost: Name") & ": " & vItem.HostName & " " & WinUI.GetTranslatedText("Host Name") & ": " & vItem.Domain & " " & WinUI.GetTranslatedText("Root Directory") & ": " & vItem.Root & " " & WinUI.GetTranslatedText("Log File") & ": " & vItem.Log & vbCrLf
        Next
    
184     Report = strReport
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

Report_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Report", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function Save(strCurConfigFile As String) As Boolean
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
    On Error GoTo Save_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPConfig.Save")
    '</EhHeader>
    Dim XML As CHILKATXMLLib.XmlFactory
    Dim ConfigXML As CHILKATXMLLib.IChilkatXml
    Dim ConfigXML2 As CHILKATXMLLib.IChilkatXml
    Dim vItem As Variant

100     Set XML = New XmlFactory
104     Set ConfigXML = XML.NewXml
108     Set ConfigXML2 = XML.NewXml
    
112     Set ConfigXML = ConfigXML.NewChild("sws", "")
116     ConfigXML.NewChild2 "ServerName", mvarServerName
120     ConfigXML.NewChild2 "Port", mvarPort
124     ConfigXML.NewChild2 "Webroot", IIf(Right$(mvarWebRoot, 1) = "\", Left$(mvarWebRoot, (Len(mvarWebRoot) - 1)), mvarWebRoot)
128     ConfigXML.NewChild2 "ErrorPages", IIf(Right$(mvarErrorPages, 1) = "\", Left$(mvarErrorPages, (Len(mvarErrorPages) - 1)), mvarErrorPages)
132     ConfigXML.NewChild2 "MaxConnections", mvarMaxConnections
136     ConfigXML.NewChild2 "LogFile", mvarLogFile
140     ConfigXML.NewChild2 "ErrorLog", mvarErrorLog
144     If WinUI.Server.HTTP.Config.ListeningAddress <> "" Then
148         ConfigXML.NewChild2 "ListeningAddress", mvarListeningAddress
        End If
152     ConfigXML.NewChild2 "AllowIndex", mvarAllowIndex
    
156     For Each vItem In Index
160         ConfigXML.NewChild2 "IndexFile", vItem.FileName
        Next
    
164     For Each vItem In CGI
168         Set ConfigXML2 = ConfigXML2.NewChild("CGI", "")
172         ConfigXML2.NewChild2 "Interpreter", vItem.Interpreter
176         ConfigXML2.NewChild2 "Extension", vItem.Extension
180         ConfigXML.AddChildTree ConfigXML2
        Next
    
184     For Each vItem In VirtHost
188         Set ConfigXML2 = ConfigXML2.NewChild("VirtualHost", "")
192         ConfigXML2.NewChild2 "vhName", vItem.HostName
196         ConfigXML2.NewChild2 "vhHostName", vItem.Domain
200         ConfigXML2.NewChild2 "vhRoot", vItem.Root
204         ConfigXML2.NewChild2 "vhLogFile", vItem.Log
208         ConfigXML.AddChildTree ConfigXML2
        Next
    
212     ConfigXML.SaveXml strCurConfigFile

216     Save = True
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

Save_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPConfig.Save", Erl, False
    Resume Next
    '</EhFooter>
End Function
