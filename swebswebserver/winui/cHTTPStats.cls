VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPStats"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

'local variable(s) to hold property value(s)
Private mvarLastRestart As Date 'local copy
Private mvarRequestCount As Long 'local copy
Private mvarTotalBytesSent As Double 'local copy

Public Sub Reload()
Dim XML As CHILKATXMLLib.XmlFactory
Dim StatsXML As CHILKATXMLLib.IChilkatXml
Dim Node As CHILKATXMLLib.IChilkatXml
    
    Set XML = New XmlFactory
    Set StatsXML = XML.NewXml
    If Dir$(WinUI.StatsFile) <> "" Then
        WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cStats.Reload", "Loading Stats File: " & WinUI.StatsFile
        StatsXML.LoadXmlFile WinUI.StatsFile
    Else
        WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cStats.Reload", "Stats File not found."
        Exit Sub
    End If
    
    '<TotalBytesSent>
    Set Node = StatsXML.SearchForTag(Nothing, "TotalBytesSent")
    If Node Is Nothing Then
        mvarTotalBytesSent = 0
    Else
        mvarTotalBytesSent = Node.Content
    End If
    
    '<LastRestart>
    Set Node = StatsXML.SearchForTag(Nothing, "LastRestart")
    If Node Is Nothing Then
        mvarLastRestart = CDate(Now)
    Else
        mvarLastRestart = CDate(Node.Content)
    End If
    
    '<RequestCount>
    Set Node = StatsXML.SearchForTag(Nothing, "RequestCount")
    If Node Is Nothing Then
        mvarRequestCount = 0
    Else
        mvarRequestCount = Val(Node.Content)
    End If
    
    'clean up
    Set XML = Nothing
    Set StatsXML = Nothing
    Set Node = Nothing
End Sub

Public Property Get TotalBytesSent() As Double
    TotalBytesSent = mvarTotalBytesSent
End Property

Public Property Get RequestCount() As Long
    RequestCount = mvarRequestCount
End Property

Public Property Get LastRestart() As Date
    LastRestart = mvarLastRestart
End Property



