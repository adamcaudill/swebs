VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinUI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private mvarStatsFile As String 'local copy
Private mvarPath As String 'local copy
Private mvarVersion As String 'local copy

Private strLang As String

Public Config As cConfig
Public Stats As cStats
Public EventLog As cEventLog
Public Registration As cRegistration
Public Update As cUpdate
Public Network As cNetwork
Public Dialog As cDialog
Public Util As cUtil

Public Property Get Version() As String
    '<EhHeader>
    On Error GoTo Version_Err
    '</EhHeader>
100     Version = mvarVersion
    '<EhFooter>
    Exit Property

Version_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.Version", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Path() As String
    '<EhHeader>
    On Error GoTo Path_Err
    '</EhHeader>
100     Path = mvarPath
    '<EhFooter>
    Exit Property

Path_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.Path", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get StatsFile() As String
    '<EhHeader>
    On Error GoTo StatsFile_Err
    '</EhHeader>
100     StatsFile = mvarStatsFile
    '<EhFooter>
    Exit Property

StatsFile_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.StatsFile", Erl, False
    Resume Next
    '</EhFooter>
End Property

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set Util = New cUtil
104     Set Dialog = New cDialog
108     Set EventLog = New cEventLog
112     Set Stats = New cStats
116     Set Registration = New cRegistration
120     Set Update = New cUpdate
124     Set Network = New cNetwork
128     Set Config = New cConfig
    
132     mvarStatsFile = mWinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
136     mvarVersion = mWinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
140     mvarPath = mWinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
144     mvarPath = IIf(Right(mvarPath, 1) <> "\", mvarPath & "\", mvarPath)
    
148     LoadLang "1033"
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set EventLog = Nothing
    Set Stats = Nothing
    Set Registration = Nothing
    Set Update = Nothing
    Set Network = Nothing
    Set Config = Nothing
    Set Dialog = Nothing
    Set Util = Nothing
    DestroyUI
End Sub

Private Sub LoadLang(strLangID As String)
    '<EhHeader>
    On Error GoTo LoadLang_Err
    '</EhHeader>
    Dim strLangTemp As String

100     If Dir$(mvarPath & "lang.xml") <> "" Then
104         strLangTemp = Space$(FileLen(mvarPath & "lang.xml"))
108         Open mvarPath & "lang.xml" For Binary As 1
112             Get #1, 1, strLangTemp
116         Close 1
120         strLang = mWinUI.Util.GetTaggedData(strLangTemp, strLangID)
124         If strLang <> "" Then
128             EventLog.AddEvent "WinUI.basMain.LoadLang", "Loaded lang: " & strLangID
            Else
132             EventLog.AddEvent "WinUI.basMain.LoadLang", "Failed to load lang: " & strLangID
            End If
136         strLang = Trim$(strLang)
140         strLang = Replace(strLang, vbCrLf, "")
144         strLang = Replace(strLang, Chr$(9), "")
        Else
148         EventLog.AddEvent "WinUI.basMain.LoadLang", "Lang.xml file is missing."
        End If
    '<EhFooter>
    Exit Sub

LoadLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.LoadLang", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function GetTranslatedText(strString As String) As String
    '<EhHeader>
    On Error GoTo GetTranslatedText_Err
    '</EhHeader>
    Dim strResult As String

100     strResult = mWinUI.Util.GetTaggedData(strLang, strString)
104     strResult = mWinUI.Util.CUnescape(strResult)
108     If strResult <> "" Then
112         GetTranslatedText = strResult
        Else
116         GetTranslatedText = mWinUI.Util.CUnescape(strString)
        End If
    '<EhFooter>
    Exit Function

GetTranslatedText_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cWinUI.GetTranslatedText", Erl, False
    Resume Next
    '</EhFooter>
End Function
