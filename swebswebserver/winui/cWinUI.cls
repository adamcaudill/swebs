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
Private mvarStatsFile As String
Private mvarPath As String
Private mvarVersion As String
Private mvarConfigFile As String

Private strLang As String

Public EventLog As cEventLog
Public Registration As cRegistration
Public Update As cUpdate
Public Net As cNet
Public Util As cUtil
Public Debuger As cDebug
Public Server As cServer

Public Property Get Version() As String
    Version = mvarVersion
End Property

Public Property Get Path() As String
    Path = mvarPath
End Property

Public Property Get ConfigFile() As String
    ConfigFile = mvarConfigFile
End Property

Public Property Get StatsFile() As String
    StatsFile = mvarStatsFile
End Property

Private Sub Class_Initialize()
    Set Util = New cUtil
    Set Debuger = New cDebug
    Set EventLog = New cEventLog
    Set Registration = New cRegistration
    Set Update = New cUpdate
    Set Net = New cNet
    Set Server = New cServer
    
    mvarStatsFile = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "StatsFile")
    mvarConfigFile = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")
    mvarVersion = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
    mvarPath = Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "AppPath")
    mvarPath = IIf(Right(mvarPath, 1) <> "\", mvarPath & "\", mvarPath)
    Debuger.DisableUpdate = IIf(LCase(Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "UpdateEnabled")) = "false", True, False)
    
    LoadLang "1033"
    Util.LoadUser32 True
    Util.InitCommonControlsVB
End Sub

Private Sub Class_Terminate()
    Set EventLog = Nothing
    Set Registration = Nothing
    Set Update = Nothing
    Set Net = Nothing
    Set Util = Nothing
    Set Debuger = Nothing
    Set Server = Nothing
End Sub

Private Sub LoadLang(strLangID As String)
Dim strLangTemp As String

    If Dir$(mvarPath & "lang.xml") <> "" Then
        strLangTemp = Space$(FileLen(mvarPath & "lang.xml"))
        Open mvarPath & "lang.xml" For Binary As 1
            Get #1, 1, strLangTemp
        Close 1
        strLang = Util.GetTaggedData(strLangTemp, "lang-code id=" & Chr(34) & strLangID & Chr(34), "lang-code")
        If strLang <> "" Then
            EventLog.AddEvent "SWEBS_WinUI_DLL.cWinUI.LoadLang", "Loaded lang: " & strLangID
        Else
            EventLog.AddEvent "SWEBS_WinUI_DLL.cWinUI.LoadLang", "Failed to load lang: " & strLangID
        End If
        strLang = Trim$(strLang)
        strLang = Replace(strLang, vbCrLf, "")
        strLang = Replace(strLang, Chr$(9), "")
    Else
        EventLog.AddEvent "SWEBS_WinUI_DLL.cWinUI.LoadLang", "Lang.xml file is missing."
        If Debuger.DebugLang = True Then
            MsgBox "Failed to load lang.xml", vbInformation
        End If
    End If
End Sub

Public Function GetTranslatedText(strString As String) As String
Dim strResult As String

    If strString <> "Lang-Maintainer" Then
        strResult = WinUI.Util.GetTaggedData(strLang, "lang base=" & Chr(34) & strString & Chr(34), "lang")
    Else
        strResult = WinUI.Util.GetTaggedData(strLang, strString)
    End If
    strResult = WinUI.Util.CUnescape(strResult)
    If strResult <> "" Then
        GetTranslatedText = strResult
    Else
        GetTranslatedText = WinUI.Util.CUnescape(strString)
        If Debuger.DebugLang = True Then
            MsgBox "Failed to load translated string:" & vbCrLf & vbCrLf & strString, vbInformation
        End If
    End If
End Function
