VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUpdate"
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
Private mvarIsAvailable As Boolean 'local copy
Private mvarVersion As String 'local copy
Private mvarReleaseDate As String 'local copy
Private mvarInfoURL As String 'local copy
Private mvarDownloadURL As String 'local copy
Private mvarDescription As String 'local copy
Private mvarUpdateLevel As String 'local copy
Private mvarFileSize As Long 'local copy

Public Sub Check()
    '<EhHeader>
    On Error GoTo Check_Err
    '</EhHeader>
    Dim strData As String
    Dim strNewVer() As String
    Dim strCurVer() As String
    Dim i As Long

100     If mWinUI.Network.IsOnline = True Then
104         strData = Replace(mWinUI.Network.PageSource("http://swebs.sf.net/upgrade.xml"), vbLf, vbCrLf)
        End If
108     If InStr(1, strData, "Server at swebs.sourceforge.net Port 80") = 0 And strData <> "" Then
112         mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update Data Found, Processing."
116         mvarReleaseDate = mWinUI.Util.GetTaggedData(strData, "Date")
120         mvarDescription = mWinUI.Util.GetTaggedData(strData, "Description")
124         mvarDownloadURL = mWinUI.Util.GetTaggedData(strData, "DownloadURL")
128         mvarInfoURL = mWinUI.Util.GetTaggedData(strData, "InfoURL")
132         mvarVersion = mWinUI.Util.GetTaggedData(strData, "Version")
136         mvarUpdateLevel = mWinUI.Util.GetTaggedData(strData, "UpgradeLevel")
140         mvarFileSize = Val(mWinUI.Util.GetTaggedData(strData, "FileSize"))
        
            'check to see if this is newer
144         strNewVer() = Split(mvarVersion, ".")
148         strCurVer() = Split(mWinUI.Version, ".")
152         For i = 0 To UBound(strNewVer)
156             If Val(strNewVer(i)) > Val(strCurVer(i)) Then
160                 mvarIsAvailable = True
164                 mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update Available. Old Version: " & mWinUI.Version & "; New Version: " & mWinUI.Update.Version
                End If
            Next
168     ElseIf mvarIsAvailable = True Then
172         mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update status already true."
        Else
176         mvarIsAvailable = False
180         mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "No update data or update file not found."
        End If
    '<EhFooter>
    Exit Sub

Check_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.Check", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Get FileSize() As Long
    '<EhHeader>
    On Error GoTo FileSize_Err
    '</EhHeader>
100     FileSize = mvarFileSize
    '<EhFooter>
    Exit Property

FileSize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.FileSize", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get UpdateLevel() As String
    '<EhHeader>
    On Error GoTo UpdateLevel_Err
    '</EhHeader>
100     UpdateLevel = mvarUpdateLevel
    '<EhFooter>
    Exit Property

UpdateLevel_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.UpdateLevel", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Description() As String
    '<EhHeader>
    On Error GoTo Description_Err
    '</EhHeader>
100     Description = mvarDescription
    '<EhFooter>
    Exit Property

Description_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.Description", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DownloadURL() As String
    '<EhHeader>
    On Error GoTo DownloadURL_Err
    '</EhHeader>
100     DownloadURL = mvarDownloadURL
    '<EhFooter>
    Exit Property

DownloadURL_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.DownloadURL", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get InfoURL() As String
    '<EhHeader>
    On Error GoTo InfoURL_Err
    '</EhHeader>
100     InfoURL = mvarInfoURL
    '<EhFooter>
    Exit Property

InfoURL_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.InfoURL", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ReleaseDate() As String
    '<EhHeader>
    On Error GoTo ReleaseDate_Err
    '</EhHeader>
100     ReleaseDate = mvarReleaseDate
    '<EhFooter>
    Exit Property

ReleaseDate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.ReleaseDate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Version() As String
    '<EhHeader>
    On Error GoTo Version_Err
    '</EhHeader>
100     Version = mvarVersion
    '<EhFooter>
    Exit Property

Version_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.Version", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get IsAvailable() As Boolean
    '<EhHeader>
    On Error GoTo IsAvailable_Err
    '</EhHeader>
100     IsAvailable = mvarIsAvailable
    '<EhFooter>
    Exit Property

IsAvailable_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUpdate.IsAvailable", Erl, False
    Resume Next
    '</EhFooter>
End Property



