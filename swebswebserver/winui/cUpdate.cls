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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.Check")
    '</EhHeader>
    Dim strData As String
    Dim strNewVer() As String
    Dim strCurVer() As String
    Dim i As Long

100     If WinUI.Debuger.DisableUpdate <> True Then
104         If WinUI.Net.IsOnline = True Then
108             strData = Replace(WinUI.Net.PageSource("http://swebs.sf.net/upgrade.xml"), vbLf, vbCrLf)
            End If
112         If InStr(1, strData, "Server at swebs.sourceforge.net Port 80") = 0 And strData <> "" Then
116             WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update Data Found, Processing."
120             mvarReleaseDate = WinUI.Util.GetTaggedData(strData, "Date")
124             mvarDescription = WinUI.Util.GetTaggedData(strData, "Description")
128             mvarDownloadURL = WinUI.Util.GetTaggedData(strData, "DownloadURL")
132             mvarInfoURL = WinUI.Util.GetTaggedData(strData, "InfoURL")
136             mvarVersion = WinUI.Util.GetTaggedData(strData, "Version")
140             mvarUpdateLevel = WinUI.Util.GetTaggedData(strData, "UpgradeLevel")
144             mvarFileSize = Val(WinUI.Util.GetTaggedData(strData, "FileSize"))
            
                'check to see if this is newer
                'this is a bad way to do things, needs to be fixed, again
148             strNewVer() = Split(mvarVersion, ".")
152             strCurVer() = Split(WinUI.Version, ".")
156             For i = 0 To UBound(strNewVer)
160                 If Val(strNewVer(i)) > Val(strCurVer(i)) Then
164                     mvarIsAvailable = True
168                     WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update Available. Old Version: " & WinUI.Version & "; New Version: " & WinUI.Update.Version
                    End If
                Next
172         ElseIf mvarIsAvailable = True Then
176             WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "Update status already true."
            Else
180             mvarIsAvailable = False
184             WinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.cUpdate.Check", "No update data or update file not found."
            End If
        End If
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Check_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.Check", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Get FileSize() As Long
    '<EhHeader>
    On Error GoTo FileSize_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.FileSize")
    '</EhHeader>
100     FileSize = mvarFileSize
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

FileSize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.FileSize", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get UpdateLevel() As String
    '<EhHeader>
    On Error GoTo UpdateLevel_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.UpdateLevel")
    '</EhHeader>
100     UpdateLevel = mvarUpdateLevel
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

UpdateLevel_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.UpdateLevel", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Description() As String
    '<EhHeader>
    On Error GoTo Description_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.Description")
    '</EhHeader>
100     Description = mvarDescription
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Description_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.Description", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DownloadURL() As String
    '<EhHeader>
    On Error GoTo DownloadURL_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.DownloadURL")
    '</EhHeader>
100     DownloadURL = mvarDownloadURL
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DownloadURL_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.DownloadURL", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get InfoURL() As String
    '<EhHeader>
    On Error GoTo InfoURL_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.InfoURL")
    '</EhHeader>
100     InfoURL = mvarInfoURL
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

InfoURL_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.InfoURL", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get ReleaseDate() As String
    '<EhHeader>
    On Error GoTo ReleaseDate_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.ReleaseDate")
    '</EhHeader>
100     ReleaseDate = mvarReleaseDate
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

ReleaseDate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.ReleaseDate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Version() As String
    '<EhHeader>
    On Error GoTo Version_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.Version")
    '</EhHeader>
100     Version = mvarVersion
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Version_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.Version", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get IsAvailable() As Boolean
    '<EhHeader>
    On Error GoTo IsAvailable_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cUpdate.IsAvailable")
    '</EhHeader>
100     IsAvailable = mvarIsAvailable
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

IsAvailable_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cUpdate.IsAvailable", Erl, False
    Resume Next
    '</EhFooter>
End Property



