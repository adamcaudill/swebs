VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDebug"
Attribute VB_GlobalNameSpace = False
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

Public CallStack As cCallStack

'local variable(s) to hold property value(s)
Private mvarEnabled As Boolean 'local copy
Private mvarDebugLang As Boolean 'local copy
Private mvarDisableSplash As Boolean 'local copy
Private mvarDisableFade As Boolean 'local copy
Private mvarDisableUpdate As Boolean 'local copy
Private mvarDisableTips As Boolean 'local copy
Private mvarDebugMode As Boolean 'local copy

'CSEH: WinUI - Custom(No Stack)
Public Function ErrorReport(strErrMsg As String, strLine As String, strLocation As String) As String
    '<EhHeader>
    On Error GoTo ErrorReport_Err
    '</EhHeader>
    Dim strMessage As String
    Dim i As Long

100     strMessage = String(30, "*") & vbCrLf
104     strMessage = strMessage & "WinUI Bug Report - Version: " & WinUI.Version & vbCrLf & vbCrLf
108     strMessage = strMessage & "Time: (" & Now & ") Location: (" & strLocation & ") Line: (" & strLine & ")" & vbCrLf
112     strMessage = strMessage & "Message: (" & strErrMsg & ")" & vbCrLf
116     strMessage = strMessage & "Call Stack:" & vbCrLf
120     For i = 1 To CallStack.Count
124         strMessage = strMessage & Chr(9) & CallStack.Peek(i) & vbCrLf
        Next
128     strMessage = strMessage & vbCrLf & "Please see http://swebs.sf.net for information on this error." & vbCrLf
132     strMessage = strMessage & String(30, "*") & vbCrLf
136     Open WinUI.Path & "WinUI_ErrLog.log" For Append As 1
140         Print #1, strMessage
144     Close 1
148     ErrorReport = WinUI.Path & "WinUI_ErrLog.log"
    '<EhFooter>
    Exit Function

ErrorReport_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.ErrorReport", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Let DebugMode(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DebugMode_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DebugMode")
    '</EhHeader>
100     mvarDebugMode = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DebugMode_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugMode", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DebugMode() As Boolean
    '<EhHeader>
    On Error GoTo DebugMode_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DebugMode")
    '</EhHeader>
100     DebugMode = mvarDebugMode
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DebugMode_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugMode", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub KillUpdate()
    '<EhHeader>
    On Error GoTo KillUpdate_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.KillUpdate")
    '</EhHeader>
100     WinUI.Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "UpdateEnabled", "False"
104     mvarDisableUpdate = True
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

KillUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.KillUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Let DisableTips(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableTips_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableTips")
    '</EhHeader>
100     mvarDisableTips = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableTips_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableTips", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableTips() As Boolean
    '<EhHeader>
    On Error GoTo DisableTips_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableTips")
    '</EhHeader>
100     DisableTips = mvarDisableTips
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableTips_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableTips", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableUpdate(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableUpdate_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableUpdate")
    '</EhHeader>
100     mvarDisableUpdate = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableUpdate() As Boolean
    '<EhHeader>
    On Error GoTo DisableUpdate_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableUpdate")
    '</EhHeader>
100     DisableUpdate = mvarDisableUpdate
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableFade(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableFade_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableFade")
    '</EhHeader>
100     mvarDisableFade = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableFade_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableFade", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableFade() As Boolean
    '<EhHeader>
    On Error GoTo DisableFade_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableFade")
    '</EhHeader>
100     DisableFade = mvarDisableFade
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableFade_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableFade", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisableSplash(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DisableSplash_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableSplash")
    '</EhHeader>
100     mvarDisableSplash = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableSplash_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableSplash", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableSplash() As Boolean
    '<EhHeader>
    On Error GoTo DisableSplash_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DisableSplash")
    '</EhHeader>
100     DisableSplash = mvarDisableSplash
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DisableSplash_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DisableSplash", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let DebugLang(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo DebugLang_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DebugLang")
    '</EhHeader>
100     mvarDebugLang = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DebugLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugLang", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get DebugLang() As Boolean
    '<EhHeader>
    On Error GoTo DebugLang_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.DebugLang")
    '</EhHeader>
100     DebugLang = mvarDebugLang
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

DebugLang_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.DebugLang", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.Enabled")
    '</EhHeader>
100     mvarEnabled = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.Enabled")
    '</EhHeader>
100     Enabled = mvarEnabled
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set CallStack = New cCallStack
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set CallStack = Nothing
End Sub

Public Sub Setup()
    '<EhHeader>
    On Error GoTo Setup_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cDebug.Setup")
    '</EhHeader>
100     mvarDisableUpdate = IIf(LCase$(WinUI.Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "UpdateEnabled")) = "false", True, False)
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Setup_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cDebug.Setup", Erl, False
    Resume Next
    '</EhFooter>
End Sub
