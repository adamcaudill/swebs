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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Public PerfMon As cPerfMon

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
Dim strMessage As String
Dim i As Long

    strMessage = String(30, "*") & vbCrLf
    strMessage = strMessage & Translator.GetText("WinUI Bug Report - Version:") & " " & gCore.Version & vbCrLf & vbCrLf
    strMessage = strMessage & Translator.GetText("Time:") & " (" & Now & ") " & Translator.GetText("Location:") & " (" & strLocation & ") " & Translator.GetText("Line:") & " (" & strLine & ")" & vbCrLf
    strMessage = strMessage & Translator.GetText("Message:") & " (" & strErrMsg & ")" & vbCrLf
    strMessage = strMessage & Translator.GetText("Call Stack:") & vbCrLf
    For i = 1 To CallStack.Count
        strMessage = strMessage & Chr(9) & CallStack.Peek(i) & vbCrLf
    Next
    strMessage = strMessage & vbCrLf & Translator.GetText("Please see http://swebs.sf.net for information on this error.") & vbCrLf
    strMessage = strMessage & String(30, "*") & vbCrLf
    Open gCore.Path & "WinUI_ErrLog.log" For Append As 1
        Print #1, strMessage
    Close 1
    ErrorReport = gCore.Path & "WinUI_ErrLog.log"
End Function

Public Property Let DebugMode(ByVal vData As Boolean)
    mvarDebugMode = vData
End Property

Public Property Get DebugMode() As Boolean
    DebugMode = mvarDebugMode
End Property

Public Sub KillUpdate()
    Util.SaveRegistryString &H80000002, "SOFTWARE\SWS", "UpdateEnabled", "False"
    mvarDisableUpdate = True
End Sub

Public Property Let DisableUpdate(ByVal vData As Boolean)
    mvarDisableUpdate = vData
End Property

Public Property Get DisableUpdate() As Boolean
    DisableUpdate = mvarDisableUpdate
End Property

Public Property Let DebugLang(ByVal vData As Boolean)
    mvarDebugLang = vData
End Property

Public Property Get DebugLang() As Boolean
    DebugLang = mvarDebugLang
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
    Set CallStack = New cCallStack
    Set PerfMon = New cPerfMon
End Sub

Private Sub Class_Terminate()
    Set CallStack = Nothing
    Set PerfMon = Nothing
End Sub

Public Sub Setup()
    mvarDisableUpdate = IIf(LCase$(Util.GetRegistryString(&H80000002, "SOFTWARE\SWS", "UpdateEnabled")) = "false", True, False)
End Sub
