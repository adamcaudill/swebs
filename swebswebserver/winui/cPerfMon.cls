VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPerfMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'CSEH: WinUI - Custom(No Stack)
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

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private mCol As Collection
Private curFreq As Currency
Private strPerfLog As String
Private mvarEnabled As Boolean
Private lngCallCount As Long
Private curTotalTime As Currency

Public Property Let Enabled(ByVal vData As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
100     mvarEnabled = vData
    '<EhFooter>
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
100     Enabled = mvarEnabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Enabled", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub Add(strFunction As String)
    '<EhHeader>
    On Error GoTo Add_Err
    '</EhHeader>
    Dim tData As cPerfMonData
    Dim curStartTime As Currency

100     Set tData = New cPerfMonData
104     tData.Name = strFunction
108     QueryPerformanceCounter curStartTime
112     tData.StartTime = curStartTime
116     tData.StopTime = 0
120     mCol.Add tData, strFunction
124     Set tData = Nothing
    '<EhFooter>
    Exit Sub

Add_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Add", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub Remove(strFunction As String)
    '<EhHeader>
    On Error GoTo Remove_Err
    '</EhHeader>
    Dim tData As cPerfMonData
    Dim curExecTime As Currency
    Dim curStopTime As Currency
    Dim strFixedFunction As String * 50

100     Set tData = mCol.Item(strFunction)
104     QueryPerformanceCounter curStopTime
108     tData.StopTime = curStopTime
112     curExecTime = 1000 * (tData.StopTime - tData.StartTime) / curFreq
116     curTotalTime = curTotalTime + curExecTime
120     lngCallCount = lngCallCount + 1
124     strFixedFunction = strFunction
128     strPerfLog = strPerfLog & "Function:  " & strFixedFunction & String(2, Chr(9)) & "Execution Time:  " & Space(18 - Len(Format(Str(curExecTime), "#.0000"))) & Format(Str(curExecTime), "#.0000") & vbCrLf
132     mCol.Remove strFunction
136     Set tData = Nothing
    '<EhFooter>
    Exit Sub

Remove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Remove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function Save(strFile As String) As Boolean
    '<EhHeader>
    On Error GoTo Save_Err
    '</EhHeader>
100     strPerfLog = strPerfLog & "Statistics: Call Count: " & lngCallCount & " Average Execution Time: " & (curTotalTime / lngCallCount) & vbCrLf & String(103, "*") & vbCrLf
104     Open strFile For Append As 1
108         Print #1, strPerfLog
112     Close 1
116     Save = True
    '<EhFooter>
    Exit Function

Save_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Save", Erl, False
    Resume Next
    '</EhFooter>
End Function

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set mCol = New Collection
104     QueryPerformanceFrequency curFreq
108     strPerfLog = String(103, "*") & vbCrLf & "Application Started: " & Now & vbCrLf
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMon.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set mCol = Nothing
End Sub
