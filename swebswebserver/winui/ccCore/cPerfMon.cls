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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Sub Add(strFunction As String)
Dim tData As cPerfMonData
Dim curStartTime As Currency

    Set tData = New cPerfMonData
    tData.Name = strFunction
    QueryPerformanceCounter curStartTime
    tData.StartTime = curStartTime
    tData.StopTime = 0
    mCol.Add tData, strFunction
    Set tData = Nothing
End Sub

Public Sub Remove(strFunction As String)
Dim tData As cPerfMonData
Dim curExecTime As Currency
Dim curStopTime As Currency
Dim strFixedFunction As String * 50

    Set tData = mCol.Item(strFunction)
    QueryPerformanceCounter curStopTime
    tData.StopTime = curStopTime
    curExecTime = 1000 * (tData.StopTime - tData.StartTime) / curFreq
    curTotalTime = curTotalTime + curExecTime
    lngCallCount = lngCallCount + 1
    strFixedFunction = strFunction
    strPerfLog = strPerfLog & "Function:  " & strFixedFunction & String(2, Chr(9)) & "Execution Time:  " & Space(18 - Len(Format(Str(curExecTime), "#.0000"))) & Format(Str(curExecTime), "#.0000") & vbCrLf
    mCol.Remove strFunction
    Set tData = Nothing
End Sub

Public Function Save(strFile As String) As Boolean
    strPerfLog = strPerfLog & "Statistics: Call Count: " & lngCallCount & " Average Execution Time: " & (curTotalTime / lngCallCount) & vbCrLf & String(103, "*") & vbCrLf
    Open strFile For Append As 1
        Print #1, strPerfLog
    Close 1
    Save = True
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
    QueryPerformanceFrequency curFreq
    strPerfLog = String(103, "*") & vbCrLf & "Application Started: " & Now & vbCrLf
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
