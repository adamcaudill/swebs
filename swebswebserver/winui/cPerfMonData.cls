VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPerfMonData"
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

'local variable(s) to hold property value(s)
Private mvarStartTime As Currency
Private mvarStopTime As Currency
Private mvarName As String

Public Property Let Name(ByVal vData As String)
    '<EhHeader>
    On Error GoTo Name_Err
    '</EhHeader>
100     mvarName = vData
    '<EhFooter>
    Exit Property

Name_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.Name", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Name() As String
    '<EhHeader>
    On Error GoTo Name_Err
    '</EhHeader>
100     Name = mvarName
    '<EhFooter>
    Exit Property

Name_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.Name", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let StopTime(ByVal vData As Currency)
    '<EhHeader>
    On Error GoTo StopTime_Err
    '</EhHeader>
100     mvarStopTime = vData
    '<EhFooter>
    Exit Property

StopTime_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.StopTime", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get StopTime() As Currency
    '<EhHeader>
    On Error GoTo StopTime_Err
    '</EhHeader>
100     StopTime = mvarStopTime
    '<EhFooter>
    Exit Property

StopTime_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.StopTime", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let StartTime(ByVal vData As Currency)
    '<EhHeader>
    On Error GoTo StartTime_Err
    '</EhHeader>
100     mvarStartTime = vData
    '<EhFooter>
    Exit Property

StartTime_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.StartTime", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get StartTime() As Currency
    '<EhHeader>
    On Error GoTo StartTime_Err
    '</EhHeader>
100     StartTime = mvarStartTime
    '<EhFooter>
    Exit Property

StartTime_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cPerfMonData.StartTime", Erl, False
    Resume Next
    '</EhFooter>
End Property



