VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIndex"
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
Private mvarFileName As String 'local copy

Public Property Let FileName(ByVal vData As String)
    '<EhHeader>
    On Error GoTo FileName_Err
    '</EhHeader>
100     mvarFileName = vData
    '<EhFooter>
    Exit Property

FileName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cIndex.FileName", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get FileName() As String
    '<EhHeader>
    On Error GoTo FileName_Err
    '</EhHeader>
100     FileName = mvarFileName
    '<EhFooter>
    Exit Property

FileName_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cIndex.FileName", Erl, False
    Resume Next
    '</EhFooter>
End Property



