VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCGI"
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
Private mvarInterpreter As String 'local copy
Private mvarExtention As String 'local copy

Public Property Let Extention(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Extention = 5
    '<EhHeader>
    On Error GoTo Extention_Err
    '</EhHeader>
100     mvarExtention = vData
    '<EhFooter>
    Exit Property

Extention_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cCGI.Extention", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Extention() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Extention
    '<EhHeader>
    On Error GoTo Extention_Err
    '</EhHeader>
100     Extention = mvarExtention
    '<EhFooter>
    Exit Property

Extention_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cCGI.Extention", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Interpreter(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Interpreter = 5
    '<EhHeader>
    On Error GoTo Interpreter_Err
    '</EhHeader>
100     mvarInterpreter = vData
    '<EhFooter>
    Exit Property

Interpreter_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cCGI.Interpreter", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Interpreter() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Interpreter
    '<EhHeader>
    On Error GoTo Interpreter_Err
    '</EhHeader>
100     Interpreter = mvarInterpreter
    '<EhFooter>
    Exit Property

Interpreter_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cCGI.Interpreter", Erl, False
    Resume Next
    '</EhFooter>
End Property



