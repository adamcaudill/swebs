VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPCGI"
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
Private mvarInterpreter As String 'local copy
Private mvarExtension As String 'local copy

Public Property Let Extension(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Extention = 5
    '<EhHeader>
    On Error GoTo Extension_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGI.Extension")
    '</EhHeader>
100     mvarExtension = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Extension_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGI.Extension", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Extension() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Extention
    '<EhHeader>
    On Error GoTo Extension_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGI.Extension")
    '</EhHeader>
100     Extension = mvarExtension
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Extension_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGI.Extension", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Let Interpreter(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Interpreter = 5
    '<EhHeader>
    On Error GoTo Interpreter_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGI.Interpreter")
    '</EhHeader>
100     mvarInterpreter = vData
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Interpreter_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGI.Interpreter", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Interpreter() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Interpreter
    '<EhHeader>
    On Error GoTo Interpreter_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGI.Interpreter")
    '</EhHeader>
100     Interpreter = mvarInterpreter
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Interpreter_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGI.Interpreter", Erl, False
    Resume Next
    '</EhFooter>
End Property



