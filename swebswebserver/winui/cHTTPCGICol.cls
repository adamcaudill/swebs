VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPCGICol"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cCGI"
Attribute VB_Ext_KEY = "Member0" ,"cCGI"
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

'local variable to hold collection
Private mCol As Collection

Public Function Add(Interpreter As String, Extension As String, Optional sKey As String) As cHTTPCGI
        'create a new object
    '<EhHeader>
    On Error GoTo Add_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.Add")
    '</EhHeader>
        Dim objNewMember As cHTTPCGI
100     Set objNewMember = New cHTTPCGI
        'set the properties passed into the method
104     objNewMember.Interpreter = Interpreter
108     objNewMember.Extension = Extension
112     If Len(sKey) = 0 Then
116         mCol.Add objNewMember
        Else
120         mCol.Add objNewMember, sKey
        End If
        'return the object created
124     Set Add = objNewMember
128     Set objNewMember = Nothing
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

Add_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Add", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Get Item(vntIndexKey As Variant) As cHTTPCGI
Attribute Item.VB_UserMemId = 0
        'used when referencing an element in the collection
        'vntIndexKey contains either the Index or Key to the collection,
        'this is why it is declared as a Variant
        'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    '<EhHeader>
    On Error GoTo Item_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.Item")
    '</EhHeader>
100   Set Item = mCol(vntIndexKey)
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Item_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Item", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Property Get Count() As Long
        'used when retrieving the number of elements in the
        'collection. Syntax: Debug.Print x.Count
    '<EhHeader>
    On Error GoTo Count_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.Count")
    '</EhHeader>
100     Count = mCol.Count
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

Count_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Count", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub Remove(vntIndexKey As Variant)
        'used when removing an element from the collection
        'vntIndexKey contains either the Index or Key, which is why
        'it is declared as a Variant
        'Syntax: x.Remove(xyz)
    '<EhHeader>
    On Error GoTo Remove_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.Remove")
    '</EhHeader>
100     mCol.Remove vntIndexKey
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Sub

Remove_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Remove", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        'this property allows you to enumerate
        'this collection with the For...Each syntax
    '<EhHeader>
    On Error GoTo NewEnum_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.NewEnum")
    '</EhHeader>
100     Set NewEnum = mCol.[_NewEnum]
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Property

NewEnum_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.NewEnum", Erl, False
    Resume Next
    '</EhFooter>
End Property

'CSEH: WinUI - Custom(No Stack)
Private Sub Class_Initialize()
        'creates the collection when this class is created
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
100     Set mCol = New Collection
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set mCol = Nothing
End Sub

Public Function Clear()
        'destroys collection
    '<EhHeader>
    On Error GoTo Clear_Err
    WinUI.Debuger.CallStack.Push ("SWEBS_WinUI.cHTTPCGICol.Clear")
    '</EhHeader>
100     Set mCol = Nothing
        'creates the collection
104     Set mCol = New Collection
    '<EhFooter>
    WinUI.Debuger.CallStack.Pop
    Exit Function

Clear_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cHTTPCGICol.Clear", Erl, False
    Resume Next
    '</EhFooter>
End Function
