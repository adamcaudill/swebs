VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCallStack"
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

Private Const BUFFER_SIZE As Long = 100

Private Items() As Variant
Private Index As Long

Public Property Get Count() As Long
    '<EhHeader>
    On Error GoTo Count_Err
    '</EhHeader>

100    Count = Index

    '<EhFooter>
    Exit Property

Count_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Count", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub Clear()
    '<EhHeader>
    On Error GoTo Clear_Err
    '</EhHeader>

100    ReDim Items(BUFFER_SIZE - 1)
104    Index = 0
   
    '<EhFooter>
    Exit Sub

Clear_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Clear", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub Push(Item As Variant)
    '<EhHeader>
    On Error GoTo Push_Err
    '</EhHeader>

100    If UBound(Items) = Index Then
104       ReDim Preserve Items(Index + BUFFER_SIZE)
       End If

108    If IsObject(Item) Then
112       Set Items(Index) = Item
       Else
116       Items(Index) = Item
       End If
120    Index = Index + 1

    '<EhFooter>
    Exit Sub

Push_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Push", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function Pop() As Variant
    '<EhHeader>
    On Error GoTo Pop_Err
    '</EhHeader>

       Dim l As Long
   
100    If Index > 0 Then
104       Index = Index - 1
108       If IsObject(Items(Index)) Then
112          Set Pop = Items(Index)
          Else
116          Pop = Items(Index)
          End If
       End If

120    If (UBound(Items) - Index) \ BUFFER_SIZE > 0 Then
124       ReDim Preserve Items(UBound(Items) - BUFFER_SIZE)
       End If
   
    '<EhFooter>
    Exit Function

Pop_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Pop", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function Peek(lngIndex As Long) As Variant
    '<EhHeader>
    On Error GoTo Peek_Err
    '</EhHeader>

100    If Index > 0 And lngIndex > 0 And lngIndex <= Index Then
104       Peek = Items(lngIndex - 1)
108       If IsObject(Items(lngIndex - 1)) Then
112          Set Peek = Items(lngIndex - 1)
          Else
116          Peek = Items(lngIndex - 1)
          End If
       End If

    '<EhFooter>
    Exit Function

Peek_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Peek", Erl, False
    Resume Next
    '</EhFooter>
End Function

Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>

100    Clear

    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cCallStack.Class_Initialize", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

   Erase Items

End Sub