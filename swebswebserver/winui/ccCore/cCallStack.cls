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
Option Base 1

Private Const BUFFER_SIZE As Long = 25

Private Items() As String
Private Index As Long

Public Property Get Count() As Long
    Count = Index
End Property

Public Sub Clear()
    ReDim Items(BUFFER_SIZE)
    Index = 0
End Sub

Public Sub Push(Item As String)
    If UBound(Items) = Index + 1 Then
        ReDim Preserve Items(Index + BUFFER_SIZE)
    End If
    
    Index = Index + 1
    Items(Index) = Item
    If gCore.Debuger.PerfMon.Enabled = True Then
        gCore.Debuger.PerfMon.Add Item
    End If
End Sub

Public Function Pop() As String
    Dim l As Long
    
    If Index > 0 Then
        Pop = Items(Index)
        If gCore.Debuger.PerfMon.Enabled = True Then
            gCore.Debuger.PerfMon.Remove Items(Index)
        End If
        Index = Index - 1
    End If
    
    If (UBound(Items) - Index) \ BUFFER_SIZE > 1 Then
        ReDim Preserve Items(UBound(Items) - BUFFER_SIZE)
    End If
End Function

Public Function Peek(lngIndex As Long) As String
    If Index > 0 And lngIndex > 0 And lngIndex <= Index Then
        Peek = Items(lngIndex)
    End If
End Function

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Erase Items
End Sub
