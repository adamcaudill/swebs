Attribute VB_Name = "basMain"
Option Explicit

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

'<GlobalVars>
Public strConfigFile As String
Public strUIPath As String
'</GlobalVars>

Public Sub Main()
    strUIPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    If GetSWSInstalled = False Then
        MsgBox "This application will now exit.", vbCritical + vbOKOnly + vbApplicationModal
        End
    End If
    strConfigFile = GetConfigLocation
    Load frmMain
    DoEvents
    frmMain.Show
End Sub
