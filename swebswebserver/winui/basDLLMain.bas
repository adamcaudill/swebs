Attribute VB_Name = "basDLLMain"
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

Public mWinUI As New cWinUI
Public strEventLog As String

Dim blnDestroyUI As Boolean

Public Sub Main()
    blnDestroyUI = True
End Sub

Public Sub DestroyUI()
    If blnDestroyUI = True Then
        blnDestroyUI = False
        Set mWinUI = Nothing
    End If
End Sub

Public Sub DisplayErrMsg(strMessage As String, strLocation As String, Optional strLine As String = "(Unknown)", Optional blnFatal As Boolean = False)
Dim strErrMsg As String

    If strMessage = "" Then
        strMessage = "There was an unknown error."
    End If
    strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "")
    MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
    mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.basDLLMain.DisplayErrMsg", "An error message was raised. The message was: " & strMessage
End Sub

