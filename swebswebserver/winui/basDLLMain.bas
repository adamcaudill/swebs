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
    '<EhHeader>
    On Error GoTo Main_Err
    '</EhHeader>
100     blnDestroyUI = True
    '<EhFooter>
    Exit Sub

Main_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.basDLLMain.Main", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub DestroyUI()
    '<EhHeader>
    On Error GoTo DestroyUI_Err
    '</EhHeader>
100     If blnDestroyUI = True Then
104         blnDestroyUI = False
108         Set mWinUI = Nothing
        End If
    '<EhFooter>
    Exit Sub

DestroyUI_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.basDLLMain.DestroyUI", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub DisplayErrMsg(strMessage As String, strLocation As String, Optional strLine As String = "(Unknown)", Optional blnFatal As Boolean = False)
    '<EhHeader>
    On Error GoTo DisplayErrMsg_Err
    '</EhHeader>
    Dim strErrMsg As String

100     If strMessage = "" Then
104         strMessage = "There was an unknown error."
        End If
108     strErrMsg = "This application has encountered a error: " & vbCrLf & vbCrLf & "Error: '" & strMessage & "'" & vbCrLf & "Location: " & strLocation & " at line: " & strLine & vbCrLf & vbCrLf & "Contact ADAM@IMSPIRE.COM to report this error." & IIf(blnFatal = True, vbCrLf & vbCrLf & "This error is fatal, this program will now close.", "")
112     MsgBox strErrMsg, vbApplicationModal + vbCritical + vbOKOnly, "SWEBS System Error"
116     mWinUI.EventLog.AddEvent "SWEBS_WinUI_DLL.basDLLMain.DisplayErrMsg", "An error message was raised. The message was: " & strMessage
    '<EhFooter>
    Exit Sub

DisplayErrMsg_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.basDLLMain.DisplayErrMsg", Erl, False
    Resume Next
    '</EhFooter>
End Sub

