Attribute VB_Name = "basUtil"
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


'Registry API's
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

'Browse For Folder API's
Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'Registry
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

'Browse For Folder
Private Const MAX_PATH As Integer = 260

'Browse For Folder
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'Browse For Folder
Private Enum FolderFlags
    BIF_RETURNONLYFSDIRS = 1
    BIF_EDITBOX = &H10
    BIF_USENEWUI = &H40
End Enum

Public Function GetRegistryString(Hkey As Long, strPath As String, strValue As String) As String
Dim keyhand As Long
Dim lresult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim r As Long
Dim lValueType As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lresult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lresult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lresult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetRegistryString = Left$(strBuf, intZeroPos - 1)
            Else
                GetRegistryString = strBuf
            End If
        End If
    End If
End Function

Public Function BrowseForFolder(ByRef poOwner As Form, Optional ByRef psTitle As String = "Select A Directory", Optional ByVal flAllowNewFolder As Boolean = False, Optional psStartDir As String = "C:\") As String
'this has a bug, I know, i'll fix it some day, just not today.
Dim lpIDList As Long
Dim szTitle As String, sBuffer As String
Dim tBrowseInfo As BrowseInfo
Dim m_CurrentDirectory As String
    
    m_CurrentDirectory = psStartDir & vbNullChar
    szTitle = psTitle
    With tBrowseInfo
        .hWndOwner = poOwner.hWnd
        '.pIDLRoot = &H11
        .lpszTitle = szTitle
        .ulFlags = FolderFlags.BIF_RETURNONLYFSDIRS + FolderFlags.BIF_EDITBOX
        If flAllowNewFolder Then
            .ulFlags = .ulFlags + FolderFlags.BIF_USENEWUI
        End If
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If
End Function
