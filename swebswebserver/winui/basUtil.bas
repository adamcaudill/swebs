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

'<APIDeclare>
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'</APIDeclare>

'<LocalConst>
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&
'</LocalConst>

Public Function GetConfigLocation() As String
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       WinUI
' Procedure  :       GetConfigLocation
' Description:       Retrives the location of the config. XML file from the registry
'
'                    Location of file info:
'                    HKEY_LOCAL_MACHINE\SOFTWARE\SWS\ConfigFile
' Created by :       Adam
' Date-Time  :       8/24/2003-1:59:20 PM
' Parameters :       none
'--------------------------------------------------------------------------------
'</CSCM>
Dim strResult As String
    GetConfigLocation = GetRegistryString(&H80000002, "SOFTWARE\SWS", "ConfigFile")

End Function

Public Function GetSWSInstalled() As Boolean
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       WinUI
' Procedure  :       GetSWSInstalled
' Description:       This will check 2 things, first is to see it SWS is even installed,
'                    then it will see if the service is installed. If it's not installed
'                    then it will offer a link to the SWS home page, if the service isnt
'                    installed, it'll try to install it.
'
'                    returns true for a useable installation, false for unusable.
'
'                    for now this does nothing except return true, till I get all the info
'                    to finish this.
' Created by :       Adam
' Date-Time  :       8/24/2003-2:09:24 PM
' Parameters :       none.
'--------------------------------------------------------------------------------
'</CSCM>

    strInstalledVer = GetRegistryString(&H80000002, "SOFTWARE\SWS", "Version")
    If strInstalledVer <> "" Then
        GetSWSInstalled = True
    Else
        GetSWSInstalled = False
    End If

End Function

Private Function GetRegistryString(Hkey As Long, strPath As String, strValue As String) As String
Dim keyhand As Long
Dim datatype As Long
Dim lresult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim r As Long
Dim lValueType As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lresult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
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
