VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNet"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'PageSource
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'LaunchURL API
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'IsOnline API
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (lpdwFlags As Long, lpszConnectionName As Long, dwNameLen As Long, ByVal dwReserved As Long) As Long

'PageSource
Private Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 256

Public Property Get CurrentIP() As String
Dim strResult As String

    If IsOnline = True Then
        strResult = PageSource("http://checkip.dyndns.org/")
        strResult = Replace(strResult, vbCr, "")
        strResult = Replace(strResult, vbLf, "")
        If InStr(strResult, "<br>") = 0 Then
            strResult = Mid$(strResult, InStr(1, strResult, "Current IP Address: "), (InStr(1, strResult, "</body>") - 1) - InStr(1, strResult, "Current IP Address: ") + 1)
        Else
            strResult = Mid$(strResult, InStr(1, strResult, "Current IP Address: "), (InStr(1, strResult, "<br>") - 1) - InStr(1, strResult, "Current IP Address: ") + 1)
        End If
        strResult = Replace(strResult, "Current IP Address: ", "")
        CurrentIP = strResult
    Else
        CurrentIP = "127.0.0.1"
    End If
End Property

Public Sub LaunchURL(ByVal strURL As String)
    Call ShellExecute(0, vbNullString, strURL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function PageSource(ByVal sURL As String, Optional ByVal strHeaders As String = "") As String
Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
Dim hInternet As Long, hSession As Long, lReturn As Long
Dim lngHeaderLen As Long

   lngHeaderLen = Len(strHeaders)
   'get the handle of the current internet connection
   hSession = InternetOpen("User-Agent: SWEBS WinUI " & WinUI.Version & " <adam@imspire.com>", 1, vbNullString, vbNullString, 0)
   'get the handle of the url
   If hSession Then hInternet = InternetOpenUrl(hSession, sURL, strHeaders, lngHeaderLen, IF_NO_CACHE_WRITE, 0)
   'if we have the handle, then start reading the web page
   If hInternet Then
       'get the first chunk & buffer it.
       iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
       sData = sBuffer
       'if there's more data then keep reading it into the buffer
       Do While lReturn <> 0
           iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
           sData = sData + Mid$(sBuffer, 1, lReturn)
           DoEvents
       Loop
   End If
    'close the URL
   iResult = InternetCloseHandle(hInternet)
   PageSource = sData
End Function

Public Property Get IsOnline() As Boolean
Dim lNameLen As Long
Dim lRetVal As Long
Dim lConnectionFlags As Long
Dim LPTR As Long
Dim lNameLenPtr As Long
Dim sConnectionName As String

    sConnectionName = Space$(256)
    lNameLen = 256
    LPTR = StrPtr(sConnectionName)
    lNameLenPtr = VarPtr(lNameLen)
    lRetVal = InternetGetConnectedStateEx(lConnectionFlags, ByVal LPTR, ByVal lNameLen, 0&)
    IsOnline = (lRetVal <> 0)
End Property
