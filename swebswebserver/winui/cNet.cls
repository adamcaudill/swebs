VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNet"
Attribute VB_GlobalNameSpace = False
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
    '<EhHeader>
    On Error GoTo CurrentIP_Err
    '</EhHeader>
    Dim strResult As String

100     If WinUI.Net.IsOnline = True Then
104         strResult = WinUI.Net.PageSource("http://checkip.dyndns.org/")
108         strResult = Replace(strResult, vbCr, "")
112         strResult = Replace(strResult, vbLf, "")
116         If InStr(strResult, "<br>") = 0 Then
120             strResult = Mid$(strResult, InStr(1, strResult, "Current IP Address: "), (InStr(1, strResult, "</body>") - 1) - InStr(1, strResult, "Current IP Address: ") + 1)
            Else
124             strResult = Mid$(strResult, InStr(1, strResult, "Current IP Address: "), (InStr(1, strResult, "<br>") - 1) - InStr(1, strResult, "Current IP Address: ") + 1)
            End If
128         strResult = Replace(strResult, "Current IP Address: ", "")
132         CurrentIP = strResult
        Else
136         CurrentIP = "127.0.0.1"
        End If
    '<EhFooter>
    Exit Property

CurrentIP_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cNet.CurrentIP", Erl, False
    Resume Next
    '</EhFooter>
End Property

Public Sub LaunchURL(ByVal strURL As String)
    '<EhHeader>
    On Error GoTo LaunchURL_Err
    '</EhHeader>
100     Call ShellExecute(0, vbNullString, strURL, vbNullString, vbNullString, vbNormalFocus)
    '<EhFooter>
    Exit Sub

LaunchURL_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cNet.LaunchURL", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function PageSource(ByVal sURL As String, Optional ByVal strHeaders As String = "") As String
    '<EhHeader>
    On Error GoTo PageSource_Err
    '</EhHeader>
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    Dim lngHeaderLen As Long

100    lngHeaderLen = Len(strHeaders)
       'get the handle of the current internet connection
104    hSession = InternetOpen("User-Agent: SWEBS WinUI " & WinUI.Version & " <adam@imspire.com>", 1, vbNullString, vbNullString, 0)
       'get the handle of the url
108    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, strHeaders, lngHeaderLen, IF_NO_CACHE_WRITE, 0)
       'if we have the handle, then start reading the web page
112    If hInternet Then
           'get the first chunk & buffer it.
116        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
120        sData = sBuffer
           'if there's more data then keep reading it into the buffer
124        Do While lReturn <> 0
128            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
132            sData = sData + Mid$(sBuffer, 1, lReturn)
136            DoEvents
           Loop
       End If
        'close the URL
140    iResult = InternetCloseHandle(hInternet)
144    PageSource = sData
    '<EhFooter>
    Exit Function

PageSource_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cNet.PageSource", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Property Get IsOnline() As Boolean
    '<EhHeader>
    On Error GoTo IsOnline_Err
    '</EhHeader>
    Dim lNameLen As Long
    Dim lRetVal As Long
    Dim lConnectionFlags As Long
    Dim LPTR As Long
    Dim lNameLenPtr As Long
    Dim sConnectionName As String

100     sConnectionName = Space$(256)
104     lNameLen = 256
108     LPTR = StrPtr(sConnectionName)
112     lNameLenPtr = VarPtr(lNameLen)
116     lRetVal = InternetGetConnectedStateEx(lConnectionFlags, ByVal LPTR, ByVal lNameLen, 0&)
120     IsOnline = (lRetVal <> 0)
    '<EhFooter>
    Exit Property

IsOnline_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.cNet.IsOnline", Erl, False
    Resume Next
    '</EhFooter>
End Property
