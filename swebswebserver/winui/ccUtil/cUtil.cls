VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

'Registry API's
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Browse For Folder API's
Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'Set foreground window by caption
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

'xp theme
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

'stop window from updating
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'prevent xp app shutdown crash. see Q309366
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'Registry
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

'Browse For Folder
Private Const MAX_PATH As Integer = 260

'xp themed
Private Const ICC_USEREX_CLASSES = &H200

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

'XP themed
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Function GetRegistryString(hKey As Long, strPath As String, strValue As String) As String
Dim keyhand As Long
Dim lresult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim R As Long
Dim lValueType As Long
    R = RegOpenKey(hKey, strPath, keyhand)
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

Public Function BrowseForFolder(Optional ByRef psTitle As String = "Select A Directory", Optional ByVal flAllowNewFolder As Boolean = False, Optional psStartDir As String = "C:\") As String
''this has a bug, I know, i'll fix it some day, just not today.
''***this needs to get replaced for the upgrade...
Dim lpIDList As Long
Dim szTitle As String, sBuffer As String
Dim tBrowseInfo As BrowseInfo
Dim m_CurrentDirectory As String

    m_CurrentDirectory = psStartDir & vbNullChar
    szTitle = psTitle
    With tBrowseInfo
        .hWndOwner = 0
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

Public Function GetTaggedData(strData As String, strTag As String, Optional strEndTag As String = "") As String
Dim lngStart As Long
Dim lngEnd As Long

    lngStart = InStr(1, strData, "<" & strTag & ">")
    If lngStart > 0 Then
        lngStart = lngStart + Len(strTag) + 2
        If strEndTag = "" Then
            lngEnd = InStr(lngStart, strData, "</" & strTag & ">")
        Else
            lngEnd = InStr(lngStart, strData, "</" & strEndTag & ">")
        End If
    End If
    If lngStart = 0 Or lngEnd = 0 Then
        GetTaggedData = ""
    Else
        GetTaggedData = Mid$(strData, lngStart, lngEnd - lngStart)
    End If
End Function

Public Function SetFocusByCaption(strCaption As String) As Boolean
Dim lngHandle As Long
Dim lngResult As Long

    lngHandle = FindWindow(vbNullString, strCaption)
    If lngHandle <> 0 Then
        lngResult = SetForegroundWindow(lngHandle)
        If lngResult = 0 Then
            SetFocusByCaption = False
        Else
            SetFocusByCaption = True
        End If
    Else
        SetFocusByCaption = False
    End If
End Function

Public Function InitCommonControlsVB() As Boolean
Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
End Function

Public Sub StopWinUpdate(Optional hWnd As Long = 0)
    Call LockWindowUpdate(hWnd)
End Sub

Public Sub LoadUser32(Optional blnLoad As Boolean = False)
Static lngUser32 As Long
    If blnLoad = True Then
        lngUser32 = LoadLibrary("shell32.dll")
    Else
        FreeLibrary lngUser32
    End If
End Sub

Public Function UrlEncode(sText As String) As String
Dim sResult As String
Dim sFinal As String
Dim sChar As String
Dim i As Long

   For i = 1 To Len(sText)
      sChar = Mid$(sText, i, 1)
      If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.@", sChar) <> 0 Then
            sResult = sResult & sChar
         ElseIf sChar = " " Then
            sResult = sResult & "+"
         ElseIf True Then
            sResult = sResult & "%" & Right$("0" & Hex$(Asc(sChar)), 2)
         End If
         If Len(sResult) > 1000 Then
            sFinal = sFinal & sResult
            sResult = ""
         End If
   Next
   UrlEncode = sFinal & sResult
End Function

Public Sub SaveRegistryString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim keyhand As Long
Dim lngResult As Long
    lngResult = RegCreateKey(hKey, strPath, keyhand)
    lngResult = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lngResult = RegCloseKey(keyhand)
End Sub

Public Function CUnescape(Source As String, Optional ForceDoubleQuote As Boolean = False) As String
' Supported escape sequences:
'
'  \b     Character 0x08 (backspace)
'  \\     Backslash
'  \n     Newline (Cr+Lf)
'  \r     Carriage return
'  \l     Line feed
'  \t     Tab
'  \"     Double-quote
'  \'     Single-quote*
'  \hnn   Hexadecimal character 0xnn

Dim lngIndex As Long
Dim strChar As String * 1
Dim strEsc As String * 1
Dim strHex As String * 2
Dim strReplace As String * 1
Dim strOutput As String

    lngIndex = 1&
    Do While lngIndex <= Len(Source)
        strChar = Mid$(Source, lngIndex, 1&)
        If (strChar <> "\") Or (lngIndex > Len(Source) - 2&) Then
            strOutput = strOutput + strChar
            lngIndex = lngIndex + 1&
        Else
            strEsc = Mid$(Source, lngIndex + 1&, 1&)
            Select Case strEsc
                Case "\"
                    strReplace = "\": lngIndex = lngIndex + 2&
                Case "b"
                    strReplace = Chr$(8&): lngIndex = lngIndex + 2&
                Case "n"
                    strReplace = vbCrLf: lngIndex = lngIndex + 2&
                Case "r"
                    strReplace = vbCr: lngIndex = lngIndex + 2&
                Case "l"
                    strReplace = vbLf: lngIndex = lngIndex + 2&
                Case "t"
                    strReplace = vbTab: lngIndex = lngIndex + 2&
                Case Chr$(34)
                    strReplace = Chr$(34): lngIndex = lngIndex + 2&
                Case "'"
                    If ForceDoubleQuote Then
                        strReplace = Chr$(34): lngIndex = lngIndex + 2&
                    Else
                        strReplace = "'": lngIndex = lngIndex + 2&
                    End If
                Case "h"
                    If lngIndex + 3& > Len(Source) Then
                        strReplace = "h"
                        lngIndex = lngIndex + 2&
                    Else
                        strHex = Mid$(Source, lngIndex + 2&, 2&)
                        If Not IsNumeric("&h" & strHex) Then
                            strReplace = "h"
                            lngIndex = lngIndex + 2&
                        Else
                            strReplace = Chr$(CLng("&h" & strHex))
                            lngIndex = lngIndex + 4&
                        End If
                    End If
                Case Else
                    strReplace = strEsc
            End Select
                strOutput = strOutput & strReplace
        End If
    Loop
    CUnescape = strOutput
End Function
