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

'Registry API's
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

'Browse For Folder API's
Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'Set foreground window by caption
Private Declare Function SetForegroundWindow Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

'xp theme
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

'stop window from updating
Private Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long

'prevent xp app shutdown crash. see Q309366
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'error handleing stuff
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

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

Public Function GetRegistryString(Hkey As Long, strPath As String, strValue As String) As String
    '<EhHeader>
    On Error GoTo GetRegistryString_Err
    '</EhHeader>
    Dim keyhand As Long
    Dim lresult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType As Long
100     r = RegOpenKey(Hkey, strPath, keyhand)
104     lresult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
108     If lValueType = REG_SZ Then
112         strBuf = String$(lDataBufSize, " ")
116         lresult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
120         If lresult = ERROR_SUCCESS Then
124             intZeroPos = InStr(strBuf, Chr$(0))
128             If intZeroPos > 0 Then
132                 GetRegistryString = Left$(strBuf, intZeroPos - 1)
                Else
136                 GetRegistryString = strBuf
                End If
            End If
        End If
    '<EhFooter>
    Exit Function

GetRegistryString_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.GetRegistryString", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function BrowseForFolder(ByRef poOwner As Variant, Optional ByRef psTitle As String = "Select A Directory", Optional ByVal flAllowNewFolder As Boolean = False, Optional psStartDir As String = "C:\") As String
    'this has a bug, I know, i'll fix it some day, just not today.
    '<EhHeader>
    On Error GoTo BrowseForFolder_Err
    '</EhHeader>
    Dim lpIDList As Long
    Dim szTitle As String, sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    Dim m_CurrentDirectory As String
    
100     m_CurrentDirectory = psStartDir & vbNullChar
104     szTitle = psTitle
108     With tBrowseInfo
112         .hWndOwner = 0
            '.pIDLRoot = &H11
116         .lpszTitle = szTitle
120         .ulFlags = FolderFlags.BIF_RETURNONLYFSDIRS + FolderFlags.BIF_EDITBOX
124         If flAllowNewFolder Then
128             .ulFlags = .ulFlags + FolderFlags.BIF_USENEWUI
            End If
        End With
132     lpIDList = SHBrowseForFolder(tBrowseInfo)
136     If (lpIDList) Then
140         sBuffer = Space$(MAX_PATH)
144         SHGetPathFromIDList lpIDList, sBuffer
148         sBuffer = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
152         BrowseForFolder = sBuffer
        Else
156         BrowseForFolder = ""
        End If
    '<EhFooter>
    Exit Function

BrowseForFolder_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.BrowseForFolder", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function GetTaggedData(strData As String, strTag As String) As String
    '<EhHeader>
    On Error GoTo GetTaggedData_Err
    '</EhHeader>
    Dim lngStart As Long
    Dim lngEnd As Long

100     lngStart = (InStr(1, strData, "<" & strTag & ">") + Len(strTag) + 2)
104     lngEnd = InStr(1, strData, "</" & strTag & ">")
108     If lngStart = 0 Or lngEnd = 0 Then
112         GetTaggedData = ""
        Else
116         GetTaggedData = Mid$(strData, lngStart, lngEnd - lngStart)
        End If
    '<EhFooter>
    Exit Function

GetTaggedData_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.GetTaggedData", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function SetFocusByCaption(strCaption As String) As Boolean
    '<EhHeader>
    On Error GoTo SetFocusByCaption_Err
    '</EhHeader>
    Dim lngHandle As Long
    Dim lngResult As Long

100     lngHandle = FindWindow(vbNullString, strCaption)
104     If lngHandle <> 0 Then
108         lngResult = SetForegroundWindow(lngHandle)
112         If lngResult = 0 Then
116             SetFocusByCaption = False
            Else
120             SetFocusByCaption = True
            End If
        Else
124         SetFocusByCaption = False
        End If
    '<EhFooter>
    Exit Function

SetFocusByCaption_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.SetFocusByCaption", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function InitCommonControlsVB() As Boolean
    '<EhHeader>
    On Error GoTo InitCommonControlsVB_Err
    '</EhHeader>
    Dim iccex As tagInitCommonControlsEx
100    With iccex
104        .lngSize = LenB(iccex)
108        .lngICC = ICC_USEREX_CLASSES
       End With
112    InitCommonControlsEx iccex
116    InitCommonControlsVB = (Err.Number = 0)
    '<EhFooter>
    Exit Function

InitCommonControlsVB_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.InitCommonControlsVB", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Sub StopWinUpdate(Optional hWnd As Long = 0)
    '<EhHeader>
    On Error GoTo StopWinUpdate_Err
    '</EhHeader>
100     Call LockWindowUpdate(hWnd)
    '<EhFooter>
    Exit Sub

StopWinUpdate_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.StopWinUpdate", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub LoadUser32(Optional blnLoad As Boolean = False)
    '<EhHeader>
    On Error GoTo LoadUser32_Err
    '</EhHeader>
    Static lngUser32 As Long
100     If blnLoad = True Then
104         lngUser32 = LoadLibrary("shell32.dll")
        Else
108         FreeLibrary lngUser32
        End If
    '<EhFooter>
    Exit Sub

LoadUser32_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.LoadUser32", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Function UrlEncode(sText As String) As String
    '<EhHeader>
    On Error GoTo UrlEncode_Err
    '</EhHeader>
    Dim sResult As String
    Dim sFinal As String
    Dim sChar As String
    Dim i As Long

100    For i = 1 To Len(sText)
104       sChar = Mid$(sText, i, 1)
108       If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.@", sChar) <> 0 Then
112             sResult = sResult & sChar
116          ElseIf sChar = " " Then
120             sResult = sResult & "+"
124          ElseIf True Then
128             sResult = sResult & "%" & Right$("0" & Hex$(Asc(sChar)), 2)
             End If
132          If Len(sResult) > 1000 Then
136             sFinal = sFinal & sResult
140             sResult = ""
             End If
       Next
144    UrlEncode = sFinal & sResult
    '<EhFooter>
    Exit Function

UrlEncode_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.UrlEncode", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Sub SaveRegistryString(Hkey As Long, strPath As String, strValue As String, strData As String)
    '<EhHeader>
    On Error GoTo SaveRegistryString_Err
    '</EhHeader>
    Dim keyhand As Long
    Dim lngResult As Long
100     lngResult = RegCreateKey(Hkey, strPath, keyhand)
104     lngResult = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
108     lngResult = RegCloseKey(keyhand)
    '<EhFooter>
    Exit Sub

SaveRegistryString_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.SaveRegistryString", Erl, False
    Resume Next
    '</EhFooter>
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
    '<EhHeader>
    On Error GoTo CUnescape_Err
    '</EhHeader>

    Dim lngIndex As Long
    Dim strChar As String * 1
    Dim strEsc As String * 1
    Dim strHex As String * 2
    Dim strReplace As String * 1
    Dim strOutput As String

100     lngIndex = 1&
104     Do While lngIndex <= Len(Source)
108         strChar = Mid$(Source, lngIndex, 1&)
112         If (strChar <> "\") Or (lngIndex > Len(Source) - 2&) Then
116             strOutput = strOutput + strChar
120             lngIndex = lngIndex + 1&
            Else
124             strEsc = Mid$(Source, lngIndex + 1&, 1&)
128             Select Case strEsc
                    Case "\"
132                     strReplace = "\": lngIndex = lngIndex + 2&
136                 Case "b"
140                     strReplace = Chr$(8&): lngIndex = lngIndex + 2&
144                 Case "n"
148                     strReplace = vbCrLf: lngIndex = lngIndex + 2&
152                 Case "r"
156                     strReplace = vbCr: lngIndex = lngIndex + 2&
160                 Case "l"
164                     strReplace = vbLf: lngIndex = lngIndex + 2&
168                 Case "t"
172                     strReplace = vbTab: lngIndex = lngIndex + 2&
176                 Case Chr$(34)
180                     strReplace = Chr$(34): lngIndex = lngIndex + 2&
184                 Case "'"
188                     If ForceDoubleQuote Then
192                         strReplace = Chr$(34): lngIndex = lngIndex + 2&
                        Else
196                         strReplace = "'": lngIndex = lngIndex + 2&
                        End If
200                 Case "h"
204                     If lngIndex + 3& > Len(Source) Then
208                         strReplace = "h"
212                         lngIndex = lngIndex + 2&
                        Else
216                         strHex = Mid$(Source, lngIndex + 2&, 2&)
220                         If Not IsNumeric("&h" & strHex) Then
224                             strReplace = "h"
228                             lngIndex = lngIndex + 2&
                            Else
232                             strReplace = Chr$(CLng("&h" & strHex))
236                             lngIndex = lngIndex + 4&
                            End If
                        End If
240                 Case Else
244                     strReplace = strEsc
                End Select
248                 strOutput = strOutput & strReplace
            End If
        Loop
252     CUnescape = strOutput
    '<EhFooter>
    Exit Function

CUnescape_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.CUnescape", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Function GetWin32ErrDesc(ErrorCode As Long) As String
    'this isn't used now, but it'll be used someday..
    '<EhHeader>
    On Error GoTo GetWin32ErrDesc_Err
    '</EhHeader>
    Dim lngRet As Long
    Dim strAPIError As String

100     strAPIError = String$(2048, " ")
104     lngRet = FormatMessage(&H1000, ByVal 0&, ErrorCode, 0, strAPIError, Len(strAPIError), 0)
108     strAPIError = Left$(strAPIError, lngRet)
112     GetWin32ErrDesc = strAPIError
    '<EhFooter>
    Exit Function

GetWin32ErrDesc_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_DLL.cUtil.GetWin32ErrDesc", Erl, False
    Resume Next
    '</EhFooter>
End Function

