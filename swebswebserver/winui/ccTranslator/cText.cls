VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Function GetText(strString As String) As String
Dim strResult As String

    If strString <> "Lang-Maintainer" Then
        strResult = Util.GetTaggedData(strLang, "lang base=" & Chr$(34) & strString & Chr$(34), "lang")
    Else
        strResult = Util.GetTaggedData(strLang, strString)
    End If
    strResult = Util.CUnescape(strResult)
    If strResult <> "" Then
        GetText = strResult
    Else
        GetText = Util.CUnescape(strString)
    End If
End Function

Public Sub LoadLang(strLangID As String, strLangFile As String)
Dim strLangTemp As String

    If Dir$(strLangFile) <> "" Then
        strLangTemp = Space$(FileLen(strLangFile))
        Open strLangFile For Binary As 1
            Get #1, 1, strLangTemp
        Close 1
        strLang = Util.GetTaggedData(strLangTemp, "lang-code id=" & Chr$(34) & strLangID & Chr$(34), "lang-code")
        strLang = Trim$(strLang)
        strLang = Replace(strLang, vbCrLf, "")
        strLang = Replace(strLang, Chr$(9), "")
    End If
End Sub
