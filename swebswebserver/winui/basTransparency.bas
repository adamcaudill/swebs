Attribute VB_Name = "basTransparency"
'CSEH: Core - Custom
'***************************************************************************
'
' SWEBS/Core
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

Private Declare Function CreateRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long

Private Declare Function CombineRgn Lib "gdi32" ( _
    ByVal hDestRgn As Long, _
    ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long _
) As Long

Private Declare Function OffsetRgn Lib "gdi32" ( _
    ByVal hRgn As Long, _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean _
) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Private Declare Function GetPixel Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Private Declare Function GetSystemMetrics Lib "user32" ( _
    ByVal nIndex As Long _
) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long _
) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8

Public Function FormRegion(frmForm As Form) As Long
Dim i As Long, ii As Long, lngPicWidth As Long, lngPicHeight As Long, lngTitleHeight As Long
Dim lngBorderWidth As Long, lngPicRegion As Long, lngPixelRegion As Long, lngPixelColor As Long
Dim lngPicDC As Long, lngPicTempBMP As Long, lngPicTransColor As Long, lngOriginalRgn As Long
Dim lngRetVal As Long

    lngPicWidth = frmForm.ScaleX(frmForm.Picture.Width, vbTwips, vbPixels)
    lngPicHeight = frmForm.ScaleY(frmForm.Picture.Height, vbTwips, vbPixels)
    
    lngPicRegion = CreateRectRgn(0, 0, lngPicWidth, lngPicHeight)
    
    lngPicDC = CreateCompatibleDC(frmForm.hdc)
    lngPicTempBMP = SelectObject(lngPicDC, frmForm.Picture.Handle)
    
    lngPicTransColor = GetPixel(lngPicDC, 0, 0)
    For i = 0 To lngPicHeight
        For ii = 0 To lngPicWidth
            lngPixelColor = GetPixel(lngPicDC, i, ii)
            If lngPixelColor = lngPicTransColor Then
                lngPixelRegion = CreateRectRgn(i, ii, i + 1, ii + 1)
                CombineRgn lngPicRegion, lngPicRegion, lngPixelRegion, RGN_XOR
                DeleteObject lngPixelRegion
            End If
        Next ii
    Next i
    
    SelectObject lngPicDC, lngPicTempBMP
    DeleteDC lngPicDC
    DeleteObject lngPicTempBMP
    
    lngOriginalRgn = SetWindowRgn(frmForm.hWnd, lngPicRegion, True)
    FormRegion = lngOriginalRgn
End Function

Public Sub FormRemoveRegion(hWnd As Long, lngOriginalRgn As Long)
    DeleteObject SetWindowRgn(hWnd, lngOriginalRgn, True)
End Sub
