Attribute VB_Name = "basExceptionFilter"
'CSEH: WinUI - Custom(No Stack)
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

Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Private Declare Sub CopyExceptionRecord Lib "kernel32" Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, ByVal LPEXCEPTION_RECORD As Long, ByVal lngBytes As Long)

Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Private Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type

Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Private Type EXCEPTION_DEBUG_INFO
    pExceptionRecord As EXCEPTION_RECORD
    dwFirstChance As Long
End Type

Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type

Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const EXCEPTION_CONTROL_C_EXIT = &HC000013A

Private Function GetExceptionText(ByVal ExceptionCode As Long) As String
    '<EhHeader>
    On Error GoTo GetExceptionText_Err
    '</EhHeader>
    Dim strExceptionString As String

100     Select Case ExceptionCode
            Case EXCEPTION_ACCESS_VIOLATION
104             strExceptionString = "Access Violation"
108         Case EXCEPTION_DATATYPE_MISALIGNMENT
112             strExceptionString = "Data Type Misalignment"
116         Case EXCEPTION_BREAKPOINT
120             strExceptionString = "Breakpoint"
124         Case EXCEPTION_SINGLE_STEP
128             strExceptionString = "Single Step"
132         Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
136             strExceptionString = "Array Bounds Exceeded"
140         Case EXCEPTION_FLT_DENORMAL_OPERAND
144             strExceptionString = "Float Denormal Operand"
148         Case EXCEPTION_FLT_DIVIDE_BY_ZERO
152             strExceptionString = "Divide By Zero"
156         Case EXCEPTION_FLT_INEXACT_RESULT
160             strExceptionString = "Floating Point Inexact Result"
164         Case EXCEPTION_FLT_INVALID_OPERATION
168             strExceptionString = "Invalid Operation"
172         Case EXCEPTION_FLT_OVERFLOW
176             strExceptionString = "Float Overflow"
180         Case EXCEPTION_FLT_STACK_CHECK
184             strExceptionString = "Float Stack Check"
188         Case EXCEPTION_FLT_UNDERFLOW
192             strExceptionString = "Float Underflow"
196         Case EXCEPTION_INT_DIVIDE_BY_ZERO
200             strExceptionString = "Integer Divide By Zero"
204         Case EXCEPTION_INT_OVERFLOW
208             strExceptionString = "Integer Overflow"
212         Case EXCEPTION_PRIVILEGED_INSTRUCTION
216             strExceptionString = "Privileged Instruction"
220         Case EXCEPTION_IN_PAGE_ERROR
224             strExceptionString = "In Page Error"
228         Case EXCEPTION_ILLEGAL_INSTRUCTION
232             strExceptionString = "Illegal Instruction"
236         Case EXCEPTION_NONCONTINUABLE_EXCEPTION
240             strExceptionString = "Non Continuable Exception"
244         Case EXCEPTION_STACK_OVERFLOW
248             strExceptionString = "Stack Overflow"
252         Case EXCEPTION_INVALID_DISPOSITION
256             strExceptionString = "Invalid Disposition"
260         Case EXCEPTION_GUARD_PAGE_VIOLATION
264             strExceptionString = "Guard Page Violation"
268         Case EXCEPTION_INVALID_HANDLE
272             strExceptionString = "Invalid Handle"
276         Case EXCEPTION_CONTROL_C_EXIT
280             strExceptionString = "Control-C Exit"
284         Case Else
288             strExceptionString = "Unknown (&H" & Right$("00000000" & Hex$(ExceptionCode), 8) & ")"
        End Select
292     GetExceptionText = strExceptionString
    '<EhFooter>
    Exit Function

GetExceptionText_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basExceptionFilter.GetExceptionText", Erl, False
    Resume Next
    '</EhFooter>
End Function

'CSEH: Skip
Private Function ExceptionFilter(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
Dim Rec As EXCEPTION_RECORD
Dim strException As String

    Rec = ExceptionPtrs.pExceptionRecord
    Do Until Rec.pExceptionRecord = 0
        CopyExceptionRecord Rec, Rec.pExceptionRecord, Len(Rec)
    Loop
    strException = GetExceptionText(Rec.ExceptionCode)
    Err.Raise 10000, "ExceptionFilter", strException
End Function

Public Sub SetExceptionFilter(blnEnable As Boolean)
    '<EhHeader>
    On Error GoTo SetExceptionFilter_Err
    '</EhHeader>
100     If blnEnable = True Then
104         Call SetUnhandledExceptionFilter(AddressOf ExceptionFilter)
        Else
108         Call SetUnhandledExceptionFilter(0)
        End If
    '<EhFooter>
    Exit Sub

SetExceptionFilter_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI.basExceptionFilter.SetExceptionFilter", Erl, False
    Resume Next
    '</EhFooter>
End Sub

