Attribute VB_Name = "basService"
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

'API Constants
Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
' Service Control
Public Const SERVICE_CONTROL_STOP = &H1
' Service State - for CurrentState
Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7
'Service Control Manager object specific access types
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
'Service object specific access types
Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Public Function ServiceStatus(ComputerName As String, ServiceName As String) As String
    '<EhHeader>
    On Error GoTo ServiceStatus_Err
    '</EhHeader>
    Dim ServiceStat As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim hServiceStatus As Long

100     ServiceStatus = ""
104     hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
108     If hSManager <> 0 Then
112         hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
116         If hService <> 0 Then
120             hServiceStatus = QueryServiceStatus(hService, ServiceStat)
124             If hServiceStatus <> 0 Then
128                 Select Case ServiceStat.dwCurrentState
                    Case SERVICE_STOPPED
132                     ServiceStatus = "Stopped"
136                 Case SERVICE_START_PENDING
140                     ServiceStatus = "Start Pending"
144                 Case SERVICE_STOP_PENDING
148                     ServiceStatus = "Stop Pending"
152                 Case SERVICE_RUNNING
156                     ServiceStatus = "Running"
160                 Case SERVICE_CONTINUE_PENDING
164                     ServiceStatus = "Continue Pending"
168                 Case SERVICE_PAUSE_PENDING
172                     ServiceStatus = "Pause Pending"
176                 Case SERVICE_PAUSED
180                     ServiceStatus = "Paused"
                    End Select
                End If
184             CloseServiceHandle hService
            End If
188         CloseServiceHandle hSManager
        End If
    '<EhFooter>
    Exit Function

ServiceStatus_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_Main.basService.ServiceStatus", Erl, False
    Resume Next
    '</EhFooter>
End Function

Public Sub ServiceStart(ComputerName As String, ServiceName As String)
    '<EhHeader>
    On Error GoTo ServiceStart_Err
    '</EhHeader>
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

100     hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
104     If hSManager <> 0 Then
108         hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
112         If hService <> 0 Then
116             res = StartService(hService, 0, 0)
120             CloseServiceHandle hService
            End If
124         CloseServiceHandle hSManager
        End If
    '<EhFooter>
    Exit Sub

ServiceStart_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_Main.basService.ServiceStart", Erl, False
    Resume Next
    '</EhFooter>
End Sub

Public Sub ServiceStop(ComputerName As String, ServiceName As String)
    '<EhHeader>
    On Error GoTo ServiceStop_Err
    '</EhHeader>
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

100     hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
104     If hSManager <> 0 Then
108         hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
112         If hService <> 0 Then
116             res = ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus)
120             CloseServiceHandle hService
            End If
124         CloseServiceHandle hSManager
        End If
    '<EhFooter>
    Exit Sub

ServiceStop_Err:
    DisplayErrMsg Err.Description, "SWEBS_WinUI_Main.basService.ServiceStop", Erl, False
    Resume Next
    '</EhFooter>
End Sub
