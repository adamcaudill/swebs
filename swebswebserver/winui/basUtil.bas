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

Public Function GetConfigLocation() As String
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       WinUI
' Procedure  :       GetConfigLocation
' Description:       Retrives the location of the config. XML file from the registry
'                    but for now it just assunes ./config.xml. To finish this I need
'                    to know where the location is stored in the registry.
' Created by :       Adam
' Date-Time  :       8/24/2003-1:59:20 PM
' Parameters :       none
'--------------------------------------------------------------------------------
'</CSCM>

    GetConfigLocation = strUIPath & "config.xml"

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

    GetSWSInstalled = True

End Function
