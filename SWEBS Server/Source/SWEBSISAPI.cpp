//----------------------------------------------------------------------------------
//			SWEBSUTILITIES.cpp
//			-------------------
//			This source is a bunch of functions that dont belong anywhere else,
//			and are used thoughout the server. They perform trivial yet
//			necessary jobs. They have been compiled into the library 
//			SWEBSUtilities.lib
//  
//----------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include <string>
#include <sstream>
#include "../Include/SWEBSISAPI.hpp"
#include "../Include/SWEBSConnection.hpp"
#include "../Include/SWEBSSocket.hpp"
#include "../Include/SWEBSUtilities.hpp"

using namespace std;

class CONNECTION;
//--------------------------------------------------------------------------------------------
//          CONNECTION::ReadClientExport()
//--------------------------------------------------------------------------------------------
BOOL WINAPI ReadClientExport(HCONN ConnID,
                                         LPVOID lpvBuffer,
                                         LPDWORD lpdwSize)
{
    char szPostData[1024];

    CONNECTION * Connection = (CONNECTION*)ConnID;

    Connection->PostData.getline(szPostData, 1024);
    
    if (*lpdwSize < (strlen(szPostData)) )
    {
        memcpy(lpvBuffer, &szPostData, *lpdwSize );
    }
    else
    {
        memcpy (lpvBuffer, &szPostData, strlen(szPostData));                     // Give them all the POST data
    }
    *lpdwSize = strlen(szPostData) + 1;
    
    return true;
}

//---------------------------------------------------------------------------------------------
//          CONNECTION::WriteClientExport
//---------------------------------------------------------------------------------------------
BOOL WINAPI WriteClientExport(HCONN      ConnID,
                                          LPVOID     Buffer,
                                          LPDWORD    lpdwBytes,
                                          DWORD      dwReserved )
{
    CONNECTION * Connection = (CONNECTION*)ConnID;
    return SWEBSSocket::Send(Connection->SFD , (char *) Buffer, strlen((char *)Buffer) );
}

//---------------------------------------------------------------------------------------------
//          CONNECTION::GetServerVariableExport
//---------------------------------------------------------------------------------------------
BOOL WINAPI GetServerVariableExport(HCONN hConn, 
                                                    LPSTR lpszVariableName, 
                                                    LPVOID lpvBuffer, 
                                                    LPDWORD lpdwSize)
{
    CONNECTION * Connection = (CONNECTION*)hConn;

    char szTemp[1024];
    char * pcTemp = "UNKNOWNVALUE";
    
    strncpy(szTemp, pcTemp, *lpdwSize);
    
    if ( !strcmpi(lpszVariableName, "ALL_HTTP") )
    {
        string ALL_HTTP;
        // Do some ALL_HTTP stuff here!
    }
    else if ( !strcmpi(lpszVariableName, "ALL_RAW"))
    {
        string ALL_RAW;
        // Do raw stuff here
    }
    else
    {
        strncpy(szTemp, Connection->CGIVariables.HTTP_MAP[lpszVariableName].c_str(), *lpdwSize);
    }
    
    *lpdwSize = strlen(szTemp) + 1;
    szTemp[*lpdwSize] = '\0';
    
    memcpy(lpvBuffer, &szTemp, *lpdwSize);
    if (*lpdwSize >= 1)
    {
        return true;
    }
    else return false;
}

//---------------------------------------------------------------------------------------------
//          CONNECTION::ServerSupportFunctionExport
//---------------------------------------------------------------------------------------------
BOOL WINAPI ServerSupportFunctionExport(HCONN      hConn,
                                           DWORD      dwHSERequest,
                                           LPVOID     lpvBuffer,
                                           LPDWORD    lpdwSize,
                                           LPDWORD    lpdwDataType)
{
    CONNECTION * Connection = (CONNECTION*)hConn;

    switch (dwHSERequest)
    {
    case HSE_REQ_SEND_RESPONSE_HEADER_EX:
        Connection->Status = StringToInt((char *)lpvBuffer);
        return true;
    break;
    default:
        return false;
    }
}

//---------------------------------------------------------------------------------------------