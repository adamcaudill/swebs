//----------------------------------------------------------------------------------
//			SWEBSSOCKET.cpp
//			-------------------
//			This source file defines some of the socket functions used by the SWEBS.
//			It provides a level of abstration over the WinSock and BSD socket I/O
//			functions. The definitions of these functions are in SWEBSSocket.lib.
//			Remember these functions use the SWEBS namespace!
//
//----------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include "../Include/SWEBSGlobals.hpp"
#include "../Include/SWEBSSocket.hpp"
#include <winsock.h>
#include <string>
#include <iostream>

using namespace std;

#pragma comment(lib, "wsock32.lib")
//----------------------------------------------------------------------------------
//			Send()
//          Our own version of send(), so that we can keep track of whats being sent
//----------------------------------------------------------------------------------
namespace SWEBSSocket
{
unsigned int Send(int SFD, string Text, int Length)
{
    // Some calls to send() have additional info that we want to get rid of, so we just ignore them here
    int NumSent = send (SFD, Text.c_str(), Length, 0);
    
    SWEBSGlobals.TotalBytesSent += NumSent;
    SWEBSGlobals.BytesSent += NumSent;
    
    if (NumSent <= 0)
        return false;
    else return true;
}

//----------------------------------------------------------------------------------
//			Recieve()
//----------------------------------------------------------------------------------
string Recieve(int SFD)
{
    string Temp;
    char Buffer[1023];
    int X = recv(SFD, Buffer, 1023, 0);
    while ( X > 0 )
    {
        Buffer[X] = '\0';
        Temp += Buffer;
        if (X >= 1023)
        {
            X = recv(SFD, Buffer, 1023, 0);
        }
        else break;
    }
    //MessageBox(NULL, Temp.c_str(), "SWEBS", MB_OK);
    return Temp;
}

};