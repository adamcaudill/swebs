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
DWORD WINAPI TimeoutThread(LPVOID lpParam);   

DWORD WINAPI TimeoutThread(LPVOID lpParam)
{
    int SFD = (int)lpParam;
    Sleep(60000);
    closesocket(SFD);
    return true;
}

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
    char Buffer[0x1000] = "";
    DWORD dwThreadId;
    HANDLE hThread;
    
    // Before we start reading, create our timeout thread
    hThread = CreateThread( 
            NULL,																    // default security attributes 
            0,                           										    // use default stack size  
            TimeoutThread,                 									    // thread function 
            &SFD,                											    // argument to thread function 
            0,                           										    // use default creation flags 
            &dwThreadId
            );                												        // returns the thread identifier 
		
        if (hThread != NULL)												        // If the thread was created, destroy it
	    {			    
            CloseHandle( hThread );
	    }

    int X = recv(SFD, Buffer, 0x1000, 0);
    while ( X > 0 )
    {
        Buffer[X] = '\0';
        Temp += Buffer;
        if (X >= 0x1000)
        {
            X = recv(SFD, Buffer, 0x1000, 0);
        }
        else break;
    }
    
    Temp = Buffer;
    return Temp;
}

};