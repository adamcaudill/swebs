#ifndef SWEBSSOCKETHPP
#define SWEBSSOCKETHPP 1
//----------------------------------------------------------------------------------
//			SWEBSSOCKET.hpp
//			-------------------
//			This header declares some of the socket functions used by the SWEBS.
//			It provides a level of abstration over the WinSock and BSD socket I/O
//			functions. The definitions of these functions are in SWEBSSocket.lib.
//			Remember these functions use the SWEBS namespace!
//
//----------------------------------------------------------------------------------																			
//			INCLUDES
//----------------------------------------------------------------------------------
#include <string>

using namespace std;
//----------------------------------------------------------------------------------
//			SWEBSSOCKET Namespace		
//----------------------------------------------------------------------------------
namespace SWEBSSocket
{
	unsigned int Send(int SFD, string Data, int Length);						    // Sends Data to socket SFD and returns bytes sent
	string Recieve(int SFD);											            // Returns Size ammount of data from SFD
}

//----------------------------------------------------------------------------------
#endif