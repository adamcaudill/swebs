//---------------------------------------------------------------------------------------------
//  SWEBSEngine
//  ----------------------
//  This is the main SWEBS processing engine, as called by the SWEBS Web Server service. This
//  application is the main workhorse of the SWEBS Web Server.
//
//  Copyright SWEBS Development team 2003
//---------------------------------------------------------------------------------------------

//---------------------------------------------------------------------------------------------
//			Includes
//---------------------------------------------------------------------------------------------
#pragma warning(disable:4786)
#pragma warning(disable:4089)
#include <windows.h>
#include <winsock.h>
#include <string>
#include "connection.hpp"
#include "stats.hpp"
#include "resource.h"

using namespace std;

#pragma comment(lib, "wsock32.lib")
// The first two of these files come from the Chilkat XML library, freely downloadable from
//  www.xml-parser.com
#pragma comment(lib, "ChilkatRelSt.lib")
#pragma comment(lib, "CkBaseRelSt.lib")
#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "rpcrt4.lib")
#pragma comment(lib, "crypt32.lib")

#pragma warning(disable:4786)

//---------------------------------------------------------------------------------------------
//			Function Declarations
//---------------------------------------------------------------------------------------------
LRESULT CALLBACK WinProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam); 
void TestLog(string);
DWORD WINAPI ProcessRequest(LPVOID lpParam );
DWORD WINAPI AcceptThreadProc(LPVOID lpParam);

//---------------------------------------------------------------------------------------------
//			Globals
//---------------------------------------------------------------------------------------------
bool SERVER_STOP = false;                                                           // Flag for if the server is running

int ReturnCode;                                                                     // Number for main() to return, can be set from any function

const int SWEBS_RETURN_UNKNOWN          = 0x00;                                     // Unknown error occured
const int SWEBS_RETURN_SUCCESS          = 0x01;                                     // Server ran fine
const int SWEBS_RETURN_COULDNOTBIND     = 0x02;                                     // Could not bind() to port
const int SWEBS_RETURN_CONFIGNOTLOADED  = 0x03;                                     // Could not load config file
const int SWEBS_RETURN_COULDNOTLISTEN   = 0x04;                                     // Could not listen()
const int SWEBS_RETURN_COULDNOTACCEPT   = 0x05;                                     // Could not accept()

const int SWEBS_NEW_ACCEPT = 1000024;
struct ARGUMENT
{
	int SFD;
	struct sockaddr_in CLA;
};

//---------------------------------------------------------------------------------------------
//			WinMain()
//---------------------------------------------------------------------------------------------
int WINAPI WinMain(HINSTANCE hinstance, HINSTANCE hprev, PSTR cmdline, int ishow)
{
	//-----------------------------------------------------------------------------------------
	// Step 1: Do stuff we must do as a windows application
	//-----------------------------------------------------------------------------------------
	WSADATA wsaData;
    WSAStartup(MAKEWORD(1,1), &wsaData);

    MSG Message;                                                                    // Current message
    WNDCLASSEX WndClassEx = {0};													// Window Class
	WndClassEx.cbSize = sizeof(WndClassEx);											// Size of itself
	WndClassEx.lpfnWndProc = WinProc;												// Message Handler
	WndClassEx.hInstance = hinstance;												// Instance of this window class
    WndClassEx.lpszClassName = "SWEBSWindowClass";									// Class Name

	//-----------------------------------------------------------------------------------------
	// Step 2: Set up options
	//-----------------------------------------------------------------------------------------
	// These are default settings, incase the configuration file is corrupt
	Options.Timeout = 20;
	Options.ErrorLog = "C:\\SWS\\Errorlog.log";

	// Read the real settings from the config file
	bool ReadConfig = Options.ReadSettings();

	if (ReadConfig == false)
	{
		// The configuration file had errors.
		Options.LogError("Warning: Could not load configuration file properly");
        ReturnCode = SWEBS_RETURN_CONFIGNOTLOADED;
		return ReturnCode;
	}

    // Set non request-specific variables
    Options.CGIVariables.GATEWAY_INTERFACE = "CGI/1.1";
    Options.CGIVariables.SERVER_NAME = Options.Servername;
    Options.CGIVariables.SERVER_SOFTWARE = Options.Servername;

	//----------------------------------------------------------------------------------------------------
	//			MIME Types
	//----------------------------------------------------------------------------------------------------
	Options.MIMETypes["hqx"] = "application/mac-binhex40";
	Options.MIMETypes["doc"] = "application/msword";
	Options.MIMETypes["bin"] = "application/octet-stream";
	Options.MIMETypes["dms"] = "application/octet-stream";
	Options.MIMETypes["lha"] = "application/octet-stream";
	Options.MIMETypes["lzh"] = "application/octet-stream";
	Options.MIMETypes["exe"] = "application/octet-stream";
	Options.MIMETypes["class"] = "application/octet-stream";
	Options.MIMETypes["pdf"] = "application/pdf";
	Options.MIMETypes["ai"] = "application/postscript";
	Options.MIMETypes["eps"] = "application/postscript";
	Options.MIMETypes["ps"] = "application/postscript";
	Options.MIMETypes["smi"] = "application/smil";
	Options.MIMETypes["smil"] = "application/smil";
	Options.MIMETypes["mif"] = "application/vnd.mif";
	Options.MIMETypes["asf"] = "application/vnd.ms-asf";
	Options.MIMETypes["xls"] = "application/vnd.ms-excel";
	Options.MIMETypes["ppt"] = "application/vnd.ms-powerpoint";
	Options.MIMETypes["vcd"] = "application/x-cdlink";
	Options.MIMETypes["Z"] = "application/x-compress";
	Options.MIMETypes["cpio"] = "application/x-cpio";
	Options.MIMETypes["csh"] = "application/x-csh";
	Options.MIMETypes["dcr"] = "application/x-director";
	Options.MIMETypes["dir"] = "application/x-director";
	Options.MIMETypes["dxr"] = "application/x-director";
	Options.MIMETypes["dvi"] = "application/x-dvi";
	Options.MIMETypes["gtar"] = "application/x-gtar";
	Options.MIMETypes["gz"] = "application/x-gzip";
	Options.MIMETypes["js"] = "application/x-javascript";
	Options.MIMETypes["latex"] = "application/x-latex";
	Options.MIMETypes["sh"] = "application/x-sh";
	Options.MIMETypes["shar"] = "application/x-shar";
	Options.MIMETypes["swf"] = "application/x-shockwave-flash";
	Options.MIMETypes["sit"] = "application/x-stuffit";
	Options.MIMETypes["tar"] = "application/x-tar";
	Options.MIMETypes["tcl"] = "application/x-tcl";
	Options.MIMETypes["tex"] = "application/x-tex";
	Options.MIMETypes["texinfo"] = "application/x-texinfo";
	Options.MIMETypes["texi"] = "application/x-texinfo";
	Options.MIMETypes["t"] = "application/x-troff";
	Options.MIMETypes["tr"] = "application/x-troff";
	Options.MIMETypes["roff"] = "application/x-troff";
	Options.MIMETypes["man"] = "application/x-troff-man";
	Options.MIMETypes["me"] = "application/x-troff-me";
	Options.MIMETypes["ms"] = "application/x-troff-ms";
	Options.MIMETypes["zip"] = "application/zip";
	Options.MIMETypes["au"] = "audio/basic";
	Options.MIMETypes["snd"] = "audio/basic";
	Options.MIMETypes["mid"] = "audio/midi";
	Options.MIMETypes["midi"] = "audio/midi";
	Options.MIMETypes["kar"] = "audio/midi";
	Options.MIMETypes["mpga"] = "audio/mpeg";
	Options.MIMETypes["mp2"] = "audio/mpeg";
	Options.MIMETypes["mp3"] = "audio/mpeg";
	Options.MIMETypes["aif"] = "audio/x-aiff";
	Options.MIMETypes["aiff"] = "audio/x-aiff";
	Options.MIMETypes["aifc"] = "audio/x-aiff";
	Options.MIMETypes["ram"] = "audio/x-pn-realaudio";
	Options.MIMETypes["rm"] = "audio/x-pn-realaudio";
	Options.MIMETypes["ra"] = "audio/x-realaudio";
	Options.MIMETypes["wav"] = "audio/x-wav";
	Options.MIMETypes["bmp"] = "image/bmp";
	Options.MIMETypes["gif"] = "image/gif";
	Options.MIMETypes["ief"] = "image/ief";
	Options.MIMETypes["jpeg"] = "image/jpeg";
	Options.MIMETypes["jpg"] = "image/jpeg";
	Options.MIMETypes["jpe"] = "image/jpeg";
	Options.MIMETypes["png"] = "image/png";
	Options.MIMETypes["tiff"] = "image/tiff";
	Options.MIMETypes["tif"] = "image/tiff";
	Options.MIMETypes["ras"] = "image/x-cmu-raster";
	Options.MIMETypes["pnm"] = "image/x-portable-anymap";
	Options.MIMETypes["pbm"] = "image/x-portable-bitmap";
	Options.MIMETypes["pgm"] = "image/x-portable-graymap";
	Options.MIMETypes["ppm"] = "image/x-portable-pixmap";
	Options.MIMETypes["rgb"] = "image/x-rgb";
	Options.MIMETypes["xbm"] = "image/x-xbitmap";
	Options.MIMETypes["xpm"] = "image/x-xpixmap";
	Options.MIMETypes["xwd"] = "image/x-xwindowdump";
	Options.MIMETypes["igs"] = "model/iges";
	Options.MIMETypes["iges"] = "model/iges";
	Options.MIMETypes["msh"] = "model/mesh";
	Options.MIMETypes["mesh"] = "model/mesh";
	Options.MIMETypes["silo"] = "model/mesh";
	Options.MIMETypes["wrl"] = "model/vrml";
	Options.MIMETypes["vrml"] = "model/vrml";
	Options.MIMETypes["css"] = "text/css";
	Options.MIMETypes["html"] = "text/html";
	Options.MIMETypes["htm"] = "text/html";
	Options.MIMETypes["asc"] = "text/plain";
	Options.MIMETypes["txt"] = "text/plain";
	Options.MIMETypes["rtx"] = "text/richtext";
	Options.MIMETypes["rtf"] = "text/rtf";
	Options.MIMETypes["sgml"] = "text/sgml";
	Options.MIMETypes["sgm"] = "text/sgml";
	Options.MIMETypes["tsv"] = "text/tab-separated-values";
	Options.MIMETypes["xml"] = "text/xml";
	Options.MIMETypes["mpeg"] = "video/mpeg";
	Options.MIMETypes["mpg"] = "video/mpeg";
	Options.MIMETypes["mpe"] = "video/mpeg";
	Options.MIMETypes["qt"] = "video/quicktime";
	Options.MIMETypes["mov"] = "video/quicktime";
	Options.MIMETypes["avi"] = "video/x-msvideo";

	//----------------------------------------------------------------------------------------------------
	//			Binary Files
	//			The following extensions should be opened as binary. Anything else should be as text
	//----------------------------------------------------------------------------------------------------
	Options.Binary["hqx"] = true;
	Options.Binary["doc"] = true;
	Options.Binary["bin"] = true;
	Options.Binary["dms"] = true;
	Options.Binary["lha"] = true;
	Options.Binary["lzh"] = true;
	Options.Binary["exe"] = true;
	Options.Binary["class"] = true;
	Options.Binary["pdf"] = true;
	Options.Binary["ai"] = true;
	Options.Binary["eps"] = true;
	Options.Binary["ps"] = true;
	Options.Binary["smi"] = true;
	Options.Binary["smil"] = true;
	Options.Binary["mif"] = true;
	Options.Binary["asf"] = true;
	Options.Binary["xls"] = true;
	Options.Binary["ppt"] = true;
	Options.Binary["vcd"] = true;
	Options.Binary["Z"] = true;
	Options.Binary["cpio"] = true;
	Options.Binary["csh"] = true;
	Options.Binary["dcr"] = true;
	Options.Binary["dir"] = true;
	Options.Binary["dxr"] = true;
	Options.Binary["dvi"] = true;
	Options.Binary["gtar"] = true;
	Options.Binary["gz"] = true;
	Options.Binary["js"] = true;
	Options.Binary["latex"] = true;
	Options.Binary["sh"] = true;
	Options.Binary["shar"] = true;
	Options.Binary["swf"] = true;
	Options.Binary["sit"] = true;
	Options.Binary["tar"] = true;
	Options.Binary["tcl"] = true;
	Options.Binary["tex"] = true;
	Options.Binary["texinfo"] = true;
	Options.Binary["texi"] = true;
	Options.Binary["t"] = true;
	Options.Binary["tr"] = true;
	Options.Binary["roff"] = true;
	Options.Binary["man"] = true;
	Options.Binary["me"] = true;
	Options.Binary["ms"] = true;
	Options.Binary["zip"] = true;
	Options.Binary["au"] = true;
	Options.Binary["snd"] = true;
	Options.Binary["mid"] = true;
	Options.Binary["midi"] = true;
	Options.Binary["kar"] = true;
	Options.Binary["mpga"] = true;
	Options.Binary["mp2"] = true;
	Options.Binary["mp3"] = true;
	Options.Binary["aif"] = true;
	Options.Binary["aiff"] = true;
	Options.Binary["aifc"] = true;
	Options.Binary["ram"] = true;
	Options.Binary["rm"] = true;
	Options.Binary["ra"] = true;
	Options.Binary["wav"] = true;
	Options.Binary["bmp"] = true;
	Options.Binary["gif"] = true;
	Options.Binary["ief"] = true;
	Options.Binary["jpeg"] = true;
	Options.Binary["jpg"] = true;
	Options.Binary["jpe"] = true;
	Options.Binary["png"] = true;
	Options.Binary["tiff"] = true;
	Options.Binary["tif"] = true;
	Options.Binary["ras"] = true;
	Options.Binary["pnm"] = true;
	Options.Binary["pbm"] = true;
	Options.Binary["pgm"] = true;
	Options.Binary["ppm"] = true;
	Options.Binary["rgb"] = true;
	Options.Binary["xbm"] = true;
	Options.Binary["xpm"] = true;
	Options.Binary["xwd"] = true;
	Options.Binary["igs"] = true;
	Options.Binary["iges"] = true;
	Options.Binary["msh"] = true;
	Options.Binary["mesh"] = true;
	Options.Binary["silo"] = true;
	Options.Binary["wrl"] = true;
	Options.Binary["vrml"] = true;
	Options.Binary["mpeg"] = true;
	Options.Binary["mpg"] = true;
	Options.Binary["mpe"] = true;
	Options.Binary["qt"] = true;
	Options.Binary["mov"] = true;
	Options.Binary["avi"] = true;
	
	//-----------------------------------------------------------------------------------------
	// Map status code numbers to text codes
	//-----------------------------------------------------------------------------------------
	Options.ErrorCode[200] = "OK";
	Options.ErrorCode[404] = "File Not Found";
	Options.ErrorCode[301] = "Moved Permanently";
	Options.ErrorCode[302] = "Moved Temporarily";
	Options.ErrorCode[500] = "Internal Server Error";

	//-----------------------------------------------------------------------------------------
	// Step 3: Start web server
	//-----------------------------------------------------------------------------------------
	int SFD_Listen;																	// Socket Descriptor we listen on
	struct sockaddr_in ServerAddress;												// Servers address structure
	int Result;																		// Result flag that will be used through the program for errors
	
    // Set socket
	SFD_Listen = socket(AF_INET, SOCK_STREAM, 0);									// Find a good socket
	if (SFD_Listen == -1)															// Socket could not be made
	{
		return SWEBS_RETURN_COULDNOTLISTEN;
	}

	// Assign server information
	ServerAddress.sin_family = AF_INET;												// Using TCP/IP
	ServerAddress.sin_port = htons(Options.Port);									// Port
	if (Options.IPAddress.length() > 1)
    {
        ServerAddress.sin_addr.s_addr = inet_addr(Options.IPAddress.c_str());       // Use the address specified
    }
    else 
    {
        ServerAddress.sin_addr.s_addr = INADDR_ANY;								    // Use any and all addresses
    }
    memset(&(ServerAddress.sin_zero), '\0', 8);										// Zero out rest

	// Bind to port
	Result = bind(SFD_Listen, (struct sockaddr *) &ServerAddress, sizeof(struct sockaddr));
	if (Result == -1)
	{
        ReturnCode = SWEBS_RETURN_COULDNOTBIND;
		return ReturnCode;
	}

	// Listen
	Result = listen(SFD_Listen, Options.MaxConnections);
	if (Result == -1)
	{
        ReturnCode = SWEBS_RETURN_COULDNOTLISTEN;
		return ReturnCode;
	}
	
	int Size = sizeof(struct sockaddr_in);

    //-----------------------------------------------------------------------------------------
    // Step 4.5: Create Stats handling function
    //-----------------------------------------------------------------------------------------
    DWORD dwThreadId2;															    // Info for the thead 
	HANDLE hThread2; 

	// CreateThread and process the request
	hThread2 = CreateThread( 
        NULL,																		// default security attributes 
        0,                           												// use default stack size  
        HandleStatsFile,                 											// thread function 
        NULL,                													    // argument to thread function 
        0,                           												// use default creation flags 
        &dwThreadId2);                												// returns the thread identifier 
		
	if (hThread2 != NULL)														    // If the thread was created, destroy it
	{
		CloseHandle( hThread2 );
	}
	//-----------------------------------------------------------------------------------------
	// Step 5: Handle Requests
	//-----------------------------------------------------------------------------------------
	SERVER_STOP = false;
    
    ReturnCode = SWEBS_RETURN_COULDNOTACCEPT;
    
    // This will create a message queue for the thread...
    PeekMessage(&Message, NULL, WM_USER, WM_USER, PM_NOREMOVE);

    // Create a new thread to do the listening
    HANDLE hAcceptThread;
    DWORD dwAcceptThreadID;

    hAcceptThread = CreateThread(NULL, 0, AcceptThreadProc, (LPVOID)SFD_Listen, 0, &dwAcceptThreadID);

    // Check for a WM_QUIT in the message queue... 
    while(!PeekMessage(&Message, NULL, WM_QUIT, WM_QUIT, PM_REMOVE))
    {
        // Loop until we get a quit messages 
        Sleep(300);
	}
    TerminateThread(hAcceptThread, 0);
    ReturnCode = SWEBS_RETURN_SUCCESS;                                              // We know the server was successful

	closesocket(SFD_Listen);
	UnregisterClass("SWEBSWindowClass", hinstance);									// Free Memory
    WSACleanup();
    return ReturnCode;
}

//---------------------------------------------------------------------------------------------
//			Request Processor - used by CreateThread()
//---------------------------------------------------------------------------------------------
DWORD WINAPI ProcessRequest(LPVOID lpParam )
{
	ARGUMENT * Arg = (ARGUMENT *)lpParam;											// Split the paramater into the arguments
	
	CONNECTION * New = new CONNECTION(Arg->SFD, Arg->CLA);
	if (New)
	{
		New->ReadRequest();															// Read in the request
		New->HandleRequest();														// Handle the request

		delete New;
	}
	closesocket(Arg->SFD);
	return 0;
}

//---------------------------------------------------------------------------------------------
//			Control Handler
//---------------------------------------------------------------------------------------------
LRESULT CALLBACK WinProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) 
{ 
	switch(message) 
	{ 
        case WM_DESTROY:
        case WM_QUIT:
        case WM_CLOSE:
		    SERVER_STOP = true;
            PostQuitMessage(0);
            break;
	    default:
            break;
	}
    // Report current status
    return DefWindowProc(hwnd, message, wparam, lparam);
}

//---------------------------------------------------------------------------------------------
//			TestLog
//---------------------------------------------------------------------------------------------
void TestLog(string Data)
{
	FILE* log;
	log = fopen(Options.Logfile.c_str(), "a+");
	if (log == NULL)
      return ;
	fprintf(log, "%s", Data.c_str());
	fclose(log);
}


//---------------------------------------------------------------------------------------------
//          AcceptThreadProc() - Thread proc that handles listening for connections
//---------------------------------------------------------------------------------------------
DWORD WINAPI AcceptThreadProc(LPVOID lpParam)
{    
    int SFD_Listen = (int) lpParam;
    int SFD_New;
    struct sockaddr_in ClientAddress ;
    int Size = sizeof(struct sockaddr);
    while (1)
    {
        DWORD dwThreadId;														// Info for the thead 
	    HANDLE hThread; 
            
        SFD_New = accept(SFD_Listen, (struct sockaddr *) &ClientAddress, &Size);
	    // Create a structure of type ARGUMENT to be passed to the new thread
	    ARGUMENT Argument;
	    Argument.CLA = ClientAddress;
        Argument.SFD = SFD_New;

		    // CreateThread and process the request
	        hThread = CreateThread( 
                NULL,																// default security attributes 
                0,                           										    // use default stack size  
                ProcessRequest,                 									// thread function 
                &Argument,                											// argument to thread function 
                0,                           										// use default creation flags 
                &dwThreadId
                );                												            // returns the thread identifier 
		
	        if (hThread != NULL)												        // If the thread was created, destroy it
	        {			    
                CloseHandle( hThread );
	        }
    }
}

//---------------------------------------------------------------------------------------------
