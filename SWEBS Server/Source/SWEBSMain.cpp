//----------------------------------------------------------------------------------
//  SWEBSEngine
//  ----------------------
//  This is the main SWEBS processing engine, as called by the SWEBS Web Server service. This
//  application is the main workhorse of the SWEBS Web Server.
//
//  Copyright SWEBS Development team 2003
//----------------------------------------------------------------------------------
//			Includes
//----------------------------------------------------------------------------------
#pragma warning(disable:4786)
#pragma warning(disable:4089)
#include <windows.h>
#include <winsock.h>
#include <memory>
#include <string>
#include "../Include/SWEBSConnection.hpp"
#include "../Include/resource.h"
#include "../Include/SWEBSGlobals.hpp"
#include "../Include/SWEBSUtilities.hpp"

using namespace std;

#pragma comment(lib, "wsock32.lib")

// The idea of using LIB's was aborted, its easier to include all the source files in the final project
//#pragma comment(lib, "../SWEBSConnection/Release/SWEBSConnection.lib")
//#pragma comment(lib, "../SWEBSSockets/Release/SWEBSSockets.lib")
//#pragma comment(lib, "../SWEBSGlobals/Release/SWEBSGlobals.lib")
//#pragma comment(lib, "../SWEBSUtilities/Release/SWEBSUtilities.lib")
// The first two of these files come from the Chilkat XML library, freely downloadable from
//  www.xml-parser.com
#pragma comment(lib, "ChilkatRelSt.lib")
#pragma comment(lib, "CkBaseRelSt.lib")
#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "rpcrt4.lib")
#pragma comment(lib, "crypt32.lib")

//----------------------------------------------------------------------------------
//			Function Declarations
//---------------------------------------------------------------------------------- 
void TestLog(string);
DWORD WINAPI ProcessRequest(LPVOID lpParam );                                       // Thread created to process a request
DWORD WINAPI StartThread(LPVOID lpParam);                                           // Thread function called when SWEBSStart() is called by the UI

//----------------------------------------------------------------------------------
//			Globals
//----------------------------------------------------------------------------------
bool SERVER_STOP = false;                                                           // Flag for if the server is running

const int SWEBS_RETURN_UNKNOWN          = 0x00;                                     // Unknown error occured
const int SWEBS_RETURN_SUCCESS          = 0x01;                                     // Server ran fine
const int SWEBS_RETURN_COULDNOTBIND     = 0x02;                                     // Could not bind() to port
const int SWEBS_RETURN_CONFIGNOTLOADED  = 0x03;                                     // Could not load config file
const int SWEBS_RETURN_COULDNOTLISTEN   = 0x04;                                     // Could not listen()
const int SWEBS_RETURN_COULDNOTACCEPT   = 0x05;                                     // Could not accept()

int ReturnCode = SWEBS_RETURN_UNKNOWN;                                              // Number for main() to return, can be set from any function

struct ARGUMENT
{
	int SFD;
	struct sockaddr_in CLA;
};

//----------------------------------------------------------------------------------
//          DllMain
//----------------------------------------------------------------------------------
int WINAPI DllMain(HINSTANCE hInstance, DWORD fdwReason, PVOID pvReserved)
{
    return true;
}

//----------------------------------------------------------------------------------
//			SWEBSStart
//----------------------------------------------------------------------------------
int SWEBSStart()
{
    DWORD dwThreadId;															    // Info for the thead 
	HANDLE hThread; 

	// CreateThread to start the server
	hThread = CreateThread( 
        NULL,																		// Default security attributes 
        0,                           												// Use default stack size  
        StartThread,                 											    // Thread function 
        NULL,                													    // Argument to thread function 
        0,                           												// Use default creation flags 
        &dwThreadId);                												// Returns the thread identifier 
		
	if (hThread != NULL)														    // If the thread was created, destroy it
	{
		CloseHandle( hThread );
	}
    return ReturnCode;
}

//----------------------------------------------------------------------------------
//			StartThread
//----------------------------------------------------------------------------------
DWORD WINAPI StartThread(LPVOID lpParam)
{
	//------------------------------------------------------------------------------
	// Step 1: Do stuff we must do as a windows application
	//------------------------------------------------------------------------------
	WSADATA wsaData;
    WSAStartup(MAKEWORD(1,1), &wsaData);
    
    //------------------------------------------------------------------------------
	// Step 2: Set up SWEBSGlobals
	//------------------------------------------------------------------------------
	// These are default settings, incase the configuration file is corrupt
	SWEBSGlobals.Timeout = 20;
	SWEBSGlobals.ErrorLog = "C:\\SWS\\Errorlog.log";

	// Read the real settings from the config file
	bool ReadConfig = SWEBSGlobals.ReadSettings();

	if (ReadConfig == false)
	{
		// The configuration file had errors.
		SWEBSGlobals.LogError("Warning: Could not load configuration file properly");
        ReturnCode = SWEBS_RETURN_CONFIGNOTLOADED;
		return ReturnCode;
	}

    // Set non request-specific variables
    SWEBSGlobals.CGIVariables.GATEWAY_INTERFACE = "CGI/1.0";
    SWEBSGlobals.CGIVariables.SERVER_NAME = SWEBSGlobals.Servername;
    SWEBSGlobals.CGIVariables.SERVER_SOFTWARE = SWEBSGlobals.Servername;

	//------------------------------------------------------------------------------
	//			MIME Types
	//------------------------------------------------------------------------------
    SWEBSGlobals.MIMETypes["hqx"] = "application/mac-binhex40";
	SWEBSGlobals.MIMETypes["doc"] = "application/msword";
	SWEBSGlobals.MIMETypes["bin"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["dms"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["lha"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["lzh"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["exe"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["class"] = "application/octet-stream";
	SWEBSGlobals.MIMETypes["pdf"] = "application/pdf";
	SWEBSGlobals.MIMETypes["ai"] = "application/postscript";
	SWEBSGlobals.MIMETypes["eps"] = "application/postscript";
	SWEBSGlobals.MIMETypes["ps"] = "application/postscript";
	SWEBSGlobals.MIMETypes["smi"] = "application/smil";
	SWEBSGlobals.MIMETypes["smil"] = "application/smil";
	SWEBSGlobals.MIMETypes["mif"] = "application/vnd.mif";
	SWEBSGlobals.MIMETypes["asf"] = "application/vnd.ms-asf";
	SWEBSGlobals.MIMETypes["xls"] = "application/vnd.ms-excel";
	SWEBSGlobals.MIMETypes["ppt"] = "application/vnd.ms-powerpoint";
	SWEBSGlobals.MIMETypes["vcd"] = "application/x-cdlink";
	SWEBSGlobals.MIMETypes["Z"] = "application/x-compress";
	SWEBSGlobals.MIMETypes["cpio"] = "application/x-cpio";
	SWEBSGlobals.MIMETypes["csh"] = "application/x-csh";
	SWEBSGlobals.MIMETypes["dcr"] = "application/x-director";
	SWEBSGlobals.MIMETypes["dir"] = "application/x-director";
	SWEBSGlobals.MIMETypes["dxr"] = "application/x-director";
	SWEBSGlobals.MIMETypes["dvi"] = "application/x-dvi";
	SWEBSGlobals.MIMETypes["gtar"] = "application/x-gtar";
	SWEBSGlobals.MIMETypes["gz"] = "application/x-gzip";
	SWEBSGlobals.MIMETypes["js"] = "application/x-javascript";
	SWEBSGlobals.MIMETypes["latex"] = "application/x-latex";
	SWEBSGlobals.MIMETypes["sh"] = "application/x-sh";
	SWEBSGlobals.MIMETypes["shar"] = "application/x-shar";
	SWEBSGlobals.MIMETypes["swf"] = "application/x-shockwave-flash";
	SWEBSGlobals.MIMETypes["sit"] = "application/x-stuffit";
	SWEBSGlobals.MIMETypes["tar"] = "application/x-tar";
	SWEBSGlobals.MIMETypes["tcl"] = "application/x-tcl";
	SWEBSGlobals.MIMETypes["tex"] = "application/x-tex";
	SWEBSGlobals.MIMETypes["texinfo"] = "application/x-texinfo";
	SWEBSGlobals.MIMETypes["texi"] = "application/x-texinfo";
	SWEBSGlobals.MIMETypes["t"] = "application/x-troff";
	SWEBSGlobals.MIMETypes["tr"] = "application/x-troff";
	SWEBSGlobals.MIMETypes["roff"] = "application/x-troff";
	SWEBSGlobals.MIMETypes["man"] = "application/x-troff-man";
	SWEBSGlobals.MIMETypes["me"] = "application/x-troff-me";
	SWEBSGlobals.MIMETypes["ms"] = "application/x-troff-ms";
	SWEBSGlobals.MIMETypes["zip"] = "application/zip";
	SWEBSGlobals.MIMETypes["au"] = "audio/basic";
	SWEBSGlobals.MIMETypes["snd"] = "audio/basic";
	SWEBSGlobals.MIMETypes["mid"] = "audio/midi";
	SWEBSGlobals.MIMETypes["midi"] = "audio/midi";
	SWEBSGlobals.MIMETypes["kar"] = "audio/midi";
	SWEBSGlobals.MIMETypes["mpga"] = "audio/mpeg";
	SWEBSGlobals.MIMETypes["mp2"] = "audio/mpeg";
	SWEBSGlobals.MIMETypes["mp3"] = "audio/mpeg";
	SWEBSGlobals.MIMETypes["aif"] = "audio/x-aiff";
	SWEBSGlobals.MIMETypes["aiff"] = "audio/x-aiff";
	SWEBSGlobals.MIMETypes["aifc"] = "audio/x-aiff";
	SWEBSGlobals.MIMETypes["ram"] = "audio/x-pn-realaudio";
	SWEBSGlobals.MIMETypes["rm"] = "audio/x-pn-realaudio";
	SWEBSGlobals.MIMETypes["ra"] = "audio/x-realaudio";
	SWEBSGlobals.MIMETypes["wav"] = "audio/x-wav";
	SWEBSGlobals.MIMETypes["bmp"] = "image/bmp";
	SWEBSGlobals.MIMETypes["gif"] = "image/gif";
	SWEBSGlobals.MIMETypes["ief"] = "image/ief";
	SWEBSGlobals.MIMETypes["jpeg"] = "image/jpeg";
	SWEBSGlobals.MIMETypes["jpg"] = "image/jpeg";
	SWEBSGlobals.MIMETypes["jpe"] = "image/jpeg";
	SWEBSGlobals.MIMETypes["png"] = "image/png";
	SWEBSGlobals.MIMETypes["tiff"] = "image/tiff";
	SWEBSGlobals.MIMETypes["tif"] = "image/tiff";
	SWEBSGlobals.MIMETypes["ras"] = "image/x-cmu-raster";
	SWEBSGlobals.MIMETypes["pnm"] = "image/x-portable-anymap";
	SWEBSGlobals.MIMETypes["pbm"] = "image/x-portable-bitmap";
	SWEBSGlobals.MIMETypes["pgm"] = "image/x-portable-graymap";
	SWEBSGlobals.MIMETypes["ppm"] = "image/x-portable-pixmap";
	SWEBSGlobals.MIMETypes["rgb"] = "image/x-rgb";
	SWEBSGlobals.MIMETypes["xbm"] = "image/x-xbitmap";
	SWEBSGlobals.MIMETypes["xpm"] = "image/x-xpixmap";
	SWEBSGlobals.MIMETypes["xwd"] = "image/x-xwindowdump";
	SWEBSGlobals.MIMETypes["igs"] = "model/iges";
	SWEBSGlobals.MIMETypes["iges"] = "model/iges";
	SWEBSGlobals.MIMETypes["msh"] = "model/mesh";
	SWEBSGlobals.MIMETypes["mesh"] = "model/mesh";
	SWEBSGlobals.MIMETypes["silo"] = "model/mesh";
	SWEBSGlobals.MIMETypes["wrl"] = "model/vrml";
	SWEBSGlobals.MIMETypes["vrml"] = "model/vrml";
	SWEBSGlobals.MIMETypes["css"] = "text/css";
	SWEBSGlobals.MIMETypes["html"] = "text/html";
	SWEBSGlobals.MIMETypes["htm"] = "text/html";
	SWEBSGlobals.MIMETypes["asc"] = "text/plain";
	SWEBSGlobals.MIMETypes["txt"] = "text/plain";
	SWEBSGlobals.MIMETypes["rtx"] = "text/richtext";
	SWEBSGlobals.MIMETypes["rtf"] = "text/rtf";
	SWEBSGlobals.MIMETypes["sgml"] = "text/sgml";
	SWEBSGlobals.MIMETypes["sgm"] = "text/sgml";
	SWEBSGlobals.MIMETypes["tsv"] = "text/tab-separated-values";
	SWEBSGlobals.MIMETypes["xml"] = "text/xml";
	SWEBSGlobals.MIMETypes["mpeg"] = "video/mpeg";
	SWEBSGlobals.MIMETypes["mpg"] = "video/mpeg";
	SWEBSGlobals.MIMETypes["mpe"] = "video/mpeg";
	SWEBSGlobals.MIMETypes["qt"] = "video/quicktime";
	SWEBSGlobals.MIMETypes["mov"] = "video/quicktime";
	SWEBSGlobals.MIMETypes["avi"] = "video/x-msvideo";

	//------------------------------------------------------------------------------
	//			Binary Files
	//			The following extensions should be opened as binary. Anything else should be as text
	//------------------------------------------------------------------------------
	SWEBSGlobals.Binary["hqx"] = true;
	SWEBSGlobals.Binary["doc"] = true;
	SWEBSGlobals.Binary["bin"] = true;
	SWEBSGlobals.Binary["dms"] = true;
	SWEBSGlobals.Binary["lha"] = true;
	SWEBSGlobals.Binary["lzh"] = true;
	SWEBSGlobals.Binary["exe"] = true;
	SWEBSGlobals.Binary["class"] = true;
	SWEBSGlobals.Binary["pdf"] = true;
	SWEBSGlobals.Binary["ai"] = true;
	SWEBSGlobals.Binary["eps"] = true;
	SWEBSGlobals.Binary["ps"] = true;
	SWEBSGlobals.Binary["smi"] = true;
	SWEBSGlobals.Binary["smil"] = true;
	SWEBSGlobals.Binary["mif"] = true;
	SWEBSGlobals.Binary["asf"] = true;
	SWEBSGlobals.Binary["xls"] = true;
	SWEBSGlobals.Binary["ppt"] = true;
	SWEBSGlobals.Binary["vcd"] = true;
	SWEBSGlobals.Binary["Z"] = true;
	SWEBSGlobals.Binary["cpio"] = true;
	SWEBSGlobals.Binary["csh"] = true;
	SWEBSGlobals.Binary["dcr"] = true;
	SWEBSGlobals.Binary["dir"] = true;
	SWEBSGlobals.Binary["dxr"] = true;
	SWEBSGlobals.Binary["dvi"] = true;
	SWEBSGlobals.Binary["gtar"] = true;
	SWEBSGlobals.Binary["gz"] = true;
	SWEBSGlobals.Binary["js"] = true;
	SWEBSGlobals.Binary["latex"] = true;
	SWEBSGlobals.Binary["sh"] = true;
	SWEBSGlobals.Binary["shar"] = true;
	SWEBSGlobals.Binary["swf"] = true;
	SWEBSGlobals.Binary["sit"] = true;
	SWEBSGlobals.Binary["tar"] = true;
	SWEBSGlobals.Binary["tcl"] = true;
	SWEBSGlobals.Binary["tex"] = true;
	SWEBSGlobals.Binary["texinfo"] = true;
	SWEBSGlobals.Binary["texi"] = true;
	SWEBSGlobals.Binary["t"] = true;
	SWEBSGlobals.Binary["tr"] = true;
	SWEBSGlobals.Binary["roff"] = true;
	SWEBSGlobals.Binary["man"] = true;
	SWEBSGlobals.Binary["me"] = true;
	SWEBSGlobals.Binary["ms"] = true;
	SWEBSGlobals.Binary["zip"] = true;
	SWEBSGlobals.Binary["au"] = true;
	SWEBSGlobals.Binary["snd"] = true;
	SWEBSGlobals.Binary["mid"] = true;
	SWEBSGlobals.Binary["midi"] = true;
	SWEBSGlobals.Binary["kar"] = true;
	SWEBSGlobals.Binary["mpga"] = true;
	SWEBSGlobals.Binary["mp2"] = true;
	SWEBSGlobals.Binary["mp3"] = true;
	SWEBSGlobals.Binary["aif"] = true;
	SWEBSGlobals.Binary["aiff"] = true;
	SWEBSGlobals.Binary["aifc"] = true;
	SWEBSGlobals.Binary["ram"] = true;
	SWEBSGlobals.Binary["rm"] = true;
	SWEBSGlobals.Binary["ra"] = true;
	SWEBSGlobals.Binary["wav"] = true;
	SWEBSGlobals.Binary["bmp"] = true;
	SWEBSGlobals.Binary["gif"] = true;
	SWEBSGlobals.Binary["ief"] = true;
	SWEBSGlobals.Binary["jpeg"] = true;
	SWEBSGlobals.Binary["jpg"] = true;
	SWEBSGlobals.Binary["jpe"] = true;
	SWEBSGlobals.Binary["png"] = true;
	SWEBSGlobals.Binary["tiff"] = true;
	SWEBSGlobals.Binary["tif"] = true;
	SWEBSGlobals.Binary["ras"] = true;
	SWEBSGlobals.Binary["pnm"] = true;
	SWEBSGlobals.Binary["pbm"] = true;
	SWEBSGlobals.Binary["pgm"] = true;
	SWEBSGlobals.Binary["ppm"] = true;
	SWEBSGlobals.Binary["rgb"] = true;
	SWEBSGlobals.Binary["xbm"] = true;
	SWEBSGlobals.Binary["xpm"] = true;
	SWEBSGlobals.Binary["xwd"] = true;
	SWEBSGlobals.Binary["igs"] = true;
	SWEBSGlobals.Binary["iges"] = true;
	SWEBSGlobals.Binary["msh"] = true;
	SWEBSGlobals.Binary["mesh"] = true;
	SWEBSGlobals.Binary["silo"] = true;
	SWEBSGlobals.Binary["wrl"] = true;
	SWEBSGlobals.Binary["vrml"] = true;
	SWEBSGlobals.Binary["mpeg"] = true;
	SWEBSGlobals.Binary["mpg"] = true;
	SWEBSGlobals.Binary["mpe"] = true;
	SWEBSGlobals.Binary["qt"] = true;
	SWEBSGlobals.Binary["mov"] = true;
	SWEBSGlobals.Binary["avi"] = true;
	
	//------------------------------------------------------------------------------
	// Map status code numbers to text codes
	//------------------------------------------------------------------------------
	SWEBSGlobals.ErrorCodes[200] = "OK";
	SWEBSGlobals.ErrorCodes[404] = "File Not Found";
	SWEBSGlobals.ErrorCodes[301] = "Moved Permanently";
	SWEBSGlobals.ErrorCodes[302] = "Moved Temporarily";
	SWEBSGlobals.ErrorCodes[500] = "Internal Server Error";

	//!=============================================================================
    // This is for ISAPI testing, remove after
    SWEBSGlobals.IsapiDLLExtensions["php"] = "C:\\PHP\\php-4.3.1-Win32\\SAPI\\php4isapi.dll";

    SWEBSGlobals.IsapiDLL["php"]= LoadLibrary(SWEBSGlobals.IsapiDLLExtensions["php"].c_str());// Loads the DLL

    //------------------------------------------------------------------------------
	// Step 3: Start web server
	//------------------------------------------------------------------------------
	int Result;																		// Result flag that will be used through the program for errors
	
    // Set socket
	SWEBSGlobals.SFD_Listen = socket(AF_INET, SOCK_STREAM, 0);									// Find a good socket
	if (SWEBSGlobals.SFD_Listen == -1)															// Socket could not be made
	{
		return SWEBS_RETURN_COULDNOTLISTEN;
	}

	// Assign server information
	SWEBSGlobals.ServerAddress.sin_family = AF_INET;												// Using TCP/IP
	SWEBSGlobals.ServerAddress.sin_port = htons(SWEBSGlobals.Port);								// Port
	if (SWEBSGlobals.IPAddress.length() > 1)
    {
        SWEBSGlobals.ServerAddress.sin_addr.s_addr = inet_addr(SWEBSGlobals.IPAddress.c_str());  // Use the address specified
    }
    else 
    {
        SWEBSGlobals.ServerAddress.sin_addr.s_addr = INADDR_ANY;								    // Use any and all addresses
    }
    memset(&(SWEBSGlobals.ServerAddress.sin_zero), '\0', 8);										// Zero out rest

	// Bind to port
	Result = bind(SWEBSGlobals.SFD_Listen, (struct sockaddr *) &(SWEBSGlobals.ServerAddress), sizeof(struct sockaddr));
	if (Result == -1)
	{
        ReturnCode = SWEBS_RETURN_COULDNOTBIND;
		return ReturnCode;
	}

	// Listen
	Result = listen(SWEBSGlobals.SFD_Listen, SWEBSGlobals.MaxConnections);
	if (Result == -1)
	{
        ReturnCode = SWEBS_RETURN_COULDNOTLISTEN;
		return ReturnCode;
	}

    //------------------------------------------------------------------------------
    // Step 4.5: Create Stats handling function
    //------------------------------------------------------------------------------
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
	//------------------------------------------------------------------------------
	// Step 5: Handle Requests
	//------------------------------------------------------------------------------
	SERVER_STOP = false;
    
    ReturnCode = SWEBS_RETURN_COULDNOTACCEPT;
    
    int SFD_New = 0;
    struct sockaddr_in ClientAddress;
    int Size = sizeof(struct sockaddr);
    DWORD dwThreadId;
    HANDLE hThread;
    
    while (SERVER_STOP != true)
    {
        SFD_New = accept(SWEBSGlobals.SFD_Listen, (struct sockaddr *) &ClientAddress, &Size);
	    // Create a structure of type ARGUMENT to be passed to the new thread
	    ARGUMENT Argument;
	    Argument.CLA = ClientAddress;
        Argument.SFD = SFD_New;

		// CreateThread and process the request
	    hThread = CreateThread( 
            NULL,																    // default security attributes 
            0,                           										    // use default stack size  
            ProcessRequest,                 									    // thread function 
            &Argument,                											    // argument to thread function 
            0,                           										    // use default creation flags 
            &dwThreadId
            );                												        // returns the thread identifier 
		
        if (hThread != NULL)												        // If the thread was created, destroy it
	    {			    
            CloseHandle( hThread );
	    }      
    }
    
    closesocket(SWEBSGlobals.SFD_Listen);
	WSACleanup();
    return ReturnCode;
}

//----------------------------------------------------------------------------------
//			Request Processor - used by CreateThread()
//----------------------------------------------------------------------------------
DWORD WINAPI ProcessRequest(LPVOID lpParam )
{
	ARGUMENT * Arg = (ARGUMENT *)lpParam;											// Split the paramater into the arguments
	
    CONNECTION NewConn(Arg->SFD, Arg->CLA);
    //if (NewConn)
	{
		NewConn.ReadRequest();														// Read in the request
		NewConn.HandleRequest();													// Handle the request
        //delete NewConn;
    }
	closesocket(Arg->SFD);
	return 0;
}

//----------------------------------------------------------------------------------
//          SWEBSReloadOptions
//----------------------------------------------------------------------------------
int SWEBSStop()
{
    SERVER_STOP = true;
    closesocket(SWEBSGlobals.SFD_Listen);
    WSACleanup();
    return 1;
}

//----------------------------------------------------------------------------------
//          SWEBSStats
//----------------------------------------------------------------------------------
const char * SWEBSStats(const char * Stat)
{
    istringstream IS(Stat);
    string TempString;
    IS >> TempString;
    
    if (TempString == "BytesSent")
        return ( IntToString(SWEBSGlobals.BytesSent).c_str() );
    else return "Not yet implemented!";
}

//----------------------------------------------------------------------------------
//			TestLog
//----------------------------------------------------------------------------------
void TestLog(string Data)
{
	FILE* log;
	log = fopen(SWEBSGlobals.Logfile.c_str(), "a+");
	if (log == NULL)
      return ;
	fprintf(log, "%s", Data.c_str());
	fclose(log);
}


//----------------------------------------------------------------------------------
