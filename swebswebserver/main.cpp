//---------------------------------------------------------------------------------------------
//			Includes
//---------------------------------------------------------------------------------------------
#pragma warning(disable:4786)
#pragma warning(disable:4089)
#include <windows.h>
#include <winsock.h>
#include <string>
#include <iostream>
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
void ServiceMain();
void ControlHandler(DWORD request); 
void TestLog(string);
void PrintAccepts(const map<string, bool>::value_type& p);
bool InstallService();
bool UninstallService();
DWORD WINAPI ProcessRequest(LPVOID lpParam );

//---------------------------------------------------------------------------------------------
//			Globals
//---------------------------------------------------------------------------------------------
bool SERVER_STOP = false;

SERVICE_STATUS          ServiceStatus; 
SERVICE_STATUS_HANDLE   hStatus; 

int ReturnCode;                                                                     // Number for main() to return, can be set from any function

const int SWEBS_RETURN_UNKNOWN          = 0x00;                                     // Unknown error occured
const int SWEBS_RETURN_SUCCESS          = 0x01;                                     // Server ran fine
const int SWEBS_RETURN_COULDNOTBIND     = 0x02;                                     // Could not bind() to port
const int SWEBS_RETURN_CONFIGNOTLOADED  = 0x03;                                     // Could not load config file
const int SWEBS_RETURN_COULDNOTLISTEN   = 0x04;                                     // Could not listen()
const int SWEBS_RETURN_COULDNOTACCEPT   = 0x05;                                     // Could not accept()

struct ARGUMENT
{
	int SFD;
	struct sockaddr_in CLA;
};

//---------------------------------------------------------------------------------------------
//			Main
//---------------------------------------------------------------------------------------------
int main(int argc, char ** argv)
{
    // Check for command line arguments
    if (argc > 1)
    {
        if (!strcmpi(argv[1], "/i"))
        {
            // We were told to install the service
            if (!InstallService())                                                  // Try to install
            {
                cout << "Could not install service.\n";
                return false;
            }
        }
        else if (!strcmpi(argv[1], "/u")) 
        {
            // We were told to uninstall
            if (!UninstallService())                                                // Try to uninstall
            {
                cout << "Could not uninstall service.\n";
                return false;
            }
        }
    }

    else 
    {
        ReturnCode = SWEBS_RETURN_UNKNOWN;
	    WSADATA wsaData;
    	WSAStartup(MAKEWORD(1,1), &wsaData);										// Do WSA Stuff

	    SERVICE_TABLE_ENTRY ServiceTable[2]; 
	    ServiceTable[0].lpServiceName = "SWEBS Web Server";							// Name of service
	    ServiceTable[0].lpServiceProc = (LPSERVICE_MAIN_FUNCTION)ServiceMain;		// Main function of service

	    ServiceTable[1].lpServiceName = NULL;										// Must create a null table
	    ServiceTable[1].lpServiceProc = NULL;
	    // Start the control dispatcher thread for our service
	    StartServiceCtrlDispatcher(ServiceTable);									// Jumps to the serice function  

	    WSACleanup();															    // End WSA Stuff
	    return ReturnCode;															// End program
    }
    return ReturnCode;
}

//---------------------------------------------------------------------------------------------
//			Service Main
//---------------------------------------------------------------------------------------------
void ServiceMain()
{
	//-----------------------------------------------------------------------------------------
	// Step 1: Do stuff we must do as a service
	//-----------------------------------------------------------------------------------------
	ServiceStatus.dwServiceType = SERVICE_WIN32;									// Win32 service
	ServiceStatus.dwCurrentState = SERVICE_START_PENDING;
	// Fields the service accepts from the SCM
	ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN;					
	ServiceStatus.dwWin32ExitCode = 0; 
	ServiceStatus.dwServiceSpecificExitCode = 0; 
	ServiceStatus.dwCheckPoint = 0; 
	ServiceStatus.dwWaitHint = 0; 

	hStatus = RegisterServiceCtrlHandler("SWEBS Web Server", (LPHANDLER_FUNCTION)ControlHandler); 
	if (hStatus == (SERVICE_STATUS_HANDLE)0) 
	{ 
      // Registering Control Handler failed
      return; 
	}  

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
		ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
        SetServiceStatus (hStatus, &ServiceStatus);
        ReturnCode = SWEBS_RETURN_CONFIGNOTLOADED;
		return;
	}

    // Set non request-specific variables
    Options.CGIVariables.GATEWAY_INTERFACE = "CGI/1.1";
    Options.CGIVariables.SERVER_NAME = Options.Servername;
    Options.CGIVariables.SERVER_SOFTWARE = Options.Servername;

  	// Report that the service is running
	ServiceStatus.dwCurrentState = SERVICE_RUNNING; 
	SetServiceStatus (hStatus, &ServiceStatus);

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
	int SFD_New;																	// Socket Descriptor for new connections
	struct sockaddr_in ServerAddress;												// Servers address structure
	struct sockaddr_in ClientAddress;												// Clients address structure
	int Result;																		// Result flag that will be used through the program for errors
	
	// Set socket
	SFD_Listen = socket(AF_INET, SOCK_STREAM, 0);									// Find a good socket
	if (SFD_Listen == -1)															// Socket could not be made
	{
		ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
		SetServiceStatus (hStatus, &ServiceStatus);
		return;
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
		ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
		SetServiceStatus (hStatus, &ServiceStatus);
        ReturnCode = SWEBS_RETURN_COULDNOTBIND;
		return;
	}

	// Listen
	Result = listen(SFD_Listen, Options.MaxConnections);
	if (Result == -1)
	{
		ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
		SetServiceStatus (hStatus, &ServiceStatus);
        ReturnCode = SWEBS_RETURN_COULDNOTLISTEN;
		return;
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
	while (!SERVER_STOP)
	{
		SFD_New = accept(SFD_Listen, (struct sockaddr *) &ClientAddress, &Size);
		
		DWORD dwThreadId;															// Info for the thead 
		HANDLE hThread; 

		// Create a structure of type ARGUMENT to be passed to the new thread
		ARGUMENT Argument;
		Argument.CLA = ClientAddress;
		Argument.SFD = SFD_New;

		// CreateThread and process the request
		hThread = CreateThread( 
        NULL,																		// default security attributes 
        0,                           												// use default stack size  
        ProcessRequest,                 											// thread function 
        &Argument,                													// argument to thread function 
        0,                           												// use default creation flags 
        &dwThreadId);                												// returns the thread identifier 
		
		if (hThread != NULL)														// If the thread was created, destroy it
		{
			CloseHandle( hThread );
		}

	}
    ReturnCode = SWEBS_RETURN_SUCCESS;                                              // We know the server was successful

	closesocket(SFD_Listen);
	return;
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

		delete New;																	// Destroy the connection
	}
	closesocket(Arg->SFD);
	return 0;
}

//---------------------------------------------------------------------------------------------
//			Control Handler
//---------------------------------------------------------------------------------------------
void ControlHandler(DWORD request) 
{ 
	switch(request) 
	{ 
	case SERVICE_CONTROL_STOP: 
		SERVER_STOP = true;

        ServiceStatus.dwWin32ExitCode = 0; 
        ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
        SetServiceStatus (hStatus, &ServiceStatus);
        return; 
 
	case SERVICE_CONTROL_SHUTDOWN: 
        SERVER_STOP = true;

        ServiceStatus.dwWin32ExitCode = 0; 
        ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
        SetServiceStatus (hStatus, &ServiceStatus);
        return; 
        
	default:
        break;
	} 
 
    // Report current status
    SetServiceStatus (hStatus, &ServiceStatus);
    return; 
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
//			PrintAccepts
//---------------------------------------------------------------------------------------------
void PrintAccepts(const map<string, bool>::value_type& p)
{
	TestLog(p.first);
	TestLog(" = ");
	TestLog((p.second ? "true" : "false"));
	TestLog("\n");
}

//---------------------------------------------------------------------------------------------
//			InstallService
//---------------------------------------------------------------------------------------------
bool InstallService()
{
    // Get the application location from the registry
    HKEY hKey;																		// Handle for the key
	unsigned long dwDisp;															// Disposition
	RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\SWS"), 0,
               NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dwDisp);

	unsigned char Buffer[_MAX_PATH];
	unsigned long DataType;
	unsigned long BufferLength = sizeof(Buffer);
	
	RegQueryValueEx(hKey, "AppPath", NULL, &DataType, Buffer, &BufferLength);

	string AppPath;
	AppPath = (char *)Buffer;											            // Copy the appPath

	if ( AppPath.empty())												            // If the key was not there
	{
		return false;
	}

	RegCloseKey(hKey);                                                              // Close
    
    string SWEBS_Exe_Location = AppPath;
    SWEBS_Exe_Location += "\\";
    SWEBS_Exe_Location += "SWEBS.exe";
    SC_HANDLE schSCManager = OpenSCManager(
		NULL,
		SERVICES_ACTIVE_DATABASE,
		SC_MANAGER_CREATE_SERVICE
	);

	LPCTSTR lpszBinaryPathName = SWEBS_Exe_Location.c_str();
	
    SC_HANDLE schService = CreateService( 
        schSCManager,                                                               // SCManager database 
        "SWEBS Web Server",                                                         // Name of service 
        "SWEBS Web Server",		                                                    // Service name to display 
        SERVICE_ALL_ACCESS,                                                         // Desired access 
        SERVICE_WIN32_OWN_PROCESS,                                                  // Service type 
        SERVICE_AUTO_START,		                                                    // Start type 
        SERVICE_ERROR_NORMAL,                                                       // Error control type 
        lpszBinaryPathName,                                                         // Service's binary 
        NULL,                                                                       // No load ordering group 
        NULL,                                                                       // No tag identifier 
        NULL,                                                                       // No dependencies 
        NULL,                                                                       // LocalSystem account 
        NULL);                     
 
    if (schService == NULL) 
	{
		return false;
    }
 
    CloseServiceHandle(schService);
	CloseServiceHandle(schSCManager);
    return true;
}

//---------------------------------------------------------------------------------------------
//			UninstallService
//---------------------------------------------------------------------------------------------
bool UninstallService()
{
    HANDLE hService;		                                                        // Handle to the service
	SC_HANDLE schSCManager;

	schSCManager = OpenSCManager(
		NULL,
		SERVICES_ACTIVE_DATABASE,
		SC_MANAGER_ALL_ACCESS
	);

	hService = OpenService                                                          // Open the service
	(
		schSCManager,
		"SWEBS Web Server",
		SC_MANAGER_ALL_ACCESS
	);

	if (DeleteService(hService))                                                    // Try to delete it
		return true;
	else
		return false;   
}