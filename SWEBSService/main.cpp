//---------------------------------------------------------------------------------------------
//  SWEBS
//  ------------------------
//  This is the source of the SWEBS Web Server service. Its job is simple - register itself
//  as a service, then start the program bin\SWEBSEngine.exe. It can also install/uninstall
//  itself with the /u and /i arguments.
//---------------------------------------------------------------------------------------------

//---------------------------------------------------------------------------------------------
//			Includes
//---------------------------------------------------------------------------------------------
#pragma warning(disable:4786)
#pragma warning(disable:4089)
#include <windows.h>
#include <string>
#include <iostream>

using namespace std;

//---------------------------------------------------------------------------------------------
//			Function Declarations
//---------------------------------------------------------------------------------------------
void ServiceMain();
void ControlHandler(DWORD request); 
void TestLog(string);
bool InstallService();
bool UninstallService();

//---------------------------------------------------------------------------------------------
//			Globals
//---------------------------------------------------------------------------------------------
bool SERVER_STOP = false;

SERVICE_STATUS          ServiceStatus; 
SERVICE_STATUS_HANDLE   hStatus; 
STARTUPINFO si;
PROCESS_INFORMATION pi;

int ReturnCode;                                                                     // Number for main() to return, can be set from any function

const int SWEBS_RETURN_UNKNOWN          = 0x00;                                     // Unknown error occured
const int SWEBS_RETURN_SUCCESS          = 0x01;                                     // Server ran fine
const int SWEBS_RETURN_COULDNOTBIND     = 0x02;                                     // Could not bind() to port
const int SWEBS_RETURN_CONFIGNOTLOADED  = 0x03;                                     // Could not load config file
const int SWEBS_RETURN_COULDNOTLISTEN   = 0x04;                                     // Could not listen()
const int SWEBS_RETURN_COULDNOTACCEPT   = 0x05;                                     // Could not accept()

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
        SERVICE_TABLE_ENTRY ServiceTable[2]; 
	    ServiceTable[0].lpServiceName = "SWEBS Web Server";							// Name of service
	    ServiceTable[0].lpServiceProc = (LPSERVICE_MAIN_FUNCTION)ServiceMain;		// Main function of service

	    ServiceTable[1].lpServiceName = NULL;										// Must create a null table
	    ServiceTable[1].lpServiceProc = NULL;
	    // Start the control dispatcher thread for our service
	    StartServiceCtrlDispatcher(ServiceTable);									// Jumps to the serice function  

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

  	// Report that the service is running
	ServiceStatus.dwCurrentState = SERVICE_RUNNING; 
	SetServiceStatus (hStatus, &ServiceStatus);

	//-----------------------------------------------------------------------------------------
	// Step 3: Start web server
	//-----------------------------------------------------------------------------------------
    // Get the application location from the registry
    string AppPath;
    string EnginePath;
    HKEY hKey;																		// Handle for the key
	unsigned long dwDisp;															// Disposition
	unsigned long ExitCode;

    RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\SWS"), 0,
               NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dwDisp);
	unsigned char Buffer[_MAX_PATH];
	unsigned long DataType;
	unsigned long BufferLength = sizeof(Buffer);
	RegQueryValueEx(hKey, "AppPath", NULL, &DataType, Buffer, &BufferLength);
	AppPath = (char *)Buffer;

    // We know where we were installed to. Run the program bin\SWEBSEngine.exe
    EnginePath = AppPath;
    EnginePath += "\\bin\\SWEBSEngine.exe";

    ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);
    ZeroMemory( &pi, sizeof(pi) );

    // Start the child process. 
    if( !CreateProcess( EnginePath.c_str(),                                         // No module name (use command line). 
        NULL,                                                                       // Command line. 
        NULL,                                                                       // Process handle not inheritable. 
        NULL,                                                                       // Thread handle not inheritable. 
        FALSE,                                                                      // Set handle inheritance to FALSE. 
        0,                                                                          // No creation flags. 
        NULL,                                                                       // Use parent's environment block. 
        AppPath.c_str(),                                                            // Starting directory. 
        &si,                                                                        // Pointer to STARTUPINFO structure.
        &pi )                                                                       // Pointer to PROCESS_INFORMATION structure.
    ) 
    {
        TestLog("Could not create process for SWEBSEngine.exe");
    }

    SERVER_STOP = false;
    while (SERVER_STOP == false)
    {
        Sleep(3000);                                                                // Check every 3 seconds if we have to stop
    }
    GetExitCodeProcess(pi.hProcess, &ExitCode);

    // Close process and thread handles. 
    CloseHandle( pi.hProcess );
    CloseHandle( pi.hThread );
	
	//-----------------------------------------------------------------------------------------
	// Step 5: Handle Requests
	//-----------------------------------------------------------------------------------------
	SERVER_STOP = false;
    
    ReturnCode = SWEBS_RETURN_SUCCESS;                                              // We know the server was successful

	return;
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
    
        SendMessage((HWND)pi.hProcess , WM_CLOSE, NULL, NULL);

        ServiceStatus.dwWin32ExitCode = 0; 
        ServiceStatus.dwCurrentState = SERVICE_STOPPED; 
        SetServiceStatus (hStatus, &ServiceStatus);
        return; 
 
	case SERVICE_CONTROL_SHUTDOWN: 
        SERVER_STOP = true;

        SendMessage((HWND)pi.hProcess, WM_CLOSE, NULL, NULL);

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
	log = fopen("c:\\sws\\testlog.log", "a+");
	if (log == NULL)
      return ;
	fprintf(log, "%s", Data.c_str());
	fclose(log);
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