#ifndef SWEBS_GLOBALS_HPP
#define SWEBS_GLOBALS_HPP
//----------------------------------------------------------------------------------
//		SWEBSGlobals.hpp
//		------------------
//		This header defines the SWEBSGlobals class. The class contains EVERYTHING
//		globally used throughout the SWEBS Web Server. General settings, options, 
//		statistics and virtual hosts are all defined here.
//		
//		Copyright Paul Stovell, 2003. All rights reserved.
//----------------------------------------------------------------------------------
//		INCLUDES
//----------------------------------------------------------------------------------
#include <windows.h>
#include <string>
#include <map>
#include <winsock.h>
#include "SWEBSCGI.hpp"

using namespace std;
//----------------------------------------------------------------------------------
//		SWEBSVirtualHost class
//----------------------------------------------------------------------------------
class SWEBSVIRTUALHOST
{
public:
	bool LoadStats();                                                               // Loads stats for this host
    
    string Name;																	// Name of this virtual host (ie, RateMyPoo)
	string HostName;																// Internet Address of host (www.ratemypoo.com)
	map <int, string> IndexFiles;													// Files that will be used as auto indexes of folders (index.htm)
	string Root;																	// Root folder of files for this VH (ie, c:\RateMyPoo)
	string Logfile;																	// Path/name of log file (C:\RateMyPoo\logfile.log)
    friend bool operator<(const SWEBSVIRTUALHOST lhs, const SWEBSVIRTUALHOST rhs);	// These must be used in order for us to create maps of the classes
    friend bool operator>(const SWEBSVIRTUALHOST lhs, const SWEBSVIRTUALHOST rhs);
    
    // Statistics
	map <string, int> PageRequests;                                               	// Page requests (per VH)
    unsigned long NumberOfRequests;                                               	// Number of requests
    unsigned long BytesSent;   														// Bytes sent to this host
};

//----------------------------------------------------------------------------------
//		SWEBSGlobals class
//----------------------------------------------------------------------------------
class SWEBSGLOBALS
{
public:
	// Methods
    bool LoadStats();																// Loads info from the stats file
	bool ReadSettings();														  	// Read in the settings from the config file
    bool LogError(string Text);
   	bool WriteStatsFile();                                                    		// Writes all statistics to the stats file    
    
   	// Settings
	string Servername;															  	// Name of this server - Ie, Central Online (SWS)
	int Port;																      	// Port number to listen on (80)
	string IPAddress;                                                             	// IP Address for the server to listen on
    string WebRoot;														 	      	// Path to root web folder (C:\WebRoot)
	int MaxConnections;															  	// Number of connections at once (20)
	string Logfile;																  	// Path/name of log file (c:\SWS\logfile.log)
	string ErrorLog;                                                              	// Error log file
    map <string, string> CGI;												      	// Map of extension/interpreter for CGI scripts (ie, CGI["php"] = "C:\PHP.exe"
	map <int, string> IndexFiles;											      	// Files that will be used as auto indexes of folders (index.htm)
	int Timeout;																  	// Idle time for each connection before time out and closure
	map <string, string> MIMETypes;												  	// MIME types
	map <string, bool> Binary;													  	// Files that should be opened as binary
	bool AllowIndex;															  	// Are we allowed to index files
    map <string, string> IsapiDLLExtensions;                                      	// Map extensions of files to an ISAPI DLL
    map <string, HINSTANCE> IsapiDLL;                                             	// Map extensions to dll instances
    map <int, string> ErrorCodes;										     	  	// List of number to string mapped error codes, ie:
	int SFD_Listen;																	// Socket Descriptor we listen on
	struct sockaddr_in ServerAddress;												// Servers address structure

    // Statistics		
    unsigned long NumberOfRequests;                                               	// Number of connections served by the server
    unsigned long TotalNumberOfRequests;                                          	// Total number of connections served
    unsigned long BytesSent;                                                      	// Total number of bytes served
    unsigned long TotalBytesSent;                                                 	// Total number of bytes sent for VH's and normal requests
    string LastRestart;                                                           	// Last time the server was restarted
    map <string, int> PageRequests;                                               	// Page requests
    map <int, string> PageRequestIndex;  
    
    // General settings
	NON_REQUEST_SPECIFIC_CGI CGIVariables;                                        	// CGI Environment variables
    string ErrorDirectory;														  	// Folder where custom error pages are kept

    // Virtual host indexing
	int NumberOfHosts;
    map <int, string> HostNumbers;                                                  // Keep a number indexed list of the virtual host names
	map <string, SWEBSVIRTUALHOST> Host;											// Map Internet address to appropriate virtual host    
};

//----------------------------------------------------------------------------------
//		Externals	
//----------------------------------------------------------------------------------
extern SWEBSGLOBALS SWEBSGlobals;													// Make our SWEBSGlobals class properly global
extern DWORD WINAPI HandleStatsFile(LPVOID lpParam );								// Make it so we can call the handle function

//----------------------------------------------------------------------------------
#endif