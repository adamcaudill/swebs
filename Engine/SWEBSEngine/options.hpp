#ifndef OPTIONSHPP
#define OPTIONSHPP 1
#pragma warning(disable:4786)
//----------------------------------------------------------------------------------------------------
#include <windows.h>
#include <map>
#include <sstream>
#include "CGI.hpp"

// These files are to be used for parsing XML documents (the config file) and must be downloaded
//  from: www.xml-parser.com
#include <CkXml.h>
#include <CkString.h>
#include <CkSettings.h>

using namespace std;

int StringToInt(string);
string IntToString(int num);
int CalcMonth(string Month);
string DecodeURL(string URL);

//----------------------------------------------------------------------------------------------------
//			Options class - derived from configuration file
//----------------------------------------------------------------------------------------------------
class OPTIONS
{
  public:
	string Servername;																// Name of this server - Ie, Central Online (SWS)
	int Port;																		// Port number to listen on (80)
	string IPAddress;                                                               // IP Address for the server to listen on
    string WebRoot;																	// Path to root web folder (C:\WebRoot)
	int MaxConnections;																// Number of connections at once (20)
	string Logfile;																	// Path/name of log file (c:\SWS\logfile.log)
	string ErrorLog;                                                                // Error log file
    map <string, string> CGI;														// Map of extension/interpreter for CGI scripts (ie, CGI["php"] = "C:\PHP.exe"
	map <int, string> IndexFiles;													// Files that will be used as auto indexes of folders (index.htm)
	int Timeout;																	// Idle time for each connection before time out and closure
	map <string, string> MIMETypes;													// MIME types
	map <string, bool> Binary;														// Files that should be opened as binary
	bool AllowIndex;																// Are we allowed to index files
	map <int, string> ErrorCode;													// List of number to string mapped error codes, ie:
																					//  ErrorCode[404] = "File Not Found";
	NON_REQUEST_SPECIFIC_CGI CGIVariables;                                          // CGI Environment variables
    string ErrorDirectory;															// Folder where custom error pages are kept
	bool ReadSettings();															// Read in the settings from the config file
    bool LogError(string Text);                                                     // Logs some error text
};
extern OPTIONS Options;


//----------------------------------------------------------------------------------------------------
//			Virtual Host class - derived from configuration file
//----------------------------------------------------------------------------------------------------
class VIRTUALHOST
{
public:
	string Name;																	// Name of this virtual host (ie, RateMyPoo)
	string HostName;																// Internet Address of host (www.ratemypoo.com)
	map <int, string> IndexFiles;													// Files that will be used as auto indexes of folders (index.htm)
	string Root;																	// Root folder of files for this VH (ie, c:\RateMyPoo)
	string Logfile;																	// Path/name of log file (C:\RateMyPoo\logfile.log)
    friend bool operator<(const VIRTUALHOST lhs, const VIRTUALHOST rhs);
    friend bool operator>(const VIRTUALHOST lhs, const VIRTUALHOST rhs);
};

//----------------------------------------------------------------------------------------------------
//			Virtual Host Index (VHI) class. Keeps an index of all the Virtual hosts
//----------------------------------------------------------------------------------------------------
class VirtualHostIndex
{
  public:
	int NumberOfHosts;
    map <int, string> HostNumbers;                                                  // Keep a number indexed list of the virtual host names
	map <string, VIRTUALHOST> Host;													// Map Internet address to appropriate virtual host
};
extern VirtualHostIndex VHI;

//----------------------------------------------------------------------------------------------------
#endif