#ifndef SWEBSCONNECTIONHPP
#define SWEBSCONNECTIONHPP 1
//--------------------------------------------------------------------------------
//			SWEBSCONNECTION.hpp
//          -------------------
//          This is it. This is the most important part of the SWEBS source code,
//          the code that handles all the requests. Everything the client will
//          see comes through this class. 
//
//          When a request comes in, the main() function creates a new instance
//          of the CONNECTION class, whos job it is to handle the connection. Main
//          calls ReadRequest() to read in all the information from the client,
//          and then calls HandleRequest(), which handles it.
//
//          This is how the CONNECTION class is created:
//          CONNECTION::CONNECTION(
//              int SFD_SET,                - Socket descriptor for this connection
//              struct sockaddr_in,         - Sockaddr_in structure used
//
//--------------------------------------------------------------------------------																				
//			INCLUDES																
//--------------------------------------------------------------------------------
#include "SWEBSGlobals.hpp"
#include "SWEBSISAPI.hpp"
#include <windows.h>
#include <string>
#include <winsock.h>
#include <map>
#include <sstream>
#include <algorithm>
#include <fstream>
#include <stdio.h>
#include <ctime>
#include <httpext.h>
#include <httpfilt.h>
#include <header.h>

using namespace std;

//---------------------------------------------------------------------------------------------
//			Connection class
//---------------------------------------------------------------------------------------------
class CONNECTION
{
  public:
	CONNECTION(int SFD_SET, struct sockaddr_in);	                                // Constructor
    CONNECTION();
    ~CONNECTION();
    
	bool LogConnection();														    // Logs connection to the appropriate log
	bool ReadRequest();														   	    // Reads the request and sets values
	bool HandleRequest();														    // Handles the request
    bool LogText(string);                                                           // Logs a string

	// Methods
	bool SetFileType();															    // Sets if the file is a script or binary
	bool IndexFolder();															    // Indexes the folder by listing all the files
	bool SendText();															    // Sends the requested file if it is text
	bool SendCGI();																    // Sends the requested file if it is a script
	bool SendBinary();															    // Sends the requested file if it is binary
	bool SendError();															    // Outputs the appropriate error code
	friend bool SendISAPI(CONNECTION * Conn);                                       // Sends the ISAPI request
    string CalculateSize();													        // Outputs the file size
	bool ModifiedSince(string Date);										        // Was the file modifed since...
	bool UnModifiedSince(string Date);											    // Is the file Unmodified since...

	// Properties
	int SFD;																   	    // Socket descriptor of connection
	struct sockaddr_in ClientAddress;											    // Client address structure
	SWEBSVIRTUALHOST * ThisHost;												    // This virtual host host
    REQUEST_SPECIFIC_CGI CGIVariables;                                              // CGI Environment variables

	string FullRequest;															    // The entire input from the client
	string RequestType;															    // Type of request (POST, GET etc)
	string FileRequested;														    // String folling GET
	  string QueryString;														    // Anything after the '?' in the file
	  string Extension;															    // Extension of file requested
	  string RealFile;															    // Real path to file
	  string RealFileDate;														    // Last Modification date of RealFile
	string HTTPVersion;															    // HTTP version of the client
	stringstream PostData;														    // Data supplied AFTER the double newline, for post requests

	string Headers;																    // Headers to be sent with the file
	map <string, bool> Accepts;													    // MIME types the client accepts
	string UserAgent;															    // Browser used by the user
	string HostRequested;														    // Host: from browser
	string From;																    // From: value (email address normally)
	string ConnectionType;														    // Connection: type (keep alive normally)
    string Referer;                                                                 // Referer: file that refered the user to this document

	string Date;																    // Date/time of this request
	string ModifiedSinceStr;													    // IfModifiedSince string
	string UnModifiedSinceStr;													    // IfUnModifiedSince string
	bool UseModDate;															    // Do we use an If-Modified-Since
	bool UseUnModDate;															    // Do we use the If-Unmodified-Since

	int Status;																	    // Status code for request (404, 200 etc)
	bool UseVH;																	    // Does the connection use a virtual host or a real one
	bool IsFolder;																    // Is the file requested a folder or a file
	bool IsBinary;																    // Is the file binary?
	bool IsScript;																    // Is the file a script
	bool IsIsapi;                                                                   // The file requested uses ISAPI
    bool IsAbsolute;															    // Did the client use an absolute address
};

//--------------------------------------------------------------------------------
#endif