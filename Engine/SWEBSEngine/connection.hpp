#ifndef CONNECTIONHPP
#define CONNECTIONHPP 1
//---------------------------------------------------------------------------------------------
/*
			CONNECTION.HPP
			--------------
			This file contains the definitions of the CONNECTION class.

			Upon creating a new connection the calling function hands the connection
			the Socket Descriptor that the connection is on, and a sockaddr_in 
			structure which contains information about the client. For example:

				CONNECTION * NewRequest = new CONNECTION (SFD, CLA);

			Where SFD is the Socket Descriptor, and CLA is a sockaddr_in structure.
			The constructor then receives the request from the client, and breaks it up
			into all the possible values that might be used in the connection, such as
			the type of request, file requested, host, MIME types accepted by the client,
			etc. 

			Then, the constructor checks that the file requested by the client exists.
			If it does, it checks if it is a folder or a file. If its a file it gets the
			extension and cross references it with the appropriate MIME type. Then, it 
			checks if the file is a CGI script or a binary file, and sets the flags
			appropriately. If it is just a plain text file, neither flags are set.

*/
//---------------------------------------------------------------------------------------------
#include <windows.h>
#include <string>
#include <winsock.h>
#include <map>
#include <sstream>
#include <algorithm>
#include <fstream>
#include <stdio.h>
#include <ctime>
#include "options.hpp"																// Contains definitions of the VHI (Virtual Host Index)
#include "stats.hpp"
#include "cgi.hpp"

using namespace std;
//---------------------------------------------------------------------------------------------
//			Connection class
//---------------------------------------------------------------------------------------------
class CONNECTION
{
  public:
	CONNECTION(int SFD_SET, struct sockaddr_in);									// Constructor
	CONNECTION();
	bool LogConnection();															// Logs connection to the appropriate log
	bool ReadRequest();																// Reads the request and sets values
	bool HandleRequest();															// Handles the request

	// Methods
	bool SetFileType();																// Sets if the file is a script or binary
	bool IndexFolder();																// Indexes the folder by listing all the files
	bool SendText();																// Sends the requested file if it is text
	bool SendCGI();																	// Sends the requested file if it is a script
	bool SendBinary();																// Sends the requested file if it is binary
	bool SendError();																// Outputs the appropriate error code
	bool Send(int SFD, string Text);                                                // Our own version of send()
    bool Send(int SFD, string Text, int Length, int Number);                        // For older calls
    bool SendSingle(int SFD, char C);                                               // Sends a single char
    bool LogText(string);															// Logs some text. Used only for testing
	string CalculateSize();															// Outputs the file size
	bool ModifiedSince(string Date);												// Was the file modifed since...
	bool UnModifiedSince(string Date);												// Is the file Unmodified since...

	// Properties
	int SFD;																		// Socket descriptor of connection
	struct sockaddr_in ClientAddress;												// Client address structure
	VIRTUALHOST *ThisHost;															// This virtual host host
    REQUEST_SPECIFIC_CGI CGIVariables;                                              // CGI Environment variables

	string FullRequest;																// The entire input from the client
	string RequestType;																// Type of request (POST, GET etc)
	string FileRequested;															// String folling GET
	  string QueryString;															// Anything after the '?' in the file
	  string Extension;																// Extension of file requested
	  string RealFile;																// Real path to file
	  string RealFileDate;															// Last Modification date of RealFile
	string HTTPVersion;																// HTTP version of the client
	string PostData;																// Data supplied AFTER the double newline, for post requests

	string Headers;																	// Headers to be sent with the file
	map <string, bool> Accepts;														// MIME types the client accepts
	string UserAgent;																// Browser used by the user
	string HostRequested;															// Host: from browser
	string From;																	// From: value (email address normally)
	string Connection;																// Connection: type (keep alive normally)
    string Referer;                                                                 // Referer: file that refered the user to this document

	string Date;																	// Date/time of this request
	string ModifiedSinceStr;														// IfModifiedSince string
	string UnModifiedSinceStr;														// IfUnModifiedSince string
	bool UseModDate;																// Do we use an If-Modified-Since
	bool UseUnModDate;																// Do we use the If-Unmodified-Since

	int Status;																		// Status code for request (404, 200 etc)
	bool UseVH;																		// Does the connection use a virtual host or a real one
	bool IsFolder;																	// Is the file requested a folder or a file
	bool IsBinary;																	// Is the file binary?
	bool IsScript;																	// Is the file a script
	bool IsAbsolute;																// Did the client use an absolute address
};

#endif