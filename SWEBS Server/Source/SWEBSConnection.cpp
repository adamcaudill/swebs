//----------------------------------------------------------------------------------
//			SWEBSConnection.cpp
//			-------------------
//		    This is it. This is the most important part of the SWEBS source code,
//          the code that handles all the requests. Everything the client will
//          see comes through this class. 
//
//          When a request comes in, the main() function creates a new instance
//          of the CONNECTION class, whos job it is to handle the connection. Main
//          calls ReadRequest() to read in all the information from the client,
//          and then calls HandleRequest(), which handles it.
//
//          Copyright Paul Stovell, 2003. All rights reserved.
//
//----------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include "../Include/SWEBSGlobals.hpp"
#include "../Include/SWEBSConnection.hpp"
#include "../Include/SWEBSUtilities.hpp"
#include "../Include/SWEBSHeaderMap.hpp"
#include "../Include/SWEBSSocket.hpp"
#include "../Include/SWEBSISAPI.hpp"
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

//----------------------------------------------------------------------------------
//			GLOBALS
//----------------------------------------------------------------------------------
#define FILE_INVALID						4294967295                              // Returned by GetFileAttributes when something doesn't exist


//----------------------------------------------------------------------------------
//			CONNECTION::CONNECTION
//----------------------------------------------------------------------------------
CONNECTION::CONNECTION(int SFD_SET, struct sockaddr_in CA)
{
	//------------------------------------------------------------------------------
	//			Change settings
	//------------------------------------------------------------------------------
	SFD = SFD_SET;											    					// Set the socket descriptor
	ClientAddress = CA;															    // Assign the client address
	UseVH = false;																    // Dont use a Virtual host by default
	IsFolder = false;															    // By default its not a folder
	IsBinary = false;															    // Nor is it binary
	IsScript = false;															    // Or a CGI script
	IsIsapi = false;															    // Or isapi!
    IsAbsolute = false;															    // And the URL is not an absolute URL
	Status = 200;																    // But the file is always served fine
	// Before we loop through keys/values, make sure HeaderMapInit is called
	HeaderMapInit();
    ThisHost = new SWEBSVIRTUALHOST;
}

CONNECTION::CONNECTION()
{
	
}

CONNECTION::~CONNECTION()
{
    delete ThisHost;
}
//---------------------------------------------------------------------------------------------
//			Connection::SetFileType
//---------------------------------------------------------------------------------------------
bool CONNECTION::SetFileType()
{
	// Get the files extension
	int X = RealFile.find_last_of('.') + 1;										  // Get the last '.' in the string

	if (X > 0)
	{
		for (X; RealFile[X] != '\0'; X++)									      // Copy everything after that to the extension
		{
			Extension+= RealFile[X];						
		}

		if (SWEBSGlobals.CGI[Extension].length() > 1)								  // If theres an enrty in the CGI map, its a script
		{
			IsIsapi = false;
			IsScript = true;
			IsBinary = false;
		}
		else if (SWEBSGlobals.Binary[Extension])									      // Or if theres an entry in the Binary map, its binary
		{
			IsIsapi = false;
			IsScript = false;
			IsBinary = true;
		}
		else if(SWEBSGlobals.IsapiDLL[Extension] != NULL)
        {
            IsIsapi = true;
            IsBinary = false;
            IsScript = false;
        }
        else																	  // It must be text then
		{
			IsIsapi = false;
			IsBinary = false;
			IsScript = false;
		}
		return true;
	}
	else return false;															  // There was no extension
}

//--------------------------------------------------------------------------------
//			Connection::ReadRequest
//--------------------------------------------------------------------------------
bool CONNECTION::ReadRequest()
{
	//----------------------------------------------------------------------------
	//			Set request variables
	//-----------------------------------------------------------------------------------------
	// First, read in the whole request
    string Word;
    FullRequest = SWEBSSocket::Recieve(SFD);
    //-----------------------------------------------------------------------------------------------------
	// Break off any POST data following a double newline
	int X = 0;
    // Find the \n\n or end of string, whichever comes first
    for (X = 0; FullRequest[X] != '\0'; X++ )
    {
        if (FullRequest[X] == '\n')
        {
            if (FullRequest[X+1] == '\n')
            {
                break;
            }
        }
    }
    // Now copy whats after that \n\n
    for (X; FullRequest[X] != '\0'; X++)
    {
        PostData << FullRequest[X];
    }
    
    //-----------------------------------------------------------------------------------------------------
	// Split it into words
	istringstream IS(FullRequest);													// Create an istringstream class

	//-----------------------------------------------------------------------------------------------------
	IS >> Word;																		// The first word will be the request type
	
	//strupr(Word.c_str());										
	if (!( strcmpi(Word.c_str(), "POST") ||											// Check to see if its a method we support
		 strcmpi(Word.c_str(), "GET")  ||
		 strcmpi(Word.c_str(), "HEAD" )))
	{
		// Its not a request we like
		Status = 501;																// Send a 501 Not Implemented
		return false;
	}
	
    //-----------------------------------------------------------------------------------------------------
	RequestType = Word;																// The request type was ok, assign it
	IS >> FileRequested;															// The next word will be the requested file
	IS >> HTTPVersion;																// Then the HTTP version
	IS >> Word;

    
	//-----------------------------------------------------------------------------------------------------
	// Map keys to values
	while (Word.length())
	{
		// Loop through, mapping keys to values
		if (HeaderMap[Word] != NULL)												// Has there been a function assigned to this header?
		{
            HeaderMap[Word](IS, this);												// If there has, call it!
		}
		// Now, check if its a MIME type. If it DOES NOT have ':' and DOES have '/', then its probably mime
		else if (!strstr(Word.c_str(), ":") && strstr(Word.c_str(), "/"))
		{
			if (strstr(Word.c_str() , ","))
			{
				// There is a comma at the end, get rid of it
				int Y = Word.length() - 1;
				Word[Y] = '\0';
			}
			Accepts[Word] = true;
		}
		IS >> Word;
	}
    
    //-----------------------------------------------------------------------------------------------------
	// First, if the request is HTTP/1.1, there must be a host field
	if ( !strcmpi (HTTPVersion.c_str(), "HTTP/1.1") )
	{
		if (HostRequested.length() <= 0)
		{
			Status = 400;															// No host was specified. Send 400 Bad Request
			return false;
		}
	}
	
    //-----------------------------------------------------------------------------------------------------
	// Cut off absolute URL, making us able to serve all future HTTP versions.
	if (strstr( FileRequested.c_str() , "http://" ))								// If its an absolute URL
	{	 
		IsAbsolute = true;															// Start at the end of the http://
		int CurrLetter = FileRequested.find_first_of("http://") + 7;
		
		for (CurrLetter; FileRequested[CurrLetter] != '/'; CurrLetter++ )
		{
			HostRequested += FileRequested[CurrLetter];								// Copy to the new host.
		}
	} 
	
    //-----------------------------------------------------------------------------------------------------
	// Figue out the virtual host
	/*ThisHost = &SWEBSGlobals.Host[HostRequested];						
	if(ThisHost->Root.length() < 1)													// If there is an entry for the host in the VHI 
		UseVH = false;																//  then it is a virtual host
	else UseVH = true;
    */
    //Beep(175, 500);
    //Sleep(2000);
	//-----------------------------------------------------------------------------------------------------
	// Cut off query string
	X = 0;
    X = FileRequested.find_last_of('?') + 1;									// Get the last '?' in the string
	if (X > 0)
	{
		int Y = X;																	// Save position of X
		for (X; FileRequested[X] != '\0'; X++)										// Copy everything after that to the extension
		{
			QueryString+= FileRequested[X];						
		}
		FileRequested[Y - 1] = '\0';												// Chop it off at the '?'
	}
	
    //-----------------------------------------------------------------------------------------------------
    // URL Encoding
    // %20 = " "
    int S = 0;
    while ((S = FileRequested.find("%20", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, " ");
    }
    // %24 = "$"
    while ((S = FileRequested.find("%24", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "$");
    }
    // %26 = "&"
    while ((S = FileRequested.find("%26", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "&");
    }
    // %2B = "+"
    while ((S = FileRequested.find("%2B", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "+");
    }
    while ((S = FileRequested.find("%2C", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, ",");
    }
    while ((S = FileRequested.find("%2F", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "/");
    }
    while ((S = FileRequested.find("%3A", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, ":");
    }
    while ((S = FileRequested.find("%3B", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, ";");
    }
    while ((S = FileRequested.find("%3D", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "=");
    }
    while ((S = FileRequested.find("%3F", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "?");
    }
    while ((S = FileRequested.find("%40", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "@");
    }
    while ((S = FileRequested.find("%22", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "\"");
    }
    while ((S = FileRequested.find("%3C", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "<");
    }
    while ((S = FileRequested.find("%3E", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, ">");
    }
    while ((S = FileRequested.find("%23", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "#");
    } 
    while ((S = FileRequested.find("%25", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "%");
    }
    while ((S = FileRequested.find("%7B", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "{");
    }
    while ((S = FileRequested.find("%7D", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "}");
    }
    while ((S = FileRequested.find("%7C", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "|");
    }
    while ((S = FileRequested.find("%5C", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "\\");
    }
    while ((S = FileRequested.find("%5E", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "^");
    }
    while ((S = FileRequested.find("%7E", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "~");
    }
    while ((S = FileRequested.find("%5B", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "[");
    }
    while ((S = FileRequested.find("%5D", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "]");
    }
    while ((S = FileRequested.find("%60", 0)) != string::npos)
    {
        FileRequested.replace(S , 3, "`");
    }

    //-----------------------------------------------------------------------------------------------------
	// Change slashes from *nix to windows
	for (int Z = 0; FileRequested[Z] != '\0'; Z++)									// Replace / with \ 
	{
		if (FileRequested[Z] == '/') FileRequested[Z] = '\\';
	}
	
    //-----------------------------------------------------------------------------------------------------
	// Assign full path based on virtual hosts
	if (UseVH) RealFile = ThisHost->Root + FileRequested;
	else RealFile = SWEBSGlobals.WebRoot + FileRequested;
		
    // Check for a "..", if found send a 404. Because this will allow them to go one folder back, and 
	//  then get files from there, effectively giving full access to the system - thanks to Adam for pointing this out!
	if (strstr(RealFile.c_str() , ".."))
	{
		Status = 404;
		return false;
	}

	//-----------------------------------------------------------------------------------------------------
	// Check if the file is a folder
	DWORD hFile = GetFileAttributes(RealFile.c_str());
	if (hFile == FILE_INVALID)
	{
		
		Status = 404;																// File does not exist. Return error 404
		return false;
	}
	
    //-----------------------------------------------------------------------------------------------------
	if (hFile & FILE_ATTRIBUTE_DIRECTORY)				
	{
		IsFolder = true;															// Is a folder
	}
	else	
	{								
		IsFolder = false;															// Is not folder
	}

    //-----------------------------------------------------------------------------------------------------	
	SetFileType();																	// Set whether the file is binary or a script
	Status = 200;																	// It passed all the tests, therefore its ok
	
    return true;
}

//---------------------------------------------------------------------------------------------
//			Connection::HandleRequest
//---------------------------------------------------------------------------------------------
bool CONNECTION::HandleRequest()
{
	// Get the time:
	time_t now;
	struct tm * tm_now = NULL;
	char buff[BUFSIZ];

	now = time ( NULL );
	tm_now = gmtime(&now);
	strftime ( buff, sizeof buff, "%a, %d %b %Y %H:%M:%S GMT", tm_now );
	Date = buff;
    
	//----------------------------------------------------------
	// Do the request
	if (Status == 200)
	{
		if (UseModDate == true)														// If we got a last modified header:
		{
			if (!ModifiedSince(ModifiedSinceStr))									// If the file has not been modifed
			{
				Status = 304;														// Send them a 304 not modifed
			}
		}
		if (UseUnModDate == true)													// If we got a last modified header:
		{
			if (!UnModifiedSince(UnModifiedSinceStr))								// If the file has not been modifed
			{
				Status = 402;														// Send them a 402 not modifed
			}
		}
		// Output the file
		if (IsFolder)																// Request was a folder
		{
			// Search for an index file
			int X = 0;
			bool Found = false;
			while ((X < SWEBSGlobals.IndexFiles.size()) && !Found)						// While theres still index files there
			{
				string File = RealFile;
				File += "\\";
				File += SWEBSGlobals.IndexFiles[X];
				ifstream hFile (File.c_str());
				if (!hFile)
					X++;															// Increase the counter
				else
					Found = true;													// Found one!
			}
			if (Found)
			{
				RealFile += "\\";
				RealFile += SWEBSGlobals.IndexFiles[X];									// Make the file were looking for the appropriate index file
				IsFolder = false;									
			}																		// Not a folder anymore, break out and send as a file
			// If we are allowed to index:
			if (SWEBSGlobals.AllowIndex == true && IsFolder)								// Ensure its still a folder
			{
				bool hResult = IndexFolder();										// Try to index the folder
				if (hResult == false)												// The folder could not be indexed. Report
				{
					Status = 404;													// Set status code
				}
			}
            SetFileType();
		}
		
        // We are ready to process. Set CGI environment varaibles
        if ( !strcmpi(RequestType.c_str(), "POST") )
            CGIVariables.HTTP_MAP["CONTENT_LENGTH"] = IntToString(PostData.gcount());
        CGIVariables.HTTP_MAP["PATH_TRANSLATED"] = RealFile;
        CGIVariables.HTTP_MAP["QUERY_STRING"] = QueryString;
        CGIVariables.HTTP_MAP["REMOTE_ADDRESS"] = inet_ntoa(ClientAddress.sin_addr);
        CGIVariables.HTTP_MAP["REQUEST_METHOD"] = RequestType;
        CGIVariables.HTTP_MAP["SCRIPT_NAME"] = FileRequested;
        CGIVariables.HTTP_MAP["SERVER_PORT"] = IntToString(SWEBSGlobals.Port);
        CGIVariables.HTTP_MAP["SERVER_PROTOCOL"] = HTTPVersion;

        CGIVariables.HTTP_MAP["HTTP_USER_AGENT"] = UserAgent;
        
        // Now process the request
		if (!IsFolder)																// Request was a file
		{
			if (IsBinary == true)
			{
				// The file is a binary file
				Headers = HTTPVersion;												// Send HTTP version
				Headers += " ";											
				Headers += IntToString(Status);										// Send status code
				Headers += " ";			
				Headers += "OK\n";													// Send OK msg
				Headers += "Server: ";
				Headers += SWEBSGlobals.Servername;
				Headers += "\nConnection: ";										// Connection type
				Headers += "close\n";
				Headers += "Date: ";
				Headers += Date;
				Headers += "\nContent-type: ";										// Content type
				if (SWEBSGlobals.MIMETypes[Extension].length() > 0)					// If we know the mime type
				{
					Headers += SWEBSGlobals.MIMETypes[Extension];					// Send it
				}
				else																// Or if we don't know it
				{
					Headers += "image/jpeg";										// Send image/jpg
				}
				Headers += "\nContent-length: ";
				Headers += CalculateSize();
				Headers += "\n\n";													// Double newlines

                SWEBSSocket::Send (SFD, Headers.c_str(), Headers.length());			// Send headers

				// Then, if its a GET of POST request, send the file requested
				if ( !strcmpi(RequestType.c_str(), "GET") || !strcmpi(RequestType.c_str(), "POST") )
					SendBinary();
			}
			else if (IsScript == true)
			{
				// The file is a CGI script
				Headers = HTTPVersion;												// Send HTTP version
				Headers += " ";											
				Headers += IntToString(Status);										// Send status code
				Headers += " ";			
				Headers += "OK\n";													// Send OK msg
				Headers += "Server: ";
				Headers += SWEBSGlobals.Servername;
				Headers += "Date: ";
				Headers += Date;
				Headers += "\nConnection: close ";									// Connection type
				// Note: We do not send the content type OR the double newlines, the
				//  CGI interpreter must do that itself.	

				SWEBSSocket::Send (SFD, Headers.c_str(), Headers.length());			// Send headers

				// Then, if its a GET of POST request, send the file requested
				if ( !strcmpi(RequestType.c_str(), "GET") || !strcmpi(RequestType.c_str(), "POST") )
					SendCGI();
			}
            else if (IsIsapi == true)
            {
                // The file is ISAPI
                // The file is a CGI script
				Headers = HTTPVersion;												// Send HTTP version
				Headers += " ";											
				Headers += IntToString(Status);										// Send status code
				Headers += " ";			
				Headers += "OK\n";													// Send OK msg
				Headers += "Server: ";
				Headers += SWEBSGlobals.Servername;
				Headers += "Date: ";
				Headers += Date;
				Headers += "\nConnection: close ";									// Connection type
				// Note: We do not send the content type OR the double newlines, the
				//  ISAPI interpreter must do that itself.	

				SWEBSSocket::Send (SFD, Headers.c_str(), Headers.length());					// Send headers

				// Then, if its a GET of POST request, send the file requested
				if ( !strcmpi(RequestType.c_str(), "GET") || !strcmpi(RequestType.c_str(), "POST") )
					SendISAPI(this);
            }
			else
			{
				// The file is plain text
				Headers = HTTPVersion;												// Send HTTP version
				Headers += " ";											
				Headers += IntToString(Status);										// Send status code
				Headers += " ";			
				Headers += "OK\n";													// Send OK msg
				Headers += "Server: ";
				Headers += SWEBSGlobals.Servername;
				Headers += "\nConnection: ";										// Connection type
				Headers += "close\n";
				Headers += "Date: ";
				Headers += Date;
				Headers += "\nContent-type: ";										// Content type
                // Make EXTENSION into small letters
                char Ext[20];
                strcpy(Ext, Extension.c_str());
                strlwr(Ext);
				if (SWEBSGlobals.MIMETypes[Ext].length() > 0)	    					// If we know the mime type
				{
					Headers += SWEBSGlobals.MIMETypes[Ext];      						// Send it
				}
				else																// Or if we don't know it
				{
					Headers += "text/plain";										// Send text/plain
				}
				Headers += "\n\n";													// Double newlines
				SWEBSSocket::Send (SFD, Headers.c_str(), Headers.length());					// Send headers
	
				// Then, if its a GET of POST request, send the file requested
				if ( !strcmpi(RequestType.c_str(), "GET") || !strcmpi(RequestType.c_str(), "POST"))
					SendText();
			}
		}
	}
	
    // Write to the stats file
    unsigned long Size = StringToInt(CalculateSize());
    SWEBSGlobals.TotalBytesSent += Size;
    SWEBSGlobals.TotalNumberOfRequests++;
    if (UseVH)
    {
        SWEBSGlobals.Host[HostRequested].BytesSent += Size;
        SWEBSGlobals.Host[HostRequested].NumberOfRequests++;
        SWEBSGlobals.Host[HostRequested].PageRequests[FileRequested] += 1;
    }
    else
    {
        SWEBSGlobals.BytesSent += Size;
        SWEBSGlobals.NumberOfRequests++;
        SWEBSGlobals.PageRequests[FileRequested] += 1;
    }
    
    // If we are logging, write the logs here:
    if (SWEBSGlobals.Logfile.length() > 0)
    {
        LogConnection();
    }
    
    // If the status wasn't 200 to start with, or it was somehow changed along the way:
	if (Status != 200)
	{
		SendError();																// Send the error
		return false;
	}
	return true;																	// No errors. Return true
}

//---------------------------------------------------------------------------------------------
//			Connection::SendText
//---------------------------------------------------------------------------------------------
bool CONNECTION::SendText()
{
	char Buffer[10000];																// Buffer to store data
	string Text;																	// String to store everything
	ifstream hFile (RealFile.c_str());		
	while (!hFile.eof())															// Go until the end
	{
		hFile.getline(Buffer, 10000);												// Read a full line
		Text += Buffer;																// Put it in the string
		Text += "\n";																// Remember to add the \n
	}
	hFile.close();																	// Close
																					// Send the data
    int Y = SWEBSSocket::Send (SFD, Text.c_str(), Text.length());
	if (Y != 0)					
		return true;																// It sent fine
	else
		return false;																// Errors sending
}

//---------------------------------------------------------------------------------------------
//			Connection::SendBinary
//---------------------------------------------------------------------------------------------
bool CONNECTION::SendBinary()
{
	// WOOHOO! Do not lose this function, it took me ages to learn how to send binary files,
	//  and now it finally works!
    char Buffer[10000];
																                    // Open the file as binary
	ifstream hFile (RealFile.c_str(), ios::binary);
	int X = 0;
	while (!hFile.eof())                                                            // Keep reading it in
    {
        hFile.read(Buffer, 10000);
        X += send(SFD, Buffer, hFile.gcount(), 0);			                        // Send data as we read it
    }
    hFile.close();                                                                  // Close  
    
    SWEBSGlobals.TotalBytesSent += X;
    if (UseVH)
    {
        SWEBSGlobals.Host[HostRequested].BytesSent += X;
    }
    else
    {
        SWEBSGlobals.BytesSent += X;
    }
    
    return true;
}

//---------------------------------------------------------------------------------------------
//			Connection::SendCGI
//---------------------------------------------------------------------------------------------
bool CONNECTION::SendCGI()
{
	// This is where we will use the program provided by Volkan Uzun.
    string Message;
    Message = "Content-type: text/html\n\n";
    Message += "<font face='Verdana'><small><center>";
    Message += "<h1>Error 500 - Internal Server Error</h1><br>";
    Message += "At this stage the SWEBS Web Server cannot use CGI due to a problem with creating";
    Message += " named pipes as an NT service. We are working on a way around this issue,";
    Message += " and it will be working for version 1.0 of the server. Thankyou for your";
    Message += " patience and support.";
    Message += "</center></small></font>";

    return true;
}

//---------------------------------------------------------------------------------------------
//			Connection::IndexFolder()
//---------------------------------------------------------------------------------------------
bool CONNECTION::IndexFolder()
{
	// Check if we are allowed
	if (SWEBSGlobals.AllowIndex == false)
	{
		Status = 404;
		return false;
	}
	else
	{	
		// Open and list folder contents
		HANDLE hFind;
		WIN32_FIND_DATA FindData;
		int ErrorCode;
		BOOL Continue = true;

		string FileIndex = RealFile + "\\*.*";										// List all files
		hFind = FindFirstFile(FileIndex.c_str(), &FindData);

		string Text;
		for (int Z = 0; FileRequested[Z] != '\0'; Z++)								// Replace \ with / 
		{
			if (FileRequested[Z] == '\\') FileRequested[Z] = '/';
		}

		//-----------------------------
		if(hFind == INVALID_HANDLE_VALUE)
		{
			Status = 404;
			return false;
		}
		else
		{	
			Headers = HTTPVersion;													// Send HTTP version
			Headers += " 200 OK\n";								
			Headers += "Server: SWS Stovell Web Server 2.0\n";						// Server name
			Headers += "Connection: close";
			Headers += "\nContent-type: text/html";									// Content type
			Headers += "\n\n";														// Double newlines
            SWEBSSocket::Send (SFD, Headers.c_str(), Headers.length());				// Send headers

			// Most of this is all HTML bieng generated													
			Text = "<html>\n<head>\
\n<title>Index of ";
			Text += FileRequested;													// Insert the file requested					
			Text += "</title>\
\n</head>\
\
\n<body>\
\n<!-- Title -->\
\n<p align=\"center\">\
\n  <font face=\"Verdana\" size=\"6\">\
\n    Index of ";
			Text += FileRequested;													// Insert the file requested
			Text += "\n  </font>\
\n</p>\
\n\
\n<hr>\
\n\
\n<center>\
\n<table border=\"0\" width=\"75%\" height=\"6\" cellspacing=\"0\" cellpadding=\"4\">\
\n  <tr>\
\n    <td width=\"40%\" height=\"1\" bgcolor=\"#0080C0\"><p align=\"center\"><strong><small><font\
\n    face=\"Verdana\" color=\"#FFFFFF\">Name</font></small></strong></td>\
\n    <td width=\"23%\" height=\"1\" bgcolor=\"#0080C0\"><p align=\"center\"><font face=\"Verdana\"\
\n    color=\"#FFFFFF\"><strong><small>Size</small></strong></font></td>\
\n    \
\n  </tr>";
																					// Now, print the first entry
			Text += "  <tr>\
\n    <td width=\"40%\" height=\"0\" bgcolor=\"#E2E2E2\"><p align=\"left\"><small><font face=\"Verdana\">";
			Text += "<a href=\"";
			Text += FileRequested;													// Current folder
			Text += "/";
			Text += FindData.cFileName;												// Name of file
			Text += "\">";
			Text += FindData.cFileName;												// Name of file
			Text += "</a></font></small></td>\n";
			Text += "\n    <td width='23%' height='0' bgcolor='#C0C0C0'>\
\n    <p align='center'><small><font face='Verdana'>";
			if (IntToString(FindData.nFileSizeHigh)[0] != '0')
				Text += IntToString(FindData.nFileSizeHigh);
			if (IntToString(FindData.nFileSizeLow)[0] != '0')
				Text += IntToString(FindData.nFileSizeLow);
			Text += "\n</font></small></td>\
\n    </font></small></td>";
            SWEBSSocket::Send(SFD, Text.c_str(), Text.length());
		}

		if (Continue)
		{
			while (FindNextFile(hFind, &FindData))
			{
				Text = "  <tr>\
\n    <td width=\"40%\" height=\"0\" bgcolor=\"#E2E2E2\"><p align=\"left\"><small><font face=\"Verdana\">";
			Text += "<a href=\"";
			if (strcmpi(FileRequested.c_str(), "/"))
				Text += FileRequested;												// Current folder
			Text += "/";
			Text += FindData.cFileName;												// Name of file
			Text += "\">";
			Text += FindData.cFileName;												// Name of file
			Text += "</a></font></small></td>\n";
			Text += "\n    <td width='23%' height='0' bgcolor='#C0C0C0'>\
\n    <p align='center'><small><font face='Verdana'>";
			
			if (IntToString(FindData.nFileSizeHigh)[0] != '0')
				Text += IntToString(FindData.nFileSizeHigh);
			if (IntToString(FindData.nFileSizeLow)[0] != '0')
				Text += IntToString(FindData.nFileSizeLow);
			Text += "\n</font></small></td>\
\n    </font></small></td>";
            SWEBSSocket::Send(SFD, Text.c_str(), Text.length());
			}

			ErrorCode = GetLastError();

			if (ErrorCode == ERROR_NO_MORE_FILES)
			{
				Text = "</tr>\
\n</table>\
\n</div>\
\n\
\n<hr>\
\n\
\n<p align='center'><small><small><font face='Verdana'>Index produced automatically by <a\
\nhref='http://swebs.sourceforge.net'>SWS Web Server</a></font></small></small></p>\
\n</body>\
\n</html>";
                SWEBSSocket::Send(SFD, Text.c_str(), Text.length());
			}

        FindClose(hFind);
       
		}
	return true;
	}
}

//---------------------------------------------------------------------------------------------
//			Connection::SendError()
//			Sends the appropriate error.
//---------------------------------------------------------------------------------------------
bool CONNECTION::SendError()
{	
	RealFile = SWEBSGlobals.ErrorDirectory;												// Where the error files are kept
	RealFile += '\\';																// Add a backslash to be safe
	RealFile += IntToString(Status);												// Error code
	RealFile += ".html";															// Extension

	char Buffer[10000];																// Buffer to store data

	string ErrorPage;

	ifstream hFile (RealFile.c_str());	
	if (hFile)																		// The file exists
	{
		ErrorPage = "HTTP/1.1 ";													// Since we have an HTML file to send them, just send a 200
		ErrorPage += "200";
		ErrorPage += ' ';
		ErrorPage += "OK";
		ErrorPage += "\nContent-type: text/html\n\n";
		while (!hFile.eof())														// Go until the end
		{
			hFile.getline(Buffer, 10000);											// Read a full line
			ErrorPage += Buffer;													// Put it in the string
			ErrorPage += "\n";														// Remember to add the \n
		}
		hFile.close();																// Close
	}
	
	else																			// There was no custom error page for this code
	{					
		ErrorPage = "HTTP/1.1 ";								
		ErrorPage += IntToString(Status);											// Send the appropriate status code
		ErrorPage += ' ';
		ErrorPage += SWEBSGlobals.ErrorCodes[Status];
		ErrorPage += "\nContent-type: text/html\n\n";

		ErrorPage += "<html><body><center><b>";										// Send a basic and boring looking error page
		ErrorPage += IntToString(Status);
		ErrorPage += " ";
		ErrorPage += SWEBSGlobals.ErrorCodes[Status];
		ErrorPage += "</b></body></html>";
	}
    int Y = SWEBSSocket::Send (SFD, ErrorPage.c_str(), ErrorPage.length());

	if (Y != 0)					
		return true;																// It sent fine
	else
		return false;																// Errors sending
	
	return false;
}

//---------------------------------------------------------------------------------------------
//			Connection::LogText()
//			Logs the string passed as an argument. This is just a temporary log, 
//			used for testing. A propper logging function must be written to replace this.
//---------------------------------------------------------------------------------------------

//---------------------------------------------------------------------------------------------
//			Connection::CalculateSize()
//			Calculates and returns the size of the file requested.
//---------------------------------------------------------------------------------------------
string CONNECTION::CalculateSize()
{
	string ReturnValue;
	unsigned long Size = 0;

	HANDLE hFile = CreateFile(
				RealFile.c_str(), 
				GENERIC_READ,														// Open for reading 
                FILE_SHARE_READ,													// Share for reading 
                NULL,																// No security 
                OPEN_EXISTING,														// Existing file only 
                FILE_ATTRIBUTE_NORMAL,												// Normal file 
                NULL);											

	if (hFile != INVALID_HANDLE_VALUE)												// If the file was opened
	{
		Size = GetFileSize(hFile, NULL);											// Get the size
	}
	
	CloseHandle(hFile);
	ReturnValue = IntToString(Size);												// Convert it to a printable string
	return ReturnValue;
}


//---------------------------------------------------------------------------------------------
//			Connection::ModifiedSince()
//			Checks if the file was modified since the given date
//---------------------------------------------------------------------------------------------
bool CONNECTION::ModifiedSince(string Date)
{
	//----------------------------------------------------------
	string Temp;
	time_t now;
	struct tm * tm_now = NULL;
	struct tm * tm_date = NULL;

	// Get the time the file was last modified:
	SYSTEMTIME myTime;																// Windows time structure
	HANDLE hFind;
    WIN32_FIND_DATA FindData;

	now = time ( NULL );															// Get the current time
	tm_date = localtime ( &now );
	
	hFind = FindFirstFile(RealFile.c_str(), &FindData);								// Open the file
	FileTimeToSystemTime(&FindData.ftLastWriteTime, &myTime);						// Get its time
	
	tm_now->tm_hour = myTime.wHour + 1;												// Copy the time settings
	tm_now->tm_min = myTime.wMinute;
	tm_now->tm_mday = myTime.wDay;
	tm_now->tm_mon = myTime.wMonth - 1;
	tm_now->tm_sec = myTime.wSecond;
	tm_now->tm_year = myTime.wYear;
	tm_now->tm_year -= 1900;

	//----------------------------------------------------------
	int Day;
	string Month;
	int Year;
	int Hour;
	int Minute;
	int Second;
	
	int Type = 0;

	istringstream DateString1(Date);
	DateString1 >> Temp;

	if (Temp[3] == ',')
	{
		Type = 1;
		//------------------------------------------------------
		for (int X = 0; X < Date.length(); X++)
		{
			if ( Date[X] == ':' )
				Date[X] = ' ';
		}

		istringstream DateString(Date);
		DateString >> Temp;
		//------------------------------------------------------
		DateString >> Day;
		DateString >> Month;
		DateString >> Year;
		DateString >> Hour;
		DateString >> Minute;
		DateString >> Second;
	}
	else if (Temp.length() > 3)
	{
		Type = 2;
		//------------------------------------------------------
		for (int X = 0; X < Date.length(); X++)
		{
			if ( Date[X] == ':' )
				Date[X] = ' ';
		}
		for (int Z = 0; Z < Date.length(); Z++)
		{
			if ( Date[Z] == '-' )
				Date[Z] = ' ';
		}
		istringstream DateString(Date);
		DateString >> Temp;
		//------------------------------------------------------
		DateString >> Day;
		DateString >> Month;
		DateString >> Year;
		Year += 2000;
		if (Year > 2050)
			return true;
		DateString >> Hour;
		DateString >> Minute;
		DateString >> Second;
	}
	else if (Temp[3] != ',')
	{
		Type = 3;
		//------------------------------------------------------
		for (int X = 0; X < Date.length(); X++)
		{
			if ( Date[X] == ':' )
				Date[X] = ' ';
		}

		istringstream DateString(Date);
		DateString >> Temp;
		//------------------------------------------------------
		
		DateString >> Month;
		DateString >> Day;
		DateString >> Hour;
		DateString >> Minute;
		DateString >> Second;
		DateString >> Year;
	}

	//----------------------------------------------------------
	tm_date->tm_hour = Hour;
	tm_date->tm_mday = Day;
	tm_date->tm_min = Minute;
	tm_date->tm_mon = CalcMonth(Month);
	tm_date->tm_sec = Second;
	tm_date->tm_year = Year - 1900;
	
	//----------------------------------------------------------
	
	int Y = difftime(mktime(tm_date), mktime(tm_now));
	delete tm_now;
	delete tm_date;
	if (Y < 0)
	{
		return true;
	}
	else return false;
}


//---------------------------------------------------------------------------------------------
//			Connection::UnModifiedSince()
//			Returns a negative of ModifiedSince()
//---------------------------------------------------------------------------------------------
bool CONNECTION::UnModifiedSince(string Date)
{
	bool X = ModifiedSince(Date);
	if (X)
	{
		return false;
	}
	else 
	{	
		return true;
	}
}



//----------------------------------------------------------------------------------
//      LogText() - writes a single string to the log file
//----------------------------------------------------------------------------------
bool CONNECTION::LogText(string Text)
{
	FILE* log = NULL;
    if (UseVH)                                                                      // Check if this connection uses a VH
    {
	    log = fopen(ThisHost->Logfile.c_str(), "a+");                                       // It does, so open its log file
    }
    else 
    {
        log = fopen(SWEBSGlobals.Logfile.c_str(), "a+");                                        // It doesn't, so use the default one
    }
	if (log == NULL)
    {
        return false;
    }
	fprintf(log, "%s", Text.c_str());                                               // Write the string
	fclose(log);                                                                    // Close it
	return true;
}

//----------------------------------------------------------------------------------
//      CONNECTION::LogConnection() - logs the current connection
//----------------------------------------------------------------------------------
bool CONNECTION::LogConnection()
{
    time_t now;
	struct tm *tm_now = NULL;
	char buff[BUFSIZ];

	now = time ( NULL );
	tm_now = gmtime(&now);
    
    LogText(inet_ntoa(ClientAddress.sin_addr));                                 // IP Address
    LogText(" - - [");                                                          // Normally the name goes here
    strftime ( buff, sizeof buff, "%d/%b/%Y:%H:%M:%S +0000", tm_now );          // Get the time (GMT)
	LogText(buff);
    LogText("] \"");
    LogText(RequestType);                                                       // Request line
    LogText(" ");
    LogText(FileRequested);
    LogText(" ");
    LogText(HTTPVersion);
    LogText("\" ");
    LogText(IntToString(Status));                                               // Status code
    LogText(" ");
    LogText( CalculateSize() );                                                 // Size
    LogText(" \"");
    LogText(Referer);                                                           // Referer
    LogText("\" \"");
    LogText(UserAgent);                                                         // User agent
    LogText("\"\n");
      
    return true;
}

//----------------------------------------------------------------------------------
//          SendISAPI()
//----------------------------------------------------------------------------------
bool SendISAPI(CONNECTION * Connection)
{
    HTTPEXTENSIONPROC pHttpExtensionProc;
    GETEXTENSIONVERSION pGetExtensionVersion;

    // First, check our DLL has been loaded
    if ( SWEBSGlobals.IsapiDLL[Connection->Extension] == NULL )
        return false;                                                               // The DLL has not been loaded. Abort
    
    // Set up the extension control block
    EXTENSION_CONTROL_BLOCK ECB;
    ECB.cbAvailable = 3;                                                            // No bytes avaliable yet, please call ReadClient()
    ECB.cbTotalBytes = 2; //Connection->PostData.length();                                           // Ammount of post data we have
    ECB.ConnID = (HCONN)(Connection);                                                        // This is a uniqe number, in our case we can use the SFD
    ECB.dwHttpStatusCode = Connection->Status;                                                  // Status code
    ECB.lpszContentType = (char *)(Connection->CGIVariables.HTTP_MAP["CONTENT_TYPE"].c_str());                                // The rest of this is just some header mapping
    ECB.lpszMethod = (char *)(Connection->CGIVariables.HTTP_MAP["REQUEST_METHOD"].c_str());
    ECB.lpszPathInfo = (char *)(Connection->CGIVariables.HTTP_MAP["PATH_INFO"].c_str());
    ECB.lpszPathTranslated = (char *)(Connection->CGIVariables.HTTP_MAP["PATH_TRANSLATED"].c_str());
    ECB.lpszQueryString = (char *)(Connection->CGIVariables.HTTP_MAP["QUERY_STRING"].c_str());                                
    
    
    ECB.GetServerVariable = GetServerVariableExport;
    ECB.ReadClient = ReadClientExport;
    ECB.WriteClient = WriteClientExport;
    ECB.ServerSupportFunction = ServerSupportFunctionExport;

    HSE_VERSION_INFO HSE;                                                           // Version info stuff. Who cares?!

    // Find the functions
    pHttpExtensionProc = (HTTPEXTENSIONPROC) GetProcAddress(SWEBSGlobals.IsapiDLL[Connection->Extension], "HttpExtensionProc");
    pGetExtensionVersion = (GETEXTENSIONVERSION) GetProcAddress(SWEBSGlobals.IsapiDLL[Connection->Extension], "GetExtensionVersion");

    pGetExtensionVersion(&HSE);                                                     // Make a call to the extension version
    pHttpExtensionProc(&ECB);                                                       // Make a call to the extension proc
    return true;
}

//---------------------------------------------------------------------------------------------







