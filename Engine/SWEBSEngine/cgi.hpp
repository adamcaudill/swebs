//---------------------------------------------------------------------------------------------
//          CGI.hpp
//          ----------------
//          This file contains information used by the CONNECTION class in CGI matters. The
//          variables here can be used by SSI, FastCGI, ISAPI, NSAPI, and any other form of CGI
//          that the SWEBS Web Server supports.
//          
//---------------------------------------------------------------------------------------------
#ifndef CGIHPP
#define CGIHPP 1
#include <string>

using namespace std;

//---------------------------------------------------------------------------------------------
//          Non-request specific CGI class
//          ------------------------------------
//          These values are set when the server runs and are not changed, and have been 
//          defined in the CGI Specification at http://hoohoo.ncsa.uiuc.edu/cgi/env.html.
//---------------------------------------------------------------------------------------------
class NON_REQUEST_SPECIFIC_CGI
{
  public:
    string SERVER_SOFTWARE;                                                         // Name and version of this software
    string SERVER_NAME;                                                             // Hostname or IP address of server
    string GATEWAY_INTERFACE;                                                       // CGI Revision of this server
};

//---------------------------------------------------------------------------------------------
//          CGI Environment Varliables structure
//          ------------------------------------
//          These values are set per-request, and have the meaning definied in the CGI
//          specification at http://hoohoo.ncsa.uiuc.edu/cgi/env.html. 
//---------------------------------------------------------------------------------------------
class REQUEST_SPECIFIC_CGI
{
  public:
    string SERVER_PROTOCOL;                                                         // HTTP Version
    string SERVER_PORT;                                                             // Port for this request
    string REQUEST_METHOD;                                                          // GET, HEAD, POST etc
    string PATH_INFO;                                                               // Extra information at end of file name
    string PATH_TRANSLATED;                                                         // Real file
    string SCRIPT_NAME;                                                             // Virtual name of this file
    string QUERY_STRING;                                                            // Query string sent by client
    string REMOTE_HOST;                                                             // Hostname of client computer
    string REMOTE_ADDRESS;                                                          // IP Address of client
    string AUTH_TYPE;                                                               // Authetithication method used
    string REMOTE_USER;                                                             // Username used
    string REMOTE_IDENT;                                                            // Remote identity (if supported)
    string CONTENT_TYPE;                                                            // Content length of POST/PUT data
    string CONTENT_LENGTH;                                                          // POST/PUT request data
    
    // HTTP Variables
    string HTTP_USER_AGENT;                                                         // User-agent data
    string HTTP_ACCEPT;                                                             // Accepts
};

//---------------------------------------------------------------------------------------------
#endif