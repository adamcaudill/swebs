#ifndef HEADERMAPHPP
#define HEADERMAPHPP 1
//---------------------------------------------------------------------------------------------
/*
			HEADERMAP.HPP
			--------------
			This file contains functions for mapping HTTP Headers
			with functions to handle the values for those headers.

			There is a map called HeaderMap, which maps a string
			such as "Host:" to a function that handles the next word.
			The function returns a bool, and takes 2 arguments.
			The first is the address of an ISTRINGSTREAM type 
			that holds the current word from the CONNECTION 
			class. The second is the address of the calling 
			CONNECTION class, so that values can be added directly.

			Functions for handling the headers are kept in a DLL, 
			and must be loaded in on startup. This is done by
			the ServiceMain() function at the beginning of startup
			when it calls HeaderMapInit(). 
			
			HeaderMap returns false if it cannot load the DLL 
			"SWEBS_headermap.dll", which will result in the server
			shutting down.

			Inside the SWEBS_headermap.dll file are a list of functions
			all look like SWEBS_hm_HEADER(), where HEADER is the
			header that the function will handle. Eg: SWEBS_hm_HOST()
			handles the host header.

			It is up to the header handling function to be sure the
			ISTRINGSTREAM passed by the calling CONNECTION class 
			is up to the next header, before it finishes.
*/
//---------------------------------------------------------------------------------------------
#include "connection.hpp"
#include <sstream>

using namespace std;
//---------------------------------------------------------------------------------------------
//			FUNCTION POINTERS
//---------------------------------------------------------------------------------------------

typedef bool (*SWEBS_HM)(istringstream &IS,				
							  CONNECTION * Connection);
extern SWEBS_HM SWEBS_hm_ACCEPT;
extern SWEBS_HM SWEBS_hm_ACCEPT_CHARSET;
extern SWEBS_HM SWEBS_hm_ACCEPT_ENCODING;
extern SWEBS_HM SWEBS_hm_ACCEPT_LANGUAGE;
extern SWEBS_HM SWEBS_hm_AGE;
extern SWEBS_HM SWEBS_hm_AUTHORIZATION;
extern SWEBS_HM SWEBS_hm_CONNECTION;
extern SWEBS_HM SWEBS_hm_CONTENT_ENCODING;
extern SWEBS_HM SWEBS_hm_CONTENT_LANGUAGE;
extern SWEBS_HM SWEBS_hm_CONTENT_LENGTH;
extern SWEBS_HM SWEBS_hm_CONTENT_LOCATION;
extern SWEBS_HM SWEBS_hm_CONTENT_TYPE;
extern SWEBS_HM SWEBS_hm_CONTENT_FROM;
extern SWEBS_HM SWEBS_hm_CONTENT_HOST;
extern SWEBS_HM SWEBS_hm_IF_MODIFIED_SINCE;
extern SWEBS_HM SWEBS_hm_LAST_MODIFIED;
extern SWEBS_HM SWEBS_hm_REFERER;
extern SWEBS_HM SWEBS_hm_USER_AGENT;
extern SWEBS_HM hm_TEST;												// This one is used for testing

extern map <string, SWEBS_HM> HeaderMap;
extern HINSTANCE gSWEBS_headermapDLL;

//---------------------------------------------------------------------------------------------
//			MAP
//---------------------------------------------------------------------------------------------

//---------------------------------------------------------------------------------------------
#endif