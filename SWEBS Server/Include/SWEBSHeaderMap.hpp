#ifndef SWEBSHEADERMAPHPP
#define SWEBSHEADERMAPHPP 1
//---------------------------------------------------------------------------------------------
/*
			SWEBSHEADERMAP.HPP
			--------------
			This file contains functions used for header mapping. Header mapping
			is the process of mapping a header such as "Host:" to a function to
			handle that header. These functions change the Connection argument 
			depending on what headers are sent. Using a map its easy to disregard
			headers we don't understand/support.
			
			Header mapping can't be used until the function HeaderMapInit has been
			called, which is usually done by the main server program.
			
			The functions for this are in the library SWEBSHeadermap.lib 
*/
//---------------------------------------------------------------------------------------------
#include <sstream>
#include "../Include/SWEBSConnection.hpp"

using namespace std;

//---------------------------------------------------------------------------------------------
//			Function Definitions
//---------------------------------------------------------------------------------------------
extern bool SWEBS_hm_ACCEPT_CHARSET(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_ACCEPT_ENCODING(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_ACCEPT_LANGUAGE(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_AUTHORIZATION(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_CONNECTION(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_CONTENT_ENCODING(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_CONTENT_LANGUAGE(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_CONTENT_LENGTH(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_CONTENT_TYPE(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_FROM(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_IF_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_IF_NOT_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_HOST(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_REFERER(istringstream &IS, CONNECTION * Connection);
extern bool SWEBS_hm_USER_AGENT(istringstream &IS, CONNECTION * Connection);

typedef bool (*SWEBS_HM)(istringstream &IS, CONNECTION * Connection);

extern bool HeaderMapInit();

//---------------------------------------------------------------------------------------------
//			MAP
//---------------------------------------------------------------------------------------------
extern HINSTANCE gSWEBS_headermapDLL;
extern map <string, SWEBS_HM>HeaderMap;

//---------------------------------------------------------------------------------------------
#endif