#pragma warning(disable:4786)
#include "headermap.hpp"

SWEBS_HM SWEBS_hm_ACCEPT;
SWEBS_HM SWEBS_hm_ACCEPT_CHARSET;
SWEBS_HM SWEBS_hm_ACCEPT_ENCODING;
SWEBS_HM SWEBS_hm_ACCEPT_LANGUAGE;
SWEBS_HM SWEBS_hm_AGE;
SWEBS_HM SWEBS_hm_AUTHORIZATION;
SWEBS_HM SWEBS_hm_CONNECTION;
SWEBS_HM SWEBS_hm_CONTENT_ENCODING;
SWEBS_HM SWEBS_hm_CONTENT_LANGUAGE;
SWEBS_HM SWEBS_hm_CONTENT_LENGTH;
SWEBS_HM SWEBS_hm_CONTENT_LOCATION;
SWEBS_HM SWEBS_hm_CONTENT_TYPE;
SWEBS_HM SWEBS_hm_CONTENT_FROM;
SWEBS_HM SWEBS_hm_CONTENT_HOST;
SWEBS_HM SWEBS_hm_IF_MODIFIED_SINCE;
SWEBS_HM SWEBS_hm_LAST_MODIFIED;
SWEBS_HM SWEBS_hm_REFERER;
SWEBS_HM SWEBS_hm_USER_AGENT;
SWEBS_HM hm_TEST;																	// This one is used for testing

map <string, SWEBS_HM> HeaderMap;
HINSTANCE gSWEBS_headermapDLL;

//---------------------------------------------------------------------------------------------
//			HeaderMapInit()
//---------------------------------------------------------------------------------------------
bool HeaderMapInit()
{
	gSWEBS_headermapDLL = LoadLibrary("SWEBS_headermap.dll");						// Load the library
	if (gSWEBS_headermapDLL == NULL)												// The library could not be loaded
	{					
		return false;
	}
	GetProcAddress(gSWEBS_headermapDLL, "SWEBS_hm_ACCEPT");							// 
	return true;	//!! added	
}
