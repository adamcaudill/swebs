//---------------------------------------------------------------------------------------------
/*
			SWEBSHEADERMAP.CPP
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

			It is up to the header handling function to be sure the
			ISTRINGSTREAM passed by the calling CONNECTION class 
			is up to the next header, before it finishes.
*/
//---------------------------------------------------------------------------------------------
#include "../Include/SWEBSHeadermap.hpp"
#include "../Include/SWEBSUtilities.hpp"
#include <sstream>
#include <string>

using namespace std;

//---------------------------------------------------------------------------------------------
//			Function Definitions
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_ACCEPT_CHARSET(istringstream &IS, CONNECTION * Connection);           	// No
bool SWEBS_hm_ACCEPT_ENCODING(istringstream &IS, CONNECTION * Connection);              // No
bool SWEBS_hm_ACCEPT_LANGUAGE(istringstream &IS, CONNECTION * Connection);              // No
bool SWEBS_hm_AUTHORIZATION(istringstream &IS, CONNECTION * Connection);                // No
bool SWEBS_hm_CONNECTION(istringstream &IS, CONNECTION * Connection);                   // Yes
bool SWEBS_hm_CONTENT_ENCODING(istringstream &IS, CONNECTION * Connection);             // No
bool SWEBS_hm_CONTENT_LANGUAGE(istringstream &IS, CONNECTION * Connection);             // No
bool SWEBS_hm_CONTENT_LENGTH(istringstream &IS, CONNECTION * Connection);               // No
bool SWEBS_hm_CONTENT_TYPE(istringstream &IS, CONNECTION * Connection);                 // Yes
bool SWEBS_hm_COOKIE(istringstream &IS, CONNECTION * Connection);                 // Yes
bool SWEBS_hm_FROM(istringstream &IS, CONNECTION * Connection);                         // Yes
bool SWEBS_hm_IF_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);            // Yes
bool SWEBS_hm_IF_NOT_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);        // Yes
bool SWEBS_hm_HOST(istringstream &IS, CONNECTION * Connection);                         // Yes
bool SWEBS_hm_REFERER(istringstream &IS, CONNECTION * Connection);                      // Yes
bool SWEBS_hm_USER_AGENT(istringstream &IS, CONNECTION * Connection);                   // No

typedef bool (*SWEBS_HM)(istringstream &IS, CONNECTION * Connection);

bool HeaderMapInit();

//---------------------------------------------------------------------------------------------
//			MAP
//---------------------------------------------------------------------------------------------
HINSTANCE gSWEBS_headermapDLL;
map <string, SWEBS_HM>HeaderMap;

//---------------------------------------------------------------------------------------------
//			HeaderMapInit()
//---------------------------------------------------------------------------------------------
bool HeaderMapInit()
{
    HeaderMap["Host:"] = SWEBS_hm_HOST;
    HeaderMap["If-Modified-Since:"] = SWEBS_hm_IF_MODIFIED_SINCE;
    HeaderMap["If-Unmodified-Since"] = SWEBS_hm_IF_NOT_MODIFIED_SINCE;
    HeaderMap["From:"] = SWEBS_hm_FROM;
    HeaderMap["Connection:"] = SWEBS_hm_CONNECTION;
    HeaderMap["Referer:"] = SWEBS_hm_REFERER;
    HeaderMap["User-Agent:"] = SWEBS_hm_USER_AGENT;
    HeaderMap["Content-Type:"] = SWEBS_hm_CONTENT_TYPE;
    HeaderMap["Content-Length:"] = SWEBS_hm_CONTENT_LENGTH;
    HeaderMap["Cookie:"] = SWEBS_hm_COOKIE;
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_HOST
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_HOST(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->HostRequested = Word;
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_IF_MODIFIED_SINCE
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_IF_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    
    Connection->UseModDate = true;
	if (Word[3] == ',')
	{
		// Its the first type
		Connection->ModifiedSinceStr = Word + ' ';					                        // We have the day and 5 more words
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
        IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
	}
	else if (Word.length() > 3)
	{
		// Its the second type
		Connection->ModifiedSinceStr = Word + ' ';					                        // We have the day and 3 more words
		IS >> Word;
	    Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
        Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;			
        Connection->ModifiedSinceStr += Word + ' ';
	}
	else if (Word[3] != ',')
	{
		// Its the third date type
		Connection->ModifiedSinceStr = Word + ' ';					                        // We have the day and 4 more words
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->ModifiedSinceStr += Word + ' ';
	}
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_IF_NOT_MODIFIED_SINCE
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_IF_NOT_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
	Connection->UseUnModDate = true;
	if (Word[3] == ',')
	{
		// Its the first type
        Connection->UnModifiedSinceStr = Word + ' ';				                        // We have the day and 5 more words
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
	}
	else if (Word.length() > 3)
	{
		// Its the second type
		Connection->UnModifiedSinceStr = Word + ' ';				                        // We have the day and 3 more words
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
	}
	else if (Word[3] != ',')
	{
		// Its the third date type
		Connection->UnModifiedSinceStr = Word + ' ';				                        // We have the day and 4 more words
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';			
        IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
		IS >> Word;
		Connection->UnModifiedSinceStr += Word + ' ';
	}
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_FROM
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_FROM(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->From = Word;
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_CONNECTION
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_CONNECTION(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->ConnectionType = Word;
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_REFERER
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_REFERER(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->Referer = Word;
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_USER_AGENT
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_USER_AGENT(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    char String[256];
    IS.getline(String, 256);                                                        // Grab this line
    Connection->UserAgent = String;                                                 // That line was the user-agent
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_CONTENT_TYPE
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_CONTENT_TYPE(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->ContentType = Word;                                                 // That word is the content type
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_CONTENT_LENGTH
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_CONTENT_LENGTH(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    IS >> Word;
    Connection->ContentLength = StringToInt(Word);                                  // That word is the content length
    return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_COOKIE
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_COOKIE(istringstream &IS, CONNECTION * Connection)
{
    string Word;
    char String[1024];
    IS.getline(String, 1024);                                                        // Grab this line
    Connection->Cookie = String;                                                    // That line was the user-agent
    return true;
}
//---------------------------------------------------------------------------------------------

