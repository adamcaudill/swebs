#pragma warning(disable:4786)
#include "headermap.hpp"

//---------------------------------------------------------------------------------------------
//			Function Definitions
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_ACCEPT(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_ACCEPT_CHARSET(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_ACCEPT_ENCODING(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_ACCEPT_LANGUAGE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_AGE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_AUTHORIZATION(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONNECTION(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONTENT_ENCODING(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONTENT_LANGUAGE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONTENT_LENGTH(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONTENT_LOCATION(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_CONTENT_TYPE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_FROM(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_IF_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_IF_NOT_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_LAST_MODIFIED(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_HOST(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_REFERER(istringstream &IS, CONNECTION * Connection);
bool SWEBS_hm_USER_AGENT(istringstream &IS, CONNECTION * Connection);

typedef bool (*SWEBS_HM)(istringstream &IS, CONNECTION * Connection);

HINSTANCE gSWEBS_headermapDLL;
map <string, SWEBS_HM>HeaderMap;

//---------------------------------------------------------------------------------------------
//			HeaderMapInit()
//---------------------------------------------------------------------------------------------
bool HeaderMapInit()
{
	// Map our supported headers
	HeaderMap["Host:"] = SWEBS_hm_HOST;
	HeaderMap["User-Agent:"] = SWEBS_hm_USER_AGENT;
	HeaderMap["From:"] = SWEBS_hm_FROM;
	HeaderMap["Connection:"] = SWEBS_hm_CONNECTION;
	HeaderMap["User-Agent:"] = SWEBS_hm_USER_AGENT;
	HeaderMap["If-Modified-Since:"] = SWEBS_hm_IF_MODIFIED_SINCE;
	HeaderMap["If-Unmodified-Since:"] = SWEBS_hm_IF_NOT_MODIFIED_SINCE;
	return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_HOST
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_HOST(istringstream &IS, CONNECTION * Connection)
{
	// The server found the header "Host:", which means the next word is the host requested
	string Word;
	IS >> Word;
	Connection->HostRequested = Word;
	return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_USER_AGENT
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_USER_AGENT(istringstream &IS, CONNECTION * Connection)
{
	// While theres no colons in the word, the word should be added to UserAgent
	string Word;
	IS >> Word;
	while ( !strstr(Word.c_str(), ":") )									
	{
		Connection->UserAgent += Word;
		Connection->UserAgent += ' ';
		IS >> Word;
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
	Connection->Connection = Word;
	return true;
}

//---------------------------------------------------------------------------------------------
//			SWEBS_hm_IF_MODIFIED_SINCE
//---------------------------------------------------------------------------------------------
bool SWEBS_hm_IF_MODIFIED_SINCE(istringstream &IS, CONNECTION * Connection)
{
	// If-Modified-Since
	// Unfortunately, there are 3 ways we can be given the time in an HTTP request, so we must handle all 3
	string Word;
	IS>>Word;
	Connection->UseModDate = true;
	if (Word[3] == ',')
	{
		// Its the first type
		Connection->ModifiedSinceStr = Word + ' ';									// We have the day and 5 more words
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
		Connection->ModifiedSinceStr = Word + ' ';									// We have the day and 3 more words
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
		Connection->ModifiedSinceStr = Word + ' ';									// We have the day and 4 more words
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
		Connection->UnModifiedSinceStr = Word + ' ';								// We have the day and 5 more words
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
		Connection->UnModifiedSinceStr = Word + ' ';								// We have the day and 3 more words
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
		Connection->UnModifiedSinceStr = Word + ' ';								// We have the day and 4 more words
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
