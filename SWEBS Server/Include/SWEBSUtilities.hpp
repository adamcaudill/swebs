#ifndef SWEBSUTILITIESHPP
#define SWEBSUTILITIESHPP 1
//--------------------------------------------------------------------------------
//			SWEBSUTILITIES.hpp
//			-------------------
//			This header is a bunch of functions that dont belong anywhere else,
//			and are used thoughout the server. They perform trivial yet
//			necessary jobs. They have been compiled into the library 
//			SWEBSUtilities.lib
//  
//--------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include <string>

using namespace std;

//----------------------------------------------------------------------------------
//          Functions
//----------------------------------------------------------------------------------
int StringToInt(string);														    // Converts a string to an integer
string IntToString(int num);													    // Converts an integer to a string
int CalcMonth(string Month);													    // Returns the number of a month from the string
string DecodeURL(string URL);					   								    // Decodes a URL-Encoded string

//----------------------------------------------------------------------------------
#endif