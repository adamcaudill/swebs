//--------------------------------------------------------------------------------
//			SWEBSUTILITIES.cpp
//			-------------------
//			This source is a bunch of functions that dont belong anywhere else,
//			and are used thoughout the server. They perform trivial yet
//			necessary jobs. They have been compiled into the library 
//			SWEBSUtilities.lib
//  
//--------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include <string>
#include <sstream>
#include "../Include/SWEBSUtilities.hpp"
using namespace std;

//----------------------------------------------------------------------------------------------------
//			IntToString();
//----------------------------------------------------------------------------------------------------
string IntToString(int num)
{
  ostringstream myStream; 															// Creates an ostringstream object
  myStream << num << flush;
  return(myStream.str()); 															// Returns the string form of the stringstream object
}

//----------------------------------------------------------------------------------------------------
//			StringToInt();
//----------------------------------------------------------------------------------------------------
int StringToInt(string str)
{
   std::istringstream is(str);
   int i;    
   
   is >> i;    
   return i;
}

//----------------------------------------------------------------------------------------------------
//			CalcMonth() - returns the integer description of a month from its string, ie, "Feb" returns 2
//----------------------------------------------------------------------------------------------------
int CalcMonth(string Month)
{
	if ( !strcmpi(Month.c_str(), "Jan")   )
		return 0;
	else if ( !strcmpi(Month.c_str(), "Feb")   )
		return 1;
	else if ( !strcmpi(Month.c_str(), "Mar")   )
		return 2;
	else if ( !strcmpi(Month.c_str(), "Apr")   )
		return 3;
	else if ( !strcmpi(Month.c_str(), "May")   )
		return 4;
	else if ( !strcmpi(Month.c_str(), "Jun")   )
		return 5;
	else if ( !strcmpi(Month.c_str(), "Jul")   )
		return 6;
	else if ( !strcmpi(Month.c_str(), "Aug")   )
		return 7;
	else if ( !strcmpi(Month.c_str(), "Sep")   )
		return 8;
	else if ( !strcmpi(Month.c_str(), "Oct")   )
		return 9;
	else if ( !strcmpi(Month.c_str(), "Nov")   )
		return 10;
	else if ( !strcmpi(Month.c_str(), "Dec")   )
		return 11;
	else return 0;
}