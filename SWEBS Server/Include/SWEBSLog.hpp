#ifndef SWEBSLOGHPP
#define SWEBSLOGHPP 1
//--------------------------------------------------------------------------------
//			SWEBSLOG.hpp
//			-------------------
//			This header contains function declarations used by SWEBS for logging.
//			The function definitions in this file can be found in SWEBSLog.lib. 
//
//--------------------------------------------------------------------------------																				
//			INCLUDES
//--------------------------------------------------------------------------------
#include "includes.hpp"

using namespace std;
//--------------------------------------------------------------------------------
extern bool WriteLogFile(string File, CONNECTION Connection);					  // Writes the log file using the variables in the CONNECTION class

//--------------------------------------------------------------------------------
#endif