#include <iostream>
#include <windows.h>
#include <string>
#include <sstream>
#include "connection.hpp"

using namespace std;

#ifndef STATSHPP
#define STATSHPP
//----------------------------------------------------------------------------------------------------
//			VHSTATS Class
//----------------------------------------------------------------------------------------------------
class VHSTATS
{
public:
    VHSTATS();
    VHSTATS(VIRTUALHOST);                                                           // Stats on a per-virtualhost basis
    friend bool operator>(const VHSTATS lhs, const VHSTATS rhs);                                        // Needed for use in map
    friend bool operator<(const VHSTATS lhs, const VHSTATS rhs);
    map <string, int> PageRequests;                                                 // Page requests (per VH)
    unsigned long NumberOfRequests;                                                 // Number of requests
    unsigned long BytesSent;                                                        // Bytes sent for this virtual host
    VIRTUALHOST ThisVH;                                                             // This virtual host
};

//----------------------------------------------------------------------------------------------------
//			STATS Class
//----------------------------------------------------------------------------------------------------
class STATS
{
  public:
    STATS();                                                                        // Constructor - must load the stats
    bool WriteStatsFile();                                                          // Writes all statistics to the stats file
    
    unsigned long NumberOfRequests;                                                 // Number of connections served by the server
    unsigned long TotalNumberOfRequests;                                            // Total number of connections served
    unsigned long BytesSent;                                                        // Total number of bytes served
    unsigned long TotalBytesSent;                                                   // Total number of bytes sent for VH's and normal requests
    string LastRestart;                                                             // Last time the server was restarted
    map <string, int> PageRequests;                                                 // Page requests
    map <VIRTUALHOST, VHSTATS> VirtualHosts;                                        // To keep stats on a per-virtualhost basis 
    map <int, string> PageRequestIndex;                                             // Index of pages requested
};

//----------------------------------------------------------------------------------------------------
//			Externals
//----------------------------------------------------------------------------------------------------
extern STATS SWEBSStats;                                                            // Make statistics avaliable to everyone
extern DWORD WINAPI HandleStatsFile(LPVOID lpParam );                               // Writes information to the stats file. 
                                                                                    // The argument is the CONNECTION class used.
#endif
