#include <iostream>
#include <windows.h>
#include <string>
#include <sstream>
#include "connection.hpp"
#include "stats.hpp"

using namespace std;

//----------------------------------------------------------------------------------------------------
//          Globals
//----------------------------------------------------------------------------------------------------
const int TimeBetweenWrites = 300000;                                               // Micro seconds between updates of stats file
string StatsFileLocation;                                                           // Location of stats file
STATS SWEBSStats;

void AddVH(const map<VIRTUALHOST, VHSTATS>::value_type& p);                         // Creates all the virtual host nodes
void AddPage(const map<string, int>::value_type& p);                                // Adds all the page requests
void AddPage2(const map<string, int>::value_type& p);                               // Page Requests for virtual hosts
bool DeleteTopLine();                                                               // Deletes the top line of the xml file
void TestLog2(string Data);                                                         // Writes some data to testlog. Only TEST stuff

//----------------------------------------------------------------------------------------------------
//		    VHSTATS::VHSTATS()
//----------------------------------------------------------------------------------------------------
VHSTATS::VHSTATS()
{
    // Do nothing
}

VHSTATS::VHSTATS(VIRTUALHOST Thishost)
{
    ThisVH = Thishost;                                                              // Set the virtual host
    BytesSent = 0;
    NumberOfRequests = 0;
	if ( StatsFileLocation.empty() == false)										// If the key was there
    {
        // We know where the file is kept, so load the data from the XML file
        CkXml LoadXML;
        LoadXML.LoadXmlFile(StatsFileLocation.c_str());							    // Use the file we just got from the registry 
	
        CkXml *node = NULL;
        CkXml *node2 = NULL;
        CkXml *node3 = NULL;

	    node = LoadXML.SearchForTag(0,"VirtualHost");							    // Find the VH tag

        string Page;
        int Count = 0;
	    while (node)                                                                // Loop through all the virtual host tags until we find this one
        {
            node = LoadXML.SearchForTag(node,"vhHostName");
            if ( !strcmpi(node->get_Content(), ThisVH.HostName.c_str()) )
            {
                node2 = node;
                // We found the VH we were looking for! load its details
                node = LoadXML.SearchForTag(node2,"BytesSent");
                if (node)
                    BytesSent = StringToInt(node->get_Content());

                node = LoadXML.SearchForTag(node2,"RequestCount");
                if (node)
                    NumberOfRequests = StringToInt(node->get_Content());
               
                node = LoadXML.SearchForTag(node2,"PageRequest");
                if (node)
                {
                    node = LoadXML.SearchForTag(node2,"Page");
                    while (node)
                    {
                        node3 = node;
                        Page = node->get_Content();
                        node = LoadXML.SearchForTag(node3,"Count");
                        if (node)
                            Count = StringToInt(node->get_Content());
                        else Count = 0;
                        PageRequests[Page] = Count;
                        node = LoadXML.SearchForTag(node3,"Page");
                    }
                    
                }
                // We got the right VH, so we can break
                break;
            }

            // Didn't find the right VH, keep looking...
            node = LoadXML.SearchForTag(node,"VirtualHost");	 
        }
    }
}

//----------------------------------------------------------------------------------------------------
//		    VHSTATS Overloaded > and < operators
//----------------------------------------------------------------------------------------------------
bool operator>(const VHSTATS lhs, const VHSTATS rhs)                                // Must have this to use VHSTATS in a map
{
    if (lhs.NumberOfRequests > rhs.NumberOfRequests)
    {
        return true;
    }
    else return false;
}

bool operator<(const VHSTATS lhs, const VHSTATS rhs)                                // We must have this to use VHSTATS in a map
{
    if (lhs.NumberOfRequests < rhs.NumberOfRequests)
    {
        return true;
    }
    else return false;
}

void TestLog2(string Data)
{
	FILE* log;
	log = fopen(Options.Logfile.c_str(), "a+");
	if (log == NULL)
      return ;
	fprintf(log, "%s", Data.c_str());
	fclose(log);
}
//----------------------------------------------------------------------------------------------------
//		    STATS::STATS()
//----------------------------------------------------------------------------------------------------
STATS::STATS()
{
    // 1: Get the location of the stats file from the registry
    HKEY hKey;																		// Handle for the key
	unsigned long dwDisp;															// Disposition
	RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\SWS"), 0,
               NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dwDisp);

	unsigned char Buffer[_MAX_PATH];
	unsigned long DataType;
	unsigned long BufferLength = sizeof(Buffer);
	
	RegQueryValueEx(hKey, "StatsFile", NULL, &DataType, Buffer, &BufferLength);
    RegCloseKey(hKey);
	StatsFileLocation = (char *)Buffer;											    // Copy the stats file location

    // 2: Load the XML file and find values
    CkXml LoadXML;

    if ( LoadXML.LoadXmlFile(StatsFileLocation.c_str()) != true )
    {
        Options.LogError("Could not load stats file as XML");
    }

    CkXml * CurrentNode = NULL;
    CkXml * StartOfPageRequest = NULL;

    //------------------------------------------------------------------------------------------------
    // Find the last restart time
    CurrentNode = LoadXML.SearchForTag(0,"LastRestart");
    if (CurrentNode)
        LastRestart = CurrentNode->get_Content();

    // Get the BytesSent
    CurrentNode = LoadXML.SearchForTag(0,"BytesSent");
    if (CurrentNode)
        BytesSent = StringToInt(CurrentNode->get_Content());
    
    // Get the TotalBytesSent
    CurrentNode = LoadXML.SearchForTag(0,"TotalBytesSent");
    if (CurrentNode)
        TotalBytesSent = StringToInt(CurrentNode->get_Content());
    
    // Get the RequestCount
    CurrentNode = LoadXML.SearchForTag(0,"RequestCount");
    if (CurrentNode)
        NumberOfRequests = StringToInt(CurrentNode->get_Content());
        
    // Page requests
    StartOfPageRequest = LoadXML.SearchForTag(0,"PageRequest");
    string Page;
    int Count;

    while (StartOfPageRequest)                                                      // Keep going while we are finding pages
    {
        CurrentNode = LoadXML.SearchForTag(StartOfPageRequest,"Page");
        if (CurrentNode)
        {
            Page = CurrentNode->get_Content();                                      // Set the page
            CurrentNode = LoadXML.SearchForTag(StartOfPageRequest,"Count");
            if (CurrentNode) 
            {
                Count = StringToInt(CurrentNode->get_Content());
                PageRequests[Page] = Count;                                         // Get the Count
            }
            
        }
        StartOfPageRequest = LoadXML.SearchForTag(StartOfPageRequest,"PageRequest");// Go on to the next page
    }

    // 3: Load all the virtual hosts
    int X = 0;
    while (VHI.HostNumbers[X].empty() == false)
    {
        string Name = VHI.HostNumbers[X];
        VHSTATS TempVHStats(VHI.Host[Name]); 
        VirtualHosts[VHI.Host[Name]] = TempVHStats;
    }

}
//----------------------------------------------------------------------------------------------------
//		    Handle Stats File
//----------------------------------------------------------------------------------------------------
DWORD WINAPI HandleStatsFile(LPVOID lpParam )
{

    // This function writes the stats file every 5 mins
    struct tm *tm_now;
    time_t now;
    char buff[1024];

    now = time ( NULL );
    tm_now = localtime ( &now );

    strftime ( buff, sizeof buff, "%I:%M %p %d/%m/%Y", tm_now );                 // Get the current time

    SWEBSStats.LastRestart = buff;                                                  // Set last restart time
    SWEBSStats.WriteStatsFile();                                                    // Write it once
    while (true)
    {
        Sleep(TimeBetweenWrites);
        SWEBSStats.WriteStatsFile();
    }
    return true;
}

//----------------------------------------------------------------------------------------------------
//      GLOBALS
//----------------------------------------------------------------------------------------------------
ofstream osOutput;


//----------------------------------------------------------------------------------------------------

void AddPage(const map<string, int>::value_type& p)
{
    osOutput << "<PageRequest>" << endl;
    osOutput << "  <Page>" << p.first << "</Page>" << endl;
    osOutput << "  <Count>" << p.second << "</Count>" << endl;
    osOutput << "</PageRequest>" << endl << endl;
}    
    
void AddPage2(const map<string, int>::value_type& p)
{
    osOutput << "  <PageRequest>" << endl;
    osOutput << "    <Page>" << p.first << "</Page>" << endl;
    osOutput << "    <Count>" << p.second << "</Count>" << endl;
    osOutput << "  </PageRequest>" << endl << endl;
}

void AddVH(const map<VIRTUALHOST, VHSTATS>::value_type& p)
{
    osOutput << "<VirtualHost>" << endl;                                            // Virtual Host
    osOutput << "  <vhHostName>" << p.first.HostName << "</vhHostName>" << endl;    // vhHostName
    osOutput << "  <BytesSent>" << p.second.BytesSent << "</BytesSent>" << endl;    // Bytes sent
    osOutput << "  <RequestCount>" << p.second.NumberOfRequests << "</RequestCount>" << endl;

    for_each(p.second.PageRequests.begin(), p.second.PageRequests.end(), AddPage2); // Page Requests

    osOutput << "</VirtualHost>" << endl;
}

bool STATS::WriteStatsFile()
{
    //------------------------------------------------------------------------------
    // Rewritten
    
    osOutput.open(StatsFileLocation.c_str());                                       // Open the file
    if (!osOutput)
    {
        return false;
    }
    osOutput << "<stats>" << endl;                                                  // Root node
    osOutput << "<LastRestart>" << LastRestart << "</LastRestart>" << endl;         // Last restart
    osOutput << "<RequestCount>" << NumberOfRequests << "</RequestCount>" << endl;  // Request Count
    osOutput << "<BytesSent>" << BytesSent << "</BytesSent>" << endl;               // Bytes sent
    osOutput << "<TotalBytesSent>" << TotalBytesSent << "</TotalBytesSent>" << endl;// Total Bytes sent
    
    for_each(PageRequests.begin(), PageRequests.end(), AddPage);                    // Add all the pages 
    for_each(VirtualHosts.begin(), VirtualHosts.end(), AddVH);                      // Add all the virtual hosts

    osOutput << "</stats>" << endl;                                                 // Finish it off
    osOutput.close();                                                               // Close osOutput
    return true;
}



