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
	log = fopen("C:\\SWS\\testlog.txt", "a+");
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
        TestLog2("Could not load stats file as XML");
    }

    CkXml * CurrentNode = NULL;
    CkXml * StartOfPageRequest = NULL;

    TestLog2("\n----STATS----\n");
    TestLog2("File Location: ");
    TestLog2(StatsFileLocation);

    //------------------------------------------------------------------------------------------------
    // Find the last restart time
    CurrentNode = LoadXML.SearchForTag(0,"LastRestart");
    if (CurrentNode)
        LastRestart = CurrentNode->get_Content();
    TestLog2("\nLastRestart: ");
    TestLog2(LastRestart);

    // Get the BytesSent
    CurrentNode = LoadXML.SearchForTag(0,"BytesSent");
    if (CurrentNode)
        BytesSent = StringToInt(CurrentNode->get_Content());
    TestLog2("\nBytesSent: ");
    TestLog2(IntToString(BytesSent));

    // Get the TotalBytesSent
    CurrentNode = LoadXML.SearchForTag(0,"TotalBytesSent");
    if (CurrentNode)
        TotalBytesSent = StringToInt(CurrentNode->get_Content());
    TestLog2("\nTotalBytesSent: ");
    TestLog2(IntToString(TotalBytesSent));

    // Get the RequestCount
    CurrentNode = LoadXML.SearchForTag(0,"RequestCount");
    if (CurrentNode)
        NumberOfRequests = StringToInt(CurrentNode->get_Content());
    TestLog2("\nRequestCount: ");
    TestLog2(IntToString(NumberOfRequests));
    
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
            TestLog2("\nPage: ");
            TestLog2(Page);
            TestLog2("\nCount: ");
            TestLog2(IntToString(Count));
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

    TestLog2("\nLoaded all stats from file\n");
    TestLog2("Bytes Sent: ");
    TestLog2(IntToString(BytesSent));

    TestLog2("\nRequests to myindex.html: ");
    TestLog2(IntToString(PageRequests["myindex.html"]));

}
//----------------------------------------------------------------------------------------------------
//		    Handle Stats File
//----------------------------------------------------------------------------------------------------
DWORD WINAPI HandleStatsFile(LPVOID lpParam )
{
    // This function writes the stats file every 5 mins
    SWEBSStats.LastRestart = "Today";
    while (true)
    {
        Sleep(300000);
        SWEBSStats.WriteStatsFile();
    }
    return true;
}

CkXml SaveXML;
CkXml PageXML;
CkXml VhXML;

void AddPage(const map<string, int>::value_type& p)
{
	PageXML.put_Tag("PageRequest");                                                 // Add a <PageRequest>
    PageXML.NewChild2("Page", p.first.c_str());                                     // Put in the page name
    PageXML.NewChild2("Count", IntToString(p.second).c_str());                              // And count
    SaveXML.AddChildTree(&PageXML);                                                 // Insert it into the rest of the XML
}

void AddPage2(const map<string, int>::value_type& p)
{
	PageXML.put_Tag("PageRequest");                                                 // Add a <PageRequest>
    PageXML.NewChild2("Page", p.first.c_str());                                     // Put in the page name
    PageXML.NewChild2("Count", IntToString(p.second).c_str());                              // And count
    VhXML.AddChildTree(&PageXML);                                                   // Insert it into the rest of the XML
}

void AddVH(const map<VIRTUALHOST, VHSTATS>::value_type& p)
{
	VhXML.put_Tag("VirtualHost");                                                   // Add a <PageRequest>
    VhXML.NewChild2("vhHostName", p.first.HostName.c_str());                        // Put in the host name
    VhXML.NewChild2("BytesSent", IntToString(p.second.BytesSent).c_str());                  // BytesSent
    VhXML.NewChild2("RequestCount", IntToString(p.second.NumberOfRequests).c_str());        // RequestCount
    
    for_each(p.second.PageRequests.begin(), p.second.PageRequests.end(), AddPage2); // Page Requests

    SaveXML.AddChildTree(&PageXML);                                                 // Insert it into the rest of the XML
}

bool DeleteTopLine()
{   
    // The chilkat library is excelent, but unfortunately when it outputs XML files it can't read them! The only way
    // around this is to delete the top line of the file (the <?xml?> line, or DTD).
    
    // TODO: Think of a faster way to do this :P
    
    char Buffer[1024];

    ifstream Input(StatsFileLocation.c_str());
    ofstream Output("c:\\sws\\tempstats.xml");

    if (!Input)
    {
        TestLog2("DeleteTopLine() Could not open ");
        TestLog2(StatsFileLocation.c_str());
        TestLog2("!\n");
        return false;
    }
    if (!Output)
    {
        TestLog2("DeleteTopLine() could not open ");
        TestLog2("c:\\sws\\tempstats.xml");
        TestLog2("!\n");
        return false;
    }

    // Get The first line, this must not be written to the file
    Input.getline(Buffer, 1024);
    if (strlen(Buffer) <= 0)
    {
        TestLog2("DeleteTopLine() The first line was not found.\n");
        return false;
    }

    while (!Input.eof())
    {
        Input.getline(Buffer, 1024);
        Output << Buffer;
        Output << "\n";
    }

    Input.close();
    Output.close();

    // Delete the original stats file
    DeleteFile(StatsFileLocation.c_str());

    // Now rename the temp file as stats.xml
    if (rename( "c:\\sws\\tempstats.xml", StatsFileLocation.c_str()))
    {
        TestLog2("DeleteFirstLine() Could not rename teststats.xml");
        return false;
    }


    // It all went fine
    return true;
}

bool STATS::WriteStatsFile()
{
    SaveXML.put_Tag("stats");                                                       // Create the root node
    
    // Put the current time here
    SaveXML.NewChild2("RequestCount", IntToString(NumberOfRequests).c_str());       // Save the request count
    SaveXML.NewChild2("BytesSent", IntToString(BytesSent).c_str());                 // Save the SytesSent
    SaveXML.NewChild2("TotalBytesSent", IntToString(TotalBytesSent).c_str());       // TotalBytesSent
    
    for_each(PageRequests.begin(), PageRequests.end(), AddPage);                    // This will go through and print each page/count pair 
    for_each(VirtualHosts.begin(), VirtualHosts.end(), AddVH);
    
    SaveXML.SaveXml(StatsFileLocation.c_str());                                     // Save the entire XML
    // The chilkat library is so stupid, it saves files as UTF-8 but cant read them! So somehow we have to go through
    // and delete the top line (DTD) from the xml file
    return DeleteTopLine();
}



