//----------------------------------------------------------------------------------
//		SWEBSGlobals.hpp
//		------------------
//		This header defines the SWEBSGlobals class. The class contains EVERYTHING
//		globally used throughout the SWEBS Web Server. General settings, options, 
//		statistics and virtual hosts are all defined here.
//		
//		Copyright Paul Stovell, 2003. All rights reserved.
//----------------------------------------------------------------------------------
//		INCLUDES
//----------------------------------------------------------------------------------
#include <windows.h>
#include <string>
#include <map>
#include <time.h>
#include <fstream>
#include <sstream>
#include <algorithm>
// These files are to be used for parsing XML documents (the config file) and must be downloaded
//  from: www.xml-parser.com
#include <CkXml.h>
#include <CkString.h>
#include <CkSettings.h>
#include "../Include/SWEBSCGI.hpp"
#include "../Include/SWEBSGlobals.hpp"
#include "../Include/SWEBSUtilities.hpp"

using namespace std;


// The first two of these files come from the Chilkat XML library, freely downloadable from
//  www.xml-parser.com
//#pragma comment(lib, "ChilkatRelSt.lib")
//#pragma comment(lib, "CkBaseRelSt.lib")

//----------------------------------------------------------------------------------
//          Global variables
//----------------------------------------------------------------------------------
SWEBSGLOBALS SWEBSGlobals;
string StatsFileLocation;

//----------------------------------------------------------------------------------------------------
//			Options::ReadSettings()
//----------------------------------------------------------------------------------------------------
bool SWEBSGLOBALS::ReadSettings()
{
	// Locate the configuration file - its location is inthe registry as
	// HKEY_LOCAL_SYSTEM\\Software\\SWS\\ConfigFile

	HKEY hKey;																		// Handle for the key
	unsigned long dwDisp;															// Disposition
	RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\SWS"), 0,
               NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dwDisp);

	unsigned char Buffer[_MAX_PATH];
	unsigned long DataType;
	unsigned long BufferLength = sizeof(Buffer);
	
	RegQueryValueEx(hKey, "ConfigFile", NULL, &DataType, Buffer, &BufferLength);

	string ConfigFileLocation;
	ConfigFileLocation = (char *)Buffer;											// Copy the config file location

	if ( ConfigFileLocation.empty())												// If the key was not there
	{
		return 0;
	}

	RegCloseKey(hKey);
	
	//===========================
	//	Open XML file
	//===========================
	CkXml xml;
    xml.LoadXmlFile(ConfigFileLocation.c_str());									// Use the file we just got from the registry 
	
	// Server's name
	CkXml *node = xml.SearchForTag(0,"ServerName");									// Find the server name first
	CkXml *node2;																	// Used later

	if (node)
	{
		Servername = node->get_Content();
	}
	// Port number
	node = xml.SearchForTag(0,"Port");												// Find the port
	if (node)
	{									
		string SPort = node->get_Content();											// Put it in a string
		Port = StringToInt(SPort);													// Convert to integer
	}

    // IP Address
    node = xml.SearchForTag(0, "ListeningAddress");
    if (node)
    {
        IPAddress = node->get_Content();
    }
    else IPAddress = "";

    // Webroot
	node = xml.SearchForTag(0,"Webroot");											// Root web folder
	if (node)
	{									
		WebRoot = node->get_Content();	
	}
	// Max connections
	node = xml.SearchForTag(0,"MaxConnections");									// Max connections
	if (node)
	{
		MaxConnections = StringToInt(node->get_Content());
	}
	
	// Log file
	node = xml.SearchForTag(0,"LogFile");
	if (node)
	{
		Logfile = node->get_Content();
	}
	
    // Error Log
    node = xml.SearchForTag(0, "ErrorLog");
    if (node)
    {
        ErrorLog = node->get_Content();
    }

	// ErrorPages
	node = xml.SearchForTag(0,"ErrorPages");
	if (node)
	{
	    ErrorDirectory = node->get_Content();
	}

	// Loop through with the CGI entries
	node = xml.SearchForTag(0,"CGI");
	while (node)
	{
		// Map Extension to Interpreter
		node2 = xml.SearchForTag(node, "Extension");
		string cExt;
		if (node2)
		{
			cExt = node2->get_Content();
		}
		
		node2 = xml.SearchForTag(node, "Interpreter");
		if (node2)
		{
			CGI[cExt] = node2->get_Content();
		}

		CkXml *curNode = node;

		node = xml.SearchForTag(curNode,"CGI");
		delete curNode;
	}
	
	// Index files
	node = xml.SearchForTag(0,"IndexFile");
	int X = 0;
	while (node)
	{
		IndexFiles[X] = node->get_Content();
		X++;
		
		CkXml *curNode = node;
		node = xml.SearchForTag(curNode,"IndexFile");
		delete curNode;
	}
	
	// Allow indexing
	node = xml.SearchForTag(0,"AllowIndex");
	if (node)
	{
		if ( !strcmpi(node->get_Content(), "true" ))
		{
			AllowIndex = true;
		}
		else AllowIndex = false;
	}

	// Virtual Hosts
	string sName;
	string sHostName;
	string sRoot;
	string sLogFile;
	string Index;
    X = 0;
	node = xml.SearchForTag(0,"VirtualHost");
	while (node)
	{
		node2 = xml.SearchForTag(node, "vhName");
		if (node2)
		{
			sName = node2->get_Content();
		}
		
		node2 = xml.SearchForTag(node, "vhHostName");
		if (node2)
		{
			sHostName = node2->get_Content();
		}

		node2 = xml.SearchForTag(node, "vhRoot");
		if (node2)
		{
			sRoot = node2->get_Content();
		}

		node2 = xml.SearchForTag(node, "vhLogFile");
		if (node2)
		{
			sLogFile = node2->get_Content();
		}

		if ( !sName.empty() && !sHostName.empty() && !sRoot.empty() && !sLogFile.empty())
		{
			Host[sHostName].HostName = sHostName;
			Host[sHostName].Logfile = sLogFile;
			Host[sHostName].Name = sName;
			Host[sHostName].Root = sRoot;
            
            HostNumbers[X] = sHostName;
		}

		CkXml *curNode = node;

		node = xml.SearchForTag(curNode,"VirtualHost");
        X++;
		delete curNode;
	}
    NumberOfHosts = X;

	return 1;
}

//----------------------------------------------------------------------------------------------------
//		    VIRTUALHOST < and > operators
//----------------------------------------------------------------------------------------------------
bool operator>(const SWEBSVIRTUALHOST lhs, const SWEBSVIRTUALHOST rhs)
{
    if (lhs.Root.length() > rhs.Root.length())
    {
        return true;
    }
    else return false;
}

bool operator<(const SWEBSVIRTUALHOST lhs, const SWEBSVIRTUALHOST rhs)
{
    if (lhs.Root.length() < rhs.Root.length())
    {
        return true;
    }
    else return false; 
}

//----------------------------------------------------------------------------------------------------
//      OPTIONS::LogError
//----------------------------------------------------------------------------------------------------
bool SWEBSGLOBALS::LogError(string Text)
{
    FILE* log;

    struct tm *tm_now;
    time_t now;
    char buff[1024];

    now = time ( NULL );
    tm_now = localtime ( &now );

    strftime ( buff, sizeof buff, "%I:%M %p %d/%m/%Y %Z", tm_now );                 // Get the current time

    log = fopen(Logfile.c_str(), "a+");                                 
	if (log == NULL)
    {
        return false;
    }
    fprintf(log, "%s - ", buff);
	fprintf(log, "%s\n", Text.c_str());                                               // Write the string
	fclose(log);                                                                    // Close it
	return true;
}

//----------------------------------------------------------------------------------
//		STATISTICS FUNCTIONS 
//----------------------------------------------------------------------------------
bool SWEBSVIRTUALHOST::LoadStats()
{
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
            if ( !strcmpi(node->get_Content(), HostName.c_str()) )
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
    return true;
}

//----------------------------------------------------------------------------------
//		SWEBSGLOBALS::SWEBSGLOBALS
//----------------------------------------------------------------------------------
bool SWEBSGLOBALS::LoadStats()
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
        LogError("Could not load stats file as XML");
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
    while (HostNumbers[X].empty() == false)
    {
        Host[HostNumbers[X]].LoadStats();
    }

    return true;

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

    SWEBSGlobals.LastRestart = buff;                                                  // Set last restart time
    SWEBSGlobals.WriteStatsFile();                                                    // Write it once
    while (true)
    {
        Sleep(300000);
        SWEBSGlobals.WriteStatsFile();
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

void AddVH(const map<string, SWEBSVIRTUALHOST>::value_type& p)
{
    osOutput << "<VirtualHost>" << endl;                                            // Virtual Host
    osOutput << "  <vhHostName>" << p.second.HostName << "</vhHostName>" << endl;    // vhHostName
    osOutput << "  <BytesSent>" << p.second.BytesSent << "</BytesSent>" << endl;    // Bytes sent
    osOutput << "  <RequestCount>" << p.second.NumberOfRequests << "</RequestCount>" << endl;

    for_each(p.second.PageRequests.begin(), p.second.PageRequests.end(), AddPage2); // Page Requests

    osOutput << "</VirtualHost>" << endl;
}

bool SWEBSGLOBALS::WriteStatsFile()
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
    for_each(Host.begin(), Host.end(), AddVH);                                      // Add all the virtual hosts

    osOutput << "</stats>" << endl;                                                 // Finish it off
    osOutput.close();                                                               // Close osOutput
    return true;
}