#pragma warning(disable:4786)
#include "options.hpp"

// Single instance classes
OPTIONS Options;
VirtualHostIndex VHI;

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
//			Options::ReadSettings()
//----------------------------------------------------------------------------------------------------
bool OPTIONS::ReadSettings()
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
	
	// ErrorPages
	node = xml.SearchForTag(0,"ErrorPages");
	if (node)
	{
		Options.ErrorDirectory = node->get_Content();
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
			VHI.Host[sHostName].HostName = sHostName;
			VHI.Host[sHostName].Logfile = sLogFile;
			VHI.Host[sHostName].Name = sName;
			VHI.Host[sHostName].Root = sRoot;
            
            VHI.HostNumbers[X] = sHostName;
		}

		CkXml *curNode = node;

		node = xml.SearchForTag(curNode,"VirtualHost");
        X++;
		delete curNode;
	}
    VHI.NumberOfHosts = X;

	return 1;
}

//----------------------------------------------------------------------------------------------------
//		    VIRTUALHOST < and > operators
//----------------------------------------------------------------------------------------------------
bool operator>(const VIRTUALHOST lhs, const VIRTUALHOST rhs)
{
    if (lhs.Root.length() > rhs.Root.length())
    {
        return true;
    }
    else return false;
}

bool operator<(const VIRTUALHOST lhs, const VIRTUALHOST rhs)
{
    if (lhs.Root.length() < rhs.Root.length())
    {
        return true;
    }
    else return false; 
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
