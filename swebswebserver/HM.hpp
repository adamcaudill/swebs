#ifndef HMHPP
#define HMHPP
//---------------------------------------------------------------------------------------------
/*
			HM
			--------------
			This is the HM class used by the SWEBS_headermap.DLL AND connection class.
*/
//---------------------------------------------------------------------------------------------
typedef bool (*SWEBS_HANDLE)(istringstream &IS,					// Function pointer
							  CONNECTION &Connection);

//---------------------------------------------------------------------------------------------

class HM
{
  public:
	CONNECTION &Connection;
	bool hm_HOST(istringstream &IS);
};

//---------------------------------------------------------------------------------------------
#endif HMHPP