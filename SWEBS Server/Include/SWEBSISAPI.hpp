#ifndef SWEBSISAPIHPP
#define SWEBSISAPIHPP 1
//----------------------------------------------------------------------------------
//			SWEBSISAPI.hpp
//			-------------------
//			This header declares all the functions used by ISAPI processing. They
//			use the SWEBSISAPI namespace, so remember to include it. A tutorial on
//			Writing an ISAPI handler can be found at http://swebs.sf.net/
//
//----------------------------------------------------------------------------------																				
//			INCLUDES
//----------------------------------------------------------------------------------
#include "SWEBSCGI.hpp"
#include <httpext.h>
#include <httpfilt.h>
#include <header.h>

//----------------------------------------------------------------------------------
extern BOOL WINAPI ReadClientExport(HCONN ConnID,
                                         LPVOID lpvBuffer,
                                         LPDWORD lpdwSize);
extern BOOL WINAPI WriteClientExport(HCONN      ConnID,
                                          LPVOID     Buffer,
                                          LPDWORD    lpdwBytes,
                                          DWORD      dwReserved );
extern BOOL WINAPI GetServerVariableExport(HCONN hConn, 
                                                LPSTR lpszVariableName, 
                                                LPVOID lpvBuffer, 
                                                LPDWORD lpdwSize);
extern BOOL WINAPI ServerSupportFunctionExport(HCONN      hConn,
                                           DWORD      dwHSERequest,
                                           LPVOID     lpvBuffer,
                                           LPDWORD    lpdwSize,
                                           LPDWORD    lpdwDataType);

typedef DWORD (WINAPI *HTTPEXTENSIONPROC)( EXTENSION_CONTROL_BLOCK *pECB );         // The HTTP Extension proc function
typedef DWORD (WINAPI *GETEXTENSIONVERSION)( HSE_VERSION_INFO* hse );               // The Get Extension Version function
//----------------------------------------------------------------------------------
#endif