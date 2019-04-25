// ODActiveX.cpp : Implementation of CODActiveXApp and DLL registration.

#include "stdafx.h"
#include "ODActiveX.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CODActiveXApp theApp;

const GUID CDECL _tlid = { 0x8007A0C4, 0x12BA, 0x44FF, { 0xA2, 0x75, 0xDA, 0xE8, 0x1D, 0x2F, 0xE5, 0x4D } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CODActiveXApp::InitInstance - DLL initialization

BOOL CODActiveXApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}



// CODActiveXApp::ExitInstance - DLL termination

int CODActiveXApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
