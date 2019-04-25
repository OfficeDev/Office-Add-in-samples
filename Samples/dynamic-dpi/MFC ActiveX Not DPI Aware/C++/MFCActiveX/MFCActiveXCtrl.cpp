/****************************** Module Header ******************************\
Module Name:  MFCActiveXCtrl.cpp
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

Implementation of the CMFCActiveXCtrl ActiveX Control class.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/en-us/openness/resources/licenses.aspx#MPL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#pragma region Includes
#include "stdafx.h"
#include "MFCActiveX.h"
#include "MFCActiveXCtrl.h"
#include "MFCActiveXPropPage.h"
#include <olectl.h>
#include <string>
#include <windows.h>

#include <OCIdl.h>

#pragma endregion

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CMFCActiveXCtrl, COleControl)



// Message map

BEGIN_MESSAGE_MAP(CMFCActiveXCtrl, COleControl)
	ON_MESSAGE(OCM_COMMAND, &CMFCActiveXCtrl::OnOcmCommand)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_WM_CREATE()
END_MESSAGE_MAP()



// Dispatch map

BEGIN_DISPATCH_MAP(CMFCActiveXCtrl, COleControl)
	DISP_FUNCTION_ID(CMFCActiveXCtrl, "HelloWorld", dispidHelloWorld, HelloWorld, VT_BSTR, VTS_NONE)
	DISP_PROPERTY_EX_ID(CMFCActiveXCtrl, "DisplayValueProperty", dispidDisplayValueProperty, GetDisplayValueProperty, SetDisplayValueProperty, VT_BSTR)
	DISP_FUNCTION_ID(CMFCActiveXCtrl, "GetProcessThreadID", dispidGetProcessThreadID, GetProcessThreadID, VT_EMPTY, VTS_PI4 VTS_PI4)
END_DISPATCH_MAP()



// Event map

BEGIN_EVENT_MAP(CMFCActiveXCtrl, COleControl)
	EVENT_CUSTOM_ID("DisplayValuePropertyChanging", eventidDisplayValuePropertyChanging, DisplayValuePropertyChanging, VTS_R4 VTS_PBOOL)
END_EVENT_MAP()



// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CMFCActiveXCtrl, 1)
	PROPPAGEID(CMFCActiveXPropPage::guid)
END_PROPPAGEIDS(CMFCActiveXCtrl)



// Initialize class factory and guid

// {A039CAF5-6934-42D2-8456-E7A92D98FA07}
IMPLEMENT_OLECREATE_EX(CMFCActiveXCtrl, "MFCACTIVEXNOTDPI.MFCActiveXCtrl.1",
	0xa039caf5, 0x6934, 0x42d2, 0x84, 0x56, 0xe7, 0xa9, 0x2d, 0x98, 0xfa, 0x7);


// Type library ID and version

IMPLEMENT_OLETYPELIB(CMFCActiveXCtrl, _tlid, _wVerMajor, _wVerMinor)



// Interface IDs

// {4C7FE461-74CD-4638-A1EE-1EAE3830953A}
const IID BASED_CODE IID_DMFCActiveX =
		{ 0x4c7fe461, 0x74cd, 0x4638,{ 0xa1, 0xee, 0x1e, 0xae, 0x38, 0x30, 0x95, 0x3a } };

// {ED7A59CB-FAB3-48AA-A6D7-C42171F633B0}
const IID BASED_CODE IID_DMFCActiveXEvents =
		{ 0xed7a59cb, 0xfab3, 0x48aa,{ 0xa6, 0xd7, 0xc4, 0x21, 0x71, 0xf6, 0x33, 0xb0 } };



// Control type information

static const DWORD BASED_CODE _dwMFCActiveXOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CMFCActiveXCtrl, IDS_MFCACTIVEX, _dwMFCActiveXOleMisc)



// CMFCActiveXCtrl::CMFCActiveXCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CMFCActiveXCtrl

BOOL CMFCActiveXCtrl::CMFCActiveXCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_MFCACTIVEX,
			IDB_MFCACTIVEX,
			afxRegApartmentThreading,
			_dwMFCActiveXOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}



// Licensing strings

static const TCHAR BASED_CODE _szLicFileName[] = _T("MFCActiveXNotDpi.lic");

static const WCHAR BASED_CODE _szLicString[] =
	L"Copyright (c) 2009 ";



// CMFCActiveXCtrl::CMFCActiveXCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CMFCActiveXCtrl::CMFCActiveXCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}



// CMFCActiveXCtrl::CMFCActiveXCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CMFCActiveXCtrl::CMFCActiveXCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}



// CMFCActiveXCtrl::CMFCActiveXCtrl - Constructor

CMFCActiveXCtrl::CMFCActiveXCtrl() 
{
	InitializeIIDs(&IID_DMFCActiveX, &IID_DMFCActiveXEvents);
	// TODO: Initialize your control's instance data here.
}



// CMFCActiveXCtrl::~CMFCActiveXCtrl - Destructor

CMFCActiveXCtrl::~CMFCActiveXCtrl()
{
	// TODO: Cleanup your control's instance data here.
}



// CMFCActiveXCtrl::OnDraw - Drawing function

void CMFCActiveXCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!pdc)
		return;

	// To size the main dialog window and fill the background
	m_MainDialog.MoveWindow(rcBounds, TRUE);
	CBrush brBackGnd(TranslateColor(AmbientBackColor()));
	pdc->FillRect(rcBounds, &brBackGnd);

	DoSuperclassPaint(pdc, rcBounds);
}



// CMFCActiveXCtrl::DoPropExchange - Persistence support

void CMFCActiveXCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
}



// CMFCActiveXCtrl::GetControlFlags -
// Flags to customize MFC's implementation of ActiveX controls.
//
DWORD CMFCActiveXCtrl::GetControlFlags()
{
	DWORD dwFlags = COleControl::GetControlFlags();


	// The control will not be redrawn when making the transition
	// between the active and inactivate state.
	dwFlags |= noFlickerActivate;
	return dwFlags;
}



// CMFCActiveXCtrl::OnResetState - Reset control to default state

void CMFCActiveXCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}



// CMFCActiveXCtrl::PreCreateWindow - Modify parameters for CreateWindowEx

BOOL CMFCActiveXCtrl::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.lpszClass = _T("STATIC");
	return COleControl::PreCreateWindow(cs);
}



// CMFCActiveXCtrl::IsSubclassedControl - This is a subclassed control

BOOL CMFCActiveXCtrl::IsSubclassedControl()
{
	return TRUE;
}



// CMFCActiveXCtrl::OnOcmCommand - Handle command messages

LRESULT CMFCActiveXCtrl::OnOcmCommand(WPARAM wParam, LPARAM lParam)
{
#ifdef _WIN32
	WORD wNotifyCode = HIWORD(wParam);
#else
	WORD wNotifyCode = HIWORD(lParam);
#endif

	// TODO: Switch on wNotifyCode here.

	return 0;
}



// CMFCActiveXCtrl message handlers

int CMFCActiveXCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;

	// To create the main dialog
	m_MainDialog.Create(IDD_MAINDIALOG, this);

	return 0;
}


BSTR CMFCActiveXCtrl::HelloWorld(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strResult;
	strResult = _T("HelloWorld");
	
	return strResult.AllocSysString();
}

CString CMFCActiveXCtrl::GetDisplayValueProperty(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your dispatch handler code here
	return this->m_cstrField;
}

void CMFCActiveXCtrl::SetDisplayValueProperty(CString newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your property handler code here
	
	// Fire the event
	VARIANT_BOOL cancel = VARIANT_FALSE; 
	DisplayValuePropertyChanging(newVal, &cancel);

	if (cancel == VARIANT_FALSE)
	{
		m_cstrField += newVal + CString(L"\n");	// Save the new value
		SetModifiedFlag();

		// Display the new value in the control UI
		TRACE(newVal);
		try {
			m_MainDialog.m_StaticDisplayValueProperty.SetWindowTextW(m_cstrField);
		}
		catch (int e) {
			TRACE(L"Exception " + e);
		}
	}
	// else, do nothing.
}

void CMFCActiveXCtrl::GetProcessThreadID(LONG* pdwProcessId, LONG* pdwThreadId)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your dispatch handler code here
	*pdwProcessId = GetCurrentProcessId();
	*pdwThreadId = GetCurrentThreadId();
}

void CMFCActiveXCtrl::OnAmbientPropertyChange(DISPID dispid)
{
	// Respond to ambient property changes
}


