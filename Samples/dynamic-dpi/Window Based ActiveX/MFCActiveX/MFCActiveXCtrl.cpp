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
	DISP_PROPERTY_EX_ID(CMFCActiveXCtrl, "FloatProperty", dispidFloatProperty, GetFloatProperty, SetFloatProperty, VT_R4)
	DISP_FUNCTION_ID(CMFCActiveXCtrl, "GetProcessThreadID", dispidGetProcessThreadID, GetProcessThreadID, VT_EMPTY, VTS_PI4 VTS_PI4)
	DISP_PROPERTY_EX_ID(CMFCActiveXCtrl, "UseDynamicDPIAwareCode", dispidUseDynamicDPIAwareCode, GetUseDynamicDPIAwareCode, SetUseDynamicDPIAwareCode, VT_BOOL)
END_DISPATCH_MAP()



// Event map

BEGIN_EVENT_MAP(CMFCActiveXCtrl, COleControl)
	EVENT_CUSTOM_ID("FloatPropertyChanging", eventidFloatPropertyChanging, FloatPropertyChanging, VTS_R4 VTS_PBOOL)
	EVENT_CUSTOM_ID("UseDynamicDPIAwareCodeChanging", eventidUseDynamicDPIAwareCodeChanging, UseDynamicDPIAwareCodeChanging, VTS_BOOL VTS_PBOOL)
END_EVENT_MAP()



// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CMFCActiveXCtrl, 1)
	PROPPAGEID(CMFCActiveXPropPage::guid)
END_PROPPAGEIDS(CMFCActiveXCtrl)



// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMFCActiveXCtrl, "MFCACTIVEX.MFCActiveXCtrl.1",
	0xe389ad6c, 0x4fb6, 0x47af, 0xb0, 0x3a, 0xa5, 0xa5, 0xc6, 0xb2, 0xb8, 0x20)



// Type library ID and version

IMPLEMENT_OLETYPELIB(CMFCActiveXCtrl, _tlid, _wVerMajor, _wVerMinor)



// Interface IDs

const IID BASED_CODE IID_DMFCActiveX =
		{ 0x327DD42, 0x7A9E, 0x415B, { 0xB9, 0xA0, 0x4A, 0xEE, 0xE1, 0xA3, 0x31, 0x9E } };
const IID BASED_CODE IID_DMFCActiveXEvents =
		{ 0x97B9B2F3, 0xE95A, 0x49D4, { 0xAC, 0xA3, 0xE2, 0xA1, 0x81, 0x42, 0x4F, 0xD8 } };



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

static const TCHAR BASED_CODE _szLicFileName[] = _T("MFCActiveX.lic");

static const WCHAR BASED_CODE _szLicString[] =
	L"Copyright (c) 2009 ";



// CMFCActiveXCtrl::CMFCActiveXCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CMFCActiveXCtrl::CMFCActiveXCtrlFactory::VerifyUserLicense()
{
	return true;/* AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);*/
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

CMFCActiveXCtrl::CMFCActiveXCtrl() : m_fField(0.0f), m_UseDynamicDPIAwareCode(VARIANT_FALSE)
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

	if (m_MainDialog) 
	{
		// To size the main dialog window and fill the background
		m_MainDialog.MoveWindow(rcBounds, TRUE);
	}

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

FLOAT CMFCActiveXCtrl::GetFloatProperty(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your dispatch handler code here
	return this->m_fField;
}

void CMFCActiveXCtrl::SetFloatProperty(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your property handler code here
	
	// Fire the event, FloatPropertyChanging
	VARIANT_BOOL* cancel = new VARIANT_BOOL(VARIANT_FALSE);
	FloatPropertyChanging(newVal, cancel);

	if (*cancel == VARIANT_FALSE)
	{
		m_fField = newVal;	// Save the new value
		SetModifiedFlag();

		// Display the new value in the control UI
		CString strFloatProp;
		strFloatProp.Format(_T("%f"), m_fField);
		m_MainDialog.m_StaticFloatProperty.SetWindowTextW(strFloatProp);
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


VARIANT_BOOL CMFCActiveXCtrl::GetUseDynamicDPIAwareCode()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return this->m_UseDynamicDPIAwareCode;
}


void CMFCActiveXCtrl::SetUseDynamicDPIAwareCode(VARIANT_BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT_BOOL* cancel = new VARIANT_BOOL(VARIANT_FALSE);
	UseDynamicDPIAwareCodeChanging(newVal, cancel);
	if (*cancel == VARIANT_FALSE)
	{
		m_UseDynamicDPIAwareCode = (BOOL)newVal;
		SetModifiedFlag();

		m_MainDialog.m_CheckUseDpi.SetCheck(m_UseDynamicDPIAwareCode ? 1 : 0);
	}
}
