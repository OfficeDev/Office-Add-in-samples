// ODActiveXCtrl.cpp : Implementation of the CODActiveXCtrl ActiveX Control class.

#include "stdafx.h"
#include "ODActiveX.h"
#include "ODActiveXCtrl.h"
#include "ODActiveXPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CODActiveXCtrl, COleControl)

// Message map

BEGIN_MESSAGE_MAP(CODActiveXCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_EDIT, OnEdit)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// Dispatch map

BEGIN_DISPATCH_MAP(CODActiveXCtrl, COleControl)
END_DISPATCH_MAP()

// Event map

BEGIN_EVENT_MAP(CODActiveXCtrl, COleControl)
END_EVENT_MAP()

// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CODActiveXCtrl, 1)
	PROPPAGEID(CODActiveXPropPage::guid)
END_PROPPAGEIDS(CODActiveXCtrl)

// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CODActiveXCtrl, "ODACTIVEX.ODActiveXCtrl.1",
	0xdc5950e1, 0x9a17, 0x4d9e, 0xb8, 0x13, 0x7d, 0x7b, 0xfa, 0xe7, 0xce, 0x7e)

// Type library ID and version

IMPLEMENT_OLETYPELIB(CODActiveXCtrl, _tlid, _wVerMajor, _wVerMinor)

// Interface IDs

const IID IID_DODActiveX = { 0x79F0A437, 0xA95A, 0x4B96, { 0xA9, 0x5E, 0x2E, 0x84, 0x0, 0x6F, 0x8F, 0x93 } };
const IID IID_DODActiveXEvents = { 0xF3EE2BA1, 0xCC2, 0x4B96, { 0xA1, 0x33, 0x5, 0x4F, 0x6A, 0x26, 0x5, 0xB5 } };

// Control type information

static const DWORD _dwODActiveXOleMisc =
	OLEMISC_SIMPLEFRAME |
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CODActiveXCtrl, IDS_ODACTIVEX, _dwODActiveXOleMisc)

// CODActiveXCtrl::CODActiveXCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CODActiveXCtrl

BOOL CODActiveXCtrl::CODActiveXCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegInsertable | afxRegApartmentThreading to afxRegInsertable.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_ODACTIVEX,
			IDB_ODACTIVEX,
			afxRegInsertable | afxRegApartmentThreading,
			_dwODActiveXOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// CODActiveXCtrl::CODActiveXCtrl - Constructor

CODActiveXCtrl::CODActiveXCtrl()
{
	InitializeIIDs(&IID_DODActiveX, &IID_DODActiveXEvents);

	EnableSimpleFrame();
	
	oldDPI = 0;
}

// CODActiveXCtrl::~CODActiveXCtrl - Destructor

CODActiveXCtrl::~CODActiveXCtrl()
{
	// TODO: Cleanup your control's instance data here.
}

// CODActiveXCtrl::OnDraw - Drawing function

void CODActiveXCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;
	
	UINT dpi = oldDPI;

	if (pdc->GetWindow() != nullptr &&
		pdc->GetWindow()->GetSafeHwnd() != nullptr)
	{
		dpi = ::GetDpiForWindow(pdc->GetWindow()->GetSafeHwnd());
	}

	dpi = pdc->GetDeviceCaps(LOGPIXELSX);

	if (oldDPI == 0)
	{
		oldDPI = dpi;
	}

	CFont* cF = pdc->GetCurrentFont();
	LOGFONT lf = { 0 };
	cF->GetLogFont(&lf);

	LOGFONT fontInfo1 = { 0 };
	fontInfo1.lfHeight = MulDiv(lf.lfHeight, dpi, 72);
	fontInfo1.lfQuality = CLEARTYPE_QUALITY;
	wcscpy_s(fontInfo1.lfFaceName, lf.lfFaceName);
	CFont font;
	font.CreateFontIndirectW(&fontInfo1);

	CFont* pOldFont = pdc->SelectObject(&font);

	ULONGLONG  current = GetTickCount64();

	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(GRAY_BRUSH)));
	pdc->Ellipse(rcBounds);
	CString str;
	str.Format(L"DPI: %d %I64u", dpi, current);
	pdc->TextOutW(rcBounds.left + rcBounds.Width()/4, rcBounds.top + rcBounds.Height()/4, str);
	if (dpi)
	{
		str.Format(L"HWND: %X DPI: %d %d", pdc->GetWindow()->GetSafeHwnd(), dpi, ::GetDpiForWindow(pdc->GetWindow()->GetSafeHwnd()));
		pdc->TextOutW(rcBounds.left + rcBounds.Width() / 4, rcBounds.top + rcBounds.Height() / 4 + fontInfo1.lfHeight, str);
	}
	else
	{
		str.Format(L"No Info about Parent HWND");
		pdc->TextOutW(rcBounds.left + rcBounds.Width() / 4, rcBounds.top + rcBounds.Height() / 4 + fontInfo1.lfHeight, str);
	}

	pdc->SelectObject(pOldFont);
	font.DeleteObject();
}

// CODActiveXCtrl::DoPropExchange - Persistence support

void CODActiveXCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
}


// CODActiveXCtrl::GetControlFlags -
// Flags to customize MFC's implementation of ActiveX controls.
//
DWORD CODActiveXCtrl::GetControlFlags()
{
	DWORD dwFlags = COleControl::GetControlFlags();


	// The control can activate without creating a window.
	// TODO: when writing the control's message handlers, avoid using
	//		the m_hWnd member variable without first checking that its
	//		value is non-NULL.
	dwFlags |= windowlessActivate;
	return dwFlags;
}


// CODActiveXCtrl::OnResetState - Reset control to default state

void CODActiveXCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


// CODActiveXCtrl message handlers
