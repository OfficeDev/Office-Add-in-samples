// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// ODActiveXPropPage.cpp : Implementation of the CODActiveXPropPage property page class.

#include "stdafx.h"
#include "ODActiveX.h"
#include "ODActiveXPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CODActiveXPropPage, COlePropertyPage)

// Message map

BEGIN_MESSAGE_MAP(CODActiveXPropPage, COlePropertyPage)
END_MESSAGE_MAP()

// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CODActiveXPropPage, "ODACTIVEX.ODActiveXPropPage.1",
	0x9c4848bb, 0x2f19, 0x4da1, 0x88, 0x99, 0x8f, 0x88, 0x2d, 0xba, 0x5a, 0x95)

// CODActiveXPropPage::CODActiveXPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CODActiveXPropPage

BOOL CODActiveXPropPage::CODActiveXPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_ODACTIVEX_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}

// CODActiveXPropPage::CODActiveXPropPage - Constructor

CODActiveXPropPage::CODActiveXPropPage() :
	COlePropertyPage(IDD, IDS_ODACTIVEX_PPG_CAPTION)
{
}

// CODActiveXPropPage::DoDataExchange - Moves data between page and properties

void CODActiveXPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// CODActiveXPropPage message handlers
