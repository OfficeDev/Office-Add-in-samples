/****************************** Module Header ******************************\
Module Name:  MFCActiveXPropPage.cpp
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

Implementation of the CMFCActiveXPropPage property page class.

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
#include "MFCActiveXPropPage.h"
#pragma endregion


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CMFCActiveXPropPage, COlePropertyPage)



// Message map

BEGIN_MESSAGE_MAP(CMFCActiveXPropPage, COlePropertyPage)
	ON_BN_CLICKED(IDC_CHECK_USEDPI, &CMFCActiveXPropPage::OnBnClickedCheckUsedpi)
END_MESSAGE_MAP()



// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMFCActiveXPropPage, "MFCACTIVEX.MFCActiveXPropPage.1",
	0xc870c834, 0x7228, 0x40e2, 0x84, 0xe4, 0x96, 0xb3, 0x87, 0x24, 0xa1, 0x4b)



// CMFCActiveXPropPage::CMFCActiveXPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CMFCActiveXPropPage

BOOL CMFCActiveXPropPage::CMFCActiveXPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_MFCACTIVEX_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}



// CMFCActiveXPropPage::CMFCActiveXPropPage - Constructor

CMFCActiveXPropPage::CMFCActiveXPropPage() :
	COlePropertyPage(IDD, IDS_MFCACTIVEX_PPG_CAPTION)
{
}



// CMFCActiveXPropPage::DoDataExchange - Moves data between page and properties

void CMFCActiveXPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_Check(pDX, IDC_CHECK_USEDPI, m_CheckUseDpi, _T("Use DDPI Code"));

	DDP_PostProcessing(pDX);
}

// CMFCActiveXPropPage message handlers


void CMFCActiveXPropPage::OnBnClickedCheckUsedpi()
{
}
