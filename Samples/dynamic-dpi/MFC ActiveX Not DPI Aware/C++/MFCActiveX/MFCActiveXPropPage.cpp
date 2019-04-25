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
END_MESSAGE_MAP()



// Initialize class factory and guid

// {1D490A55-70DD-4843-9BC5-9F84D973A115}
IMPLEMENT_OLECREATE_EX(CMFCActiveXPropPage, "MFCACTIVEXNOTDPI.MFCActiveXPropPage.1",
	0x1d490a55, 0x70dd, 0x4843, 0x9b, 0xc5, 0x9f, 0x84, 0xd9, 0x73, 0xa1, 0x15);


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
	DDP_PostProcessing(pDX);
}



// CMFCActiveXPropPage message handlers
