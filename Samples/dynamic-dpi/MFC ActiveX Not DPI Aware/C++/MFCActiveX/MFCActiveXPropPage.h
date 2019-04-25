/****************************** Module Header ******************************\
Module Name:  MFCActiveXPropPage.h
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

Declaration of the CMFCActiveXPropPage property page class.
See MFCActiveXPropPage.cpp for implementation.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/en-us/openness/resources/licenses.aspx#MPL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#pragma once


class CMFCActiveXPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CMFCActiveXPropPage)
	DECLARE_OLECREATE_EX(CMFCActiveXPropPage)

// Constructor
public:
	CMFCActiveXPropPage();

// Dialog Data
	enum { IDD = IDD_PROPPAGE_MFCACTIVEX };

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	DECLARE_MESSAGE_MAP()
};

