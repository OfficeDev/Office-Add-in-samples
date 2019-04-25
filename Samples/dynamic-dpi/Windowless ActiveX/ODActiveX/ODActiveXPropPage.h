#pragma once

// ODActiveXPropPage.h : Declaration of the CODActiveXPropPage property page class.


// CODActiveXPropPage : See ODActiveXPropPage.cpp for implementation.

class CODActiveXPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CODActiveXPropPage)
	DECLARE_OLECREATE_EX(CODActiveXPropPage)

// Constructor
public:
	CODActiveXPropPage();

// Dialog Data
	enum { IDD = IDD_PROPPAGE_ODACTIVEX };

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	DECLARE_MESSAGE_MAP()
};

