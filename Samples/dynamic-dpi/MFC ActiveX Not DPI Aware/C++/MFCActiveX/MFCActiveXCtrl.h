/****************************** Module Header ******************************\
Module Name:  MFCActiveXCtrl.h
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

Declaration of the CMFCActiveXCtrl ActiveX Control class.
See MFCActiveXCtrl.cpp for the implementation.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/en-us/openness/resources/licenses.aspx#MPL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#pragma once
#include "maindialog.h"


class CMFCActiveXCtrl : public COleControl
{
	DECLARE_DYNCREATE(CMFCActiveXCtrl)

// Constructor
public:
	CMFCActiveXCtrl();

// Overrides
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual DWORD GetControlFlags();
	virtual void OnAmbientPropertyChange(DISPID dispid);

// Implementation
protected:
	~CMFCActiveXCtrl();

	BEGIN_OLEFACTORY(CMFCActiveXCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CMFCActiveXCtrl)

	DECLARE_OLETYPELIB(CMFCActiveXCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CMFCActiveXCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CMFCActiveXCtrl)		// Type name and misc status

	// Subclassed control support
	BOOL IsSubclassedControl();
	LRESULT OnOcmCommand(WPARAM wParam, LPARAM lParam);

// Message maps
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	DECLARE_DISPATCH_MAP()

// Event maps
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
		dispidGetProcessThreadID = 3L,
		eventidDisplayValuePropertyChanging = 1L,
		dispidDisplayValueProperty = 2L,
		dispidHelloWorld = 2L,
	};
	CMainDialog m_MainDialog;
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
protected:
	CString m_cstrField;
	BSTR HelloWorld(void);
	void GetProcessThreadID(LONG* pdwProcessId, LONG* pdwThreadId);
	CString GetDisplayValueProperty(void);
	void SetDisplayValueProperty(CString newVal);

	void DisplayValuePropertyChanging(CString NewValue, VARIANT_BOOL* Cancel)
	{
		FireEvent(eventidDisplayValuePropertyChanging, EVENT_PARAM(VTS_R4 VTS_PBOOL), NewValue, Cancel);
	}
};


