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
		eventidUseDynamicDPIAwareCodeChanging = 2L,
		dispidUseDynamicDPIAwareCode = 4,
		dispidGetProcessThreadID = 3L,
		eventidFloatPropertyChanging = 1L,
		dispidFloatProperty = 2,
		dispidHelloWorld = 2L,
	};

	CMainDialog m_MainDialog;
	BOOL m_fUseDDPICode = false;

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
protected:
	// The float field used by FloatProperty
	FLOAT m_fField;

	// True to use the DDPI Aware code
	BOOL m_UseDynamicDPIAwareCode;
protected:
	BSTR HelloWorld(void);
	void GetProcessThreadID(LONG* pdwProcessId, LONG* pdwThreadId);
	FLOAT GetFloatProperty(void);
	void SetFloatProperty(FLOAT newVal);

	void FloatPropertyChanging(FLOAT NewValue, VARIANT_BOOL* Cancel)
	{
		FireEvent(eventidFloatPropertyChanging, EVENT_PARAM(VTS_R4 VTS_PBOOL), NewValue, Cancel);
	}
	VARIANT_BOOL GetUseDynamicDPIAwareCode();
	void SetUseDynamicDPIAwareCode(VARIANT_BOOL newVal);

	void UseDynamicDPIAwareCodeChanging(VARIANT_BOOL newValue, VARIANT_BOOL* cancel)
	{
		FireEvent(eventidUseDynamicDPIAwareCodeChanging, EVENT_PARAM(VTS_BOOL VTS_PBOOL), newValue, cancel);

		// HACK: cancel is always coming back as 1 (at least on 64 bit).  Possibly param passing issue with MFC?
		cancel = VARIANT_FALSE;
	}
};

