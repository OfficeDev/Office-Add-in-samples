#pragma once

// ODActiveXCtrl.h : Declaration of the CODActiveXCtrl ActiveX Control class.


// CODActiveXCtrl : See ODActiveXCtrl.cpp for implementation.

class CODActiveXCtrl : public COleControl
{
	DECLARE_DYNCREATE(CODActiveXCtrl)

// Constructor
public:
	CODActiveXCtrl();

// Overrides
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual DWORD GetControlFlags();

// Implementation
protected:
	~CODActiveXCtrl();

	DECLARE_OLECREATE_EX(CODActiveXCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CODActiveXCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CODActiveXCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CODActiveXCtrl)		// Type name and misc status

// Message maps
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	DECLARE_DISPATCH_MAP()

// Event maps
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	};

private:
	unsigned int oldDPI;
};

