/****************************** Module Header ******************************\
Module Name:  MainDialog.h
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

CMainDialog dialog

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/en-us/openness/resources/licenses.aspx#MPL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#pragma once
#include "afxwin.h"


class CMainDialog : public CDialog
{
	DECLARE_DYNAMIC(CMainDialog)

public:
	CMainDialog(CWnd* pParent = NULL);   // standard constructor
	virtual ~CMainDialog();

// Dialog Data
	enum { IDD = IDD_MAINDIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedMsgBoxBn();
	afx_msg LRESULT OnDPIMessage(WPARAM wParam, LPARAM lParam);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	
	static BOOL CALLBACK EnumChildProc(HWND hWnd, LPARAM lParam);

	CStatic m_StaticFloatProperty;
	CButton m_CheckUseDpi;

private:
	CEdit m_EditMessage;
	CStatic m_StaticDPI;
	UINT m_newDPI;
	UINT m_currentDPI;

	BOOL UseDDpiCode();

public:
	afx_msg void OnBnClickedCheck1();

};
