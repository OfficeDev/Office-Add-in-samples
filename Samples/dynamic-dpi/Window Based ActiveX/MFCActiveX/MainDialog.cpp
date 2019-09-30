/****************************** Module Header ******************************\
Module Name:  MainDialog.cpp
Project:      MFCActiveX
Copyright (c) Microsoft Corporation.

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
#include "MainDialog.h"
#include "MFCActiveXCtrl.h"
#pragma endregion

const int DESIRED_HEIGHT = 9;
#define   DESIRED_FONT_NAME L"Arial"

// CMainDialog dialog

IMPLEMENT_DYNAMIC(CMainDialog, CDialog)

CMainDialog::CMainDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CMainDialog::IDD, pParent)
{
	m_currentDPI = m_newDPI = 0;
}

CMainDialog::~CMainDialog()
{
}

void CMainDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_MSGBOX_EDIT, m_EditMessage);
	DDX_Control(pDX, IDC_FLOATPROP_STATIC, m_StaticFloatProperty);
	DDX_Control(pDX, IDC_DPI, m_StaticDPI);
	DDX_Control(pDX, IDC_CHECK_USEDPI, m_CheckUseDpi);
}


BEGIN_MESSAGE_MAP(CMainDialog, CDialog)
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_MSGBOX_BN, &CMainDialog::OnBnClickedMsgBoxBn)
	ON_MESSAGE(WM_DPICHANGED, OnDPIMessage)
	ON_WM_SIZE()
	ON_BN_CLICKED(IDC_CHECK_USEDPI, &CMainDialog::OnBnClickedCheck1)
END_MESSAGE_MAP()


// CMainDialog message handlers


void CMainDialog::OnBnClickedMsgBoxBn()
{
	CString strText;
	m_EditMessage.GetWindowText(strText);

	CString strDPI;
	strDPI.Format(_T("Click DPI: %d"), ::GetDpiForWindow(this->GetSafeHwnd()));
	strText.Append(strDPI);
	MessageBox(strText, _T("HelloWorld"), MB_ICONINFORMATION | MB_OK);
}

LRESULT CMainDialog::OnDPIMessage(WPARAM wParam, LPARAM lParam)
{
	CString strDPI;
	strDPI.Format(_T("DPI: %d"), ::GetDpiForWindow(this->GetSafeHwnd()));
	m_StaticDPI.SetWindowTextW(strDPI);
	return 1;
}

void ChangeWindowFontDPI(HWND hWnd, UINT dpi)
{
	LOGFONT fontInfo1 = { 0 };
	fontInfo1.lfHeight = -MulDiv(DESIRED_HEIGHT, dpi, 72);
	fontInfo1.lfQuality = CLEARTYPE_QUALITY;
	wcscpy_s(fontInfo1.lfFaceName, DESIRED_FONT_NAME);

	::SendMessage(hWnd, WM_SETFONT, (WPARAM)::CreateFontIndirectW(&fontInfo1), TRUE);
}

BOOL CALLBACK CMainDialog::EnumChildProc(HWND hWnd, LPARAM lParam)
{
	CMainDialog* _this = (CMainDialog*) lParam;
	if (_this != nullptr)
	{
		double zoom = (((double) _this->m_newDPI) / (((double) _this->m_currentDPI) / 100.0)) / 100;

		RECT rect = {};
		::GetWindowRect(hWnd, &rect);

		POINT pt = { rect.left, rect.top };
		::ScreenToClient(::GetParent(hWnd), &pt);

		::MoveWindow(hWnd,
			static_cast<int>(pt.x*zoom),
			static_cast<int>(pt.y*zoom),
			static_cast<int>((rect.right - rect.left)*zoom),
			static_cast<int>((rect.bottom - rect.top)*zoom),
			TRUE);

		ChangeWindowFontDPI(hWnd, _this->m_newDPI);
		return TRUE;
	}
	return FALSE;
}

void CMainDialog::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);

	m_newDPI = ::GetDpiForWindow(this->GetSafeHwnd());
	CString strDPI;
	strDPI.Format(_T("OnSize DPI: %d"), m_newDPI);
	if (m_StaticDPI) m_StaticDPI.SetWindowTextW(strDPI);

	if (!m_currentDPI)
	{
		m_currentDPI = m_newDPI;
	}

	if (UseDDpiCode())
	{
		::EnumChildWindows(this->GetSafeHwnd(), EnumChildProc, (LPARAM)this);
	}

	m_currentDPI = m_newDPI;
}


void CMainDialog::OnBnClickedCheck1()
{
	// m_CheckUseDpi
}

BOOL CMainDialog::UseDDpiCode()
{
	return m_CheckUseDpi ? m_CheckUseDpi.GetState() == BST_CHECKED : false;
}
