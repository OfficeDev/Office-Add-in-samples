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
#pragma endregion


// CMainDialog dialog

IMPLEMENT_DYNAMIC(CMainDialog, CDialog)

CMainDialog::CMainDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CMainDialog::IDD, pParent)
{

}

CMainDialog::~CMainDialog()
{
}

void CMainDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_MSGBOX_EDIT, m_EditMessage);
	DDX_Control(pDX, IDC_DISPLAYVALUE_EDIT, m_StaticDisplayValueProperty);
}


BEGIN_MESSAGE_MAP(CMainDialog, CDialog)
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_MSGBOX_BN, &CMainDialog::OnBnClickedMsgBoxBn)
END_MESSAGE_MAP()


// CMainDialog message handlers


void CMainDialog::OnBnClickedMsgBoxBn()
{
	CString strText;
	m_EditMessage.GetWindowText(strText);
	MessageBox(strText, _T("HelloWorld"), MB_ICONINFORMATION | MB_OK);
}
