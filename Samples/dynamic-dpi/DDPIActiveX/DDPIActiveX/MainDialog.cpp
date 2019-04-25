// MainDialog.cpp : implementation file
//

#include "stdafx.h"
#include "DDPIActiveX.h"
#include "MainDialog.h"
#include "afxdialogex.h"


// CMainDialog dialog

IMPLEMENT_DYNAMIC(CMainDialog, CDialogEx)

CMainDialog::CMainDialog(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_MAINDIALOG, pParent)
{

}

CMainDialog::~CMainDialog()
{
}

void CMainDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CMainDialog, CDialogEx)
END_MESSAGE_MAP()


// CMainDialog message handlers
