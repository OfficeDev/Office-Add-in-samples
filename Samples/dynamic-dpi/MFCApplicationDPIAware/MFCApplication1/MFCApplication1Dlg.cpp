// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// MFCApplication1Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

const int DESIRED_HEIGHT = 9;
#define   DESIRED_FONT_NAME L"Times New Roman"

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCApplication1Dlg dialog



CMFCApplication1Dlg::CMFCApplication1Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_MFCAPPLICATION1_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMFCApplication1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DPI_INFO, m_dpiInfo);
}

BEGIN_MESSAGE_MAP(CMFCApplication1Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_MOVE()
#ifdef WM_DPICHANGED
	ON_MESSAGE(WM_DPICHANGED, OnDPIMessage)
#else
	ON_MESSAGE(0x02E0, OnDPIMessage)
#endif
END_MESSAGE_MAP()

// CMFCApplication1Dlg message handlers

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	newDPI = currentDPI = ::GetDpiForWindow(this->GetSafeHwnd());

	::EnumChildWindows(this->GetSafeHwnd(), EnumChildProc, (LPARAM)this);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMFCApplication1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}

	{
		CClientDC dc(this); // device context for painting
		CDC* pdc = &dc;

		LOGFONT fontInfo1 = { 0 };
		fontInfo1.lfHeight = -MulDiv(DESIRED_HEIGHT, ::GetDpiForWindow(this->GetSafeHwnd()), 72);
		fontInfo1.lfQuality = CLEARTYPE_QUALITY;
		wcscpy_s(fontInfo1.lfFaceName, DESIRED_FONT_NAME);
		CFont font;
		font.CreateFontIndirectW(&fontInfo1);

		CFont* pOldFont = pdc->SelectObject(&font);

		CString str;
		str.Format(L"OWD DPI: %d", pdc->GetDeviceCaps(LOGPIXELSX));
		pdc->TextOutW(0, 0, str);

		pdc->SelectObject(pOldFont);
		font.DeleteObject();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CMFCApplication1Dlg::OnMove(int x, int y)
{
	CDialogEx::OnMove(x, y);
}


void ChangeWindowFontDPI(HWND hWnd, UINT dpi)
{	
	LOGFONT fontInfo1 = { 0 };
	fontInfo1.lfHeight = -MulDiv(DESIRED_HEIGHT, dpi, 72);
	fontInfo1.lfQuality = CLEARTYPE_QUALITY;
	wcscpy_s(fontInfo1.lfFaceName, DESIRED_FONT_NAME);

	::SendMessage(hWnd, WM_SETFONT, (WPARAM)::CreateFontIndirectW(&fontInfo1), TRUE);
}

BOOL CALLBACK CMFCApplication1Dlg::EnumChildProc(HWND hWnd, LPARAM lParam)
{
	CMFCApplication1Dlg* _this = (CMFCApplication1Dlg*) lParam;
	if (_this != nullptr)
	{
		// if you want all controls on parent to be resized
		// remove this, otherwise all compozite controls(including ActiveX) 
		// will care about themselfs
		if (_this->GetSafeHwnd() != ::GetParent(hWnd))
		{
			return TRUE;
		}

		double zoom = (((double) _this->newDPI) / (((double) _this->currentDPI) / 100.0)) / 100;

		RECT rect = {};
		::GetWindowRect(hWnd, &rect);		

		POINT pt = { rect.left, rect.top };
		::ScreenToClient(::GetParent(hWnd), &pt);

		::MoveWindow(hWnd,
			pt.x*zoom,
			pt.y*zoom,
			(rect.right - rect.left)*zoom,
			(rect.bottom - rect.top)*zoom,
			TRUE);

		ChangeWindowFontDPI(hWnd, _this->newDPI);
		return TRUE;
	}
	return FALSE;
}

LRESULT CMFCApplication1Dlg::OnDPIMessage(WPARAM wParam, LPARAM lParam)
{

	RECT* const prcNewWindow = (RECT*) lParam;
	::SetWindowPos(this->GetSafeHwnd(),
		NULL,
		prcNewWindow->left,
		prcNewWindow->top,
		prcNewWindow->right - prcNewWindow->left,
		prcNewWindow->bottom - prcNewWindow->top,
		SWP_NOZORDER | SWP_NOACTIVATE);
	::InvalidateRect(this->GetSafeHwnd(), nullptr, TRUE);

	newDPI = HIWORD(wParam);
	::EnumChildWindows(this->GetSafeHwnd(), EnumChildProc, (LPARAM)this);
	currentDPI = newDPI;

	CString str;
	str.Format(L"HWND:%X ON_DPICHANGED:%d GetDPIFromWindow:%d", this->GetSafeHwnd(), newDPI, ::GetDpiForWindow(this->GetSafeHwnd()));
	m_dpiInfo.SetWindowTextW(str);
	return 0;
}