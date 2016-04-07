// vc_StockMemberDlg.cpp : implementation file
//

#include "stdafx.h"
#include "vc_StockMember.h"
#include "vc_StockMemberDlg.h"
#include "CpDibEvent.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CVc_StockMemberDlg dialog

CVc_StockMemberDlg::CVc_StockMemberDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CVc_StockMemberDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CVc_StockMemberDlg)
	m_strOut = _T("");
	m_strJongMok = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	m_pEvent = NULL; // <= 추가한 것임
}

// <= 추가한 것임
CVc_StockMemberDlg::~CVc_StockMemberDlg()
{
	try {
		CString strKey;
		CString* pStrBuf = NULL;

		if(NULL != m_pEvent) {
			IConnectionPointContainerPtr pCPC;
			IConnectionPointPtr pCP;
			pCPC = m_CpDibObj;
			pCPC->FindConnectionPoint(__uuidof(_IDibEvents), &pCP);
			pCP->Unadvise(m_pEvent->GetCookie());
			m_pEvent->Destroy();
		}

		m_mapStr.RemoveAll();
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
}
//

void CVc_StockMemberDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CVc_StockMemberDlg)
	DDX_Text(pDX, IDC_EDT_OUT, m_strOut);
	DDX_Text(pDX, IDC_EDT_JONGMOK, m_strJongMok);
	DDV_MaxChars(pDX, m_strJongMok, 7);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CVc_StockMemberDlg, CDialog)
	//{{AFX_MSG_MAP(CVc_StockMemberDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_REQUEST, OnBtnRequest)
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
    ON_MESSAGE(WM_CPDIB_RECEIVED, OnReceived)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CVc_StockMemberDlg message handlers

BOOL CVc_StockMemberDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
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
	
	// TODO: Add extra initialization here
	// <= 추가한 것임
	DWORD dwCookie;
	CRuntimeClass* prcCpDibEvent = NULL;

	try {
		m_CpDibObj.CreateInstance(__uuidof(StockCur));
		
		prcCpDibEvent = RUNTIME_CLASS(CCpDibEvent);
		m_pEvent = (CCpDibEvent *)prcCpDibEvent->CreateObject();
		m_pEvent->SetOwner(GetSafeHwnd());
		IConnectionPointContainerPtr pCPC;
		IConnectionPointPtr pCP;
		
		pCPC = m_CpDibObj;
		pCPC->FindConnectionPoint(__uuidof(_IDibEvents), &pCP);
		IUnknownPtr pUnk = m_pEvent->GetIDispatch(TRUE);
		pCP->Advise(pUnk, &dwCookie);
		m_pEvent->SetCookie(dwCookie);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
	//

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CVc_StockMemberDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CVc_StockMemberDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

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
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CVc_StockMemberDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// <= 추가한 것임
LONG CVc_StockMemberDlg::OnReceived(WPARAM wParam, LPARAM lParam)
{
	try {
		CString strJongMokCode;
		CString strBuf;
		POSITION pos;

		strJongMokCode = (LPCTSTR)(_bstr_t)m_CpDibObj->GetHeaderValue(0);
		m_mapStr[strJongMokCode].Format(_T("%s(%s)\t시간:%04d\t현재가:%d"), (LPCTSTR)(_bstr_t)m_CpDibObj->GetHeaderValue(1), (LPCTSTR)strJongMokCode, (long)m_CpDibObj->GetHeaderValue(3), (long)m_CpDibObj->GetHeaderValue(13));
		
		for (pos = m_mapStr.GetStartPosition(), m_strOut.Empty(); pos != NULL; ) {
			m_mapStr.GetNextAssoc(pos, strJongMokCode, strBuf);
			m_strOut += strBuf + _T("\r\n");
		}

		UpdateData(FALSE);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
	
	return 0;
}

void CVc_StockMemberDlg::OnBtnRequest() 
{
	// TODO: Add your control notification handler code here
	try {
		UpdateData(TRUE);

		if ("" == m_mapStr[m_strJongMok]) {
			m_CpDibObj->SetInputValue(0, (LPCTSTR)m_strJongMok);
			m_CpDibObj->SubscribeLatest();

			m_mapStr[m_strJongMok] = "";
		}
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
}
//
