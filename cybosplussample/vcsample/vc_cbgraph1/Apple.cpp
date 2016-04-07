// Apple.cpp : implementation file
//

#include "stdafx.h"
#include "test.h"
#include "Apple.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CApple dialog


CApple::CApple(CWnd* pParent /*=NULL*/)
	: CDialog(CApple::IDD, pParent)
{
//	m_pParent = pParent;
//	m_nID= CApple::IDD;
	//{{AFX_DATA_INIT(CApple)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CApple::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CApple)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CApple, CDialog)
	//{{AFX_MSG_MAP(CApple)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CApple message handlers


void CApple::PostNcDestroy() 
{
	// TODO: Add your specialized code here and/or call the base class
	delete this;
//	CDialog::PostNcDestroy();
}

BOOL CApple::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
