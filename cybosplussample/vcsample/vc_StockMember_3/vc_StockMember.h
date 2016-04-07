// vc_StockMember.h : main header file for the VC_STOCKMEMBER application
//

#if !defined(AFX_VC_STOCKMEMBER_H__605D38A2_DE82_46A7_9D3F_A281F152AB2C__INCLUDED_)
#define AFX_VC_STOCKMEMBER_H__605D38A2_DE82_46A7_9D3F_A281F152AB2C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CVc_StockMemberApp:
// See vc_StockMember.cpp for the implementation of this class
//

class CVc_StockMemberApp : public CWinApp
{
public:
	CVc_StockMemberApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVc_StockMemberApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CVc_StockMemberApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VC_STOCKMEMBER_H__605D38A2_DE82_46A7_9D3F_A281F152AB2C__INCLUDED_)
