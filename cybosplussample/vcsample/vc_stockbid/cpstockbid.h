// cpstockbid.h : main header file for the CPSTOCKBID application
//

#if !defined(AFX_CPSTOCKBID_H__10BF5818_4660_4876_A54D_93548E805496__INCLUDED_)
#define AFX_CPSTOCKBID_H__10BF5818_4660_4876_A54D_93548E805496__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CCpstockbidApp:
// See cpstockbid.cpp for the implementation of this class
//

class CCpstockbidApp : public CWinApp
{
public:
	CCpstockbidApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCpstockbidApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CCpstockbidApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CPSTOCKBID_H__10BF5818_4660_4876_A54D_93548E805496__INCLUDED_)
