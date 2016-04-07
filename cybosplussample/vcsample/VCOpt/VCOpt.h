// VCOpt.h : main header file for the VCOPT application
//

#if !defined(AFX_VCOPT_H__BA1C1534_084F_4BDD_96D6_E4EEA5BCCCC1__INCLUDED_)
#define AFX_VCOPT_H__BA1C1534_084F_4BDD_96D6_E4EEA5BCCCC1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "VCOpt_i.h"

/////////////////////////////////////////////////////////////////////////////
// CVCOptApp:
// See VCOpt.cpp for the implementation of this class
//

class CVCOptApp : public CWinApp
{
public:
	CVCOptApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVCOptApp)
	public:
	virtual BOOL InitInstance();
		virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CVCOptApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	BOOL m_bATLInited;
private:
	BOOL InitATL();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VCOPT_H__BA1C1534_084F_4BDD_96D6_E4EEA5BCCCC1__INCLUDED_)
