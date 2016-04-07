#if !defined(AFX_APPLE_H__DB0CABB6_D278_42EF_9453_4126D5378A11__INCLUDED_)
#define AFX_APPLE_H__DB0CABB6_D278_42EF_9453_4126D5378A11__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Apple.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CApple dialog

class CApple : public CDialog
{
// Construction
public:
	CApple(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CApple)
	enum { IDD = IDD_DIALOG1 };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CApple)
	public:
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CApple)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_APPLE_H__DB0CABB6_D278_42EF_9453_4126D5378A11__INCLUDED_)
