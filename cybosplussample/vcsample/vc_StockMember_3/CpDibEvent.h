#if !defined(AFX_CPDIBEVENT_H__B591F279_D708_4C0B_BF79_4B22D8F69845__INCLUDED_)
#define AFX_CPDIBEVENT_H__B591F279_D708_4C0B_BF79_4B22D8F69845__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// CpDibEvent.h : header file
//

#define WM_CPDIB_RECEIVED	(WM_USER+100) // <= �߰��� ����

/////////////////////////////////////////////////////////////////////////////
// CCpDibEvent command target

class CCpDibEvent : public CCmdTarget
{
	DECLARE_DYNCREATE(CCpDibEvent)

	CCpDibEvent();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	// <= �߰��� ����
	void	Destroy() { delete this; }
	void	Received();
	void	SetOwner(HWND hwnd) { ASSERT(NULL != hwnd); m_hwndOwner = hwnd; }
	DWORD	GetCookie() { return m_dwCookie; }
	void	SetCookie(DWORD dwCookie) { m_dwCookie = dwCookie; }
	//

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCpDibEvent)
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CCpDibEvent();

	// Generated message map functions
	//{{AFX_MSG(CCpDibEvent)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	// <= �߰��� ����
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
	//
private:
	// <= �߰��� ����
	HWND	m_hwndOwner;	
	DWORD	m_dwCookie;
	//
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CPDIBEVENT_H__B591F279_D708_4C0B_BF79_4B22D8F69845__INCLUDED_)
