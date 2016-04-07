// CpDibEvent.cpp : implementation file
//

#include "stdafx.h"
#include "vc_StockMember.h"
#include "CpDibEvent.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCpDibEvent

IMPLEMENT_DYNCREATE(CCpDibEvent, CCmdTarget)

CCpDibEvent::CCpDibEvent()
{
	// <= 추가한 것임
	m_hwndOwner = NULL;
	EnableAutomation();
	//
}

CCpDibEvent::~CCpDibEvent()
{
}

BEGIN_MESSAGE_MAP(CCpDibEvent, CCmdTarget)
	//{{AFX_MSG_MAP(CCpDibEvent)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

// <= 추가한 것임
BEGIN_DISPATCH_MAP(CCpDibEvent, CCmdTarget)
	DISP_FUNCTION(CCpDibEvent, "Received", Received, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CCpDibEvent,CCmdTarget)
	INTERFACE_PART(CCpDibEvent, __uuidof(_IDibEvents), Dispatch)
END_INTERFACE_MAP()
//

/////////////////////////////////////////////////////////////////////////////
// CCpDibEvent message handlers

void CCpDibEvent::Received()
{
	ASSERT(NULL != m_hwndOwner);
	::SendMessage(m_hwndOwner, WM_CPDIB_RECEIVED, 0, 0);
}
