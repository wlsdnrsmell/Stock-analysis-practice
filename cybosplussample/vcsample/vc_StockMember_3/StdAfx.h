// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__88FAF6D7_51D5_4513_ABA5_1B23E77D5470__INCLUDED_)
#define AFX_STDAFX_H__88FAF6D7_51D5_4513_ABA5_1B23E77D5470__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <comdef.h> // <= 추가한 것임

#import "C:\daishin\cybosplus\cpdib.dll" no_namespace // cybosplus에 들어있는 cpdib.dll 경로

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__88FAF6D7_51D5_4513_ABA5_1B23E77D5470__INCLUDED_)
