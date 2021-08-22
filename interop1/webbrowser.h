// CWebBrowser2.h  : Declaration of ActiveX Control wrapper class(es) created by Microsoft Visual C++

#pragma once

#include <vector>
#include <afxwin.h>
#include <afxdisp.h>
#include <string>

/////////////////////////////////////////////////////////////////////////////
// CWebBrowser2

class CWebBrowser2 : public CWnd
{
protected:
	DECLARE_DYNCREATE(CWebBrowser2)
public:
	virtual BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle,
						const RECT& rect, CWnd* pParentWnd, UINT nID, 
						CCreateContext* pContext = NULL)
	{ 
		return CreateControl(CLSID_WebBrowser, lpszWindowName, dwStyle, rect, pParentWnd, nID); 
	}

    BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, 
				UINT nID, CFile* pPersist = NULL, BOOL bStorage = FALSE,
				BSTR bstrLicKey = NULL)
	{ 
		return CreateControl(CLSID_WebBrowser, lpszWindowName, dwStyle, rect, pParentWnd, nID,
		pPersist, bStorage, bstrLicKey); 
	}

// Attributes
public:

// Operations
public:
	
	void Navigate(LPCTSTR URL, VARIANT * Flags, VARIANT * TargetFrameName, VARIANT * PostData, VARIANT * Headers)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, URL, Flags, TargetFrameName, PostData, Headers);
	}
	LPDISPATCH get_Document()
	{
		LPDISPATCH result;
		InvokeHelper(0xcb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	bool calljscript( const std::wstring& csFn, const std::vector<CComVariant>& args, CComVariant& retVal );
	bool getjscript( CComPtr<IDispatch>& spDisp );
};
