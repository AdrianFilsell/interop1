// CWebBrowser2.cpp  : Definition of ActiveX Control wrapper class(es) created by Microsoft Visual C++


#include "pch.h"
#include "webbrowser.h"
#include <MsHTML.h>
#include <WinInet.h>

/////////////////////////////////////////////////////////////////////////////
// CWebBrowser2

IMPLEMENT_DYNCREATE(CWebBrowser2, CWnd)

bool CWebBrowser2::calljscript( const std::wstring& szFn, const std::vector<CComVariant>& args, CComVariant& retVal )
{
	// get the jscript
	CComPtr<IDispatch> spScript;
	if( !getjscript( spScript ) )
		return false;

	// get the member dispid
	CComBSTR bstrMember( szFn.c_str() );
	DISPID dispid = NULL;
	HRESULT hr = spScript->GetIDsOfNames( IID_NULL, &bstrMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
	if( FAILED(hr) )
		return false;

	// setup the args
	DISPPARAMS dispparams;
	memset( &dispparams, 0, sizeof(dispparams) );
	dispparams.cArgs = UINT( args.size() );
	dispparams.rgvarg = new VARIANT[dispparams.cArgs];
	for( int nCount = 0 ; nCount < int( args.size() ) ; ++nCount )
	{
		const CComVariant& src = args[args.size()-nCount-1];
		dispparams.rgvarg[nCount] = src;
	}
	dispparams.cNamedArgs = 0;

	// make the call
	EXCEPINFO excepInfo;
	memset( &excepInfo, 0, sizeof(excepInfo) );
   	CComVariant vaResult;
	UINT nArgErr = (UINT)-1;
	hr = spScript->Invoke( dispid, IID_NULL, 0, DISPATCH_METHOD, &dispparams, &vaResult, &excepInfo, &nArgErr );
	delete [] dispparams.rgvarg;
	if( FAILED(hr) )
		return false;
	retVal = vaResult;
	
	return true;
}

bool CWebBrowser2::getjscript(CComPtr<IDispatch>& spDisp)
{
	// get the html
	CComPtr<IDispatch> idoc = m_pCtrlSite ? get_Document() : NULL;
	if( !idoc )
		return false;
	CComPtr<IHTMLDocument2>	htmldoc;
	HRESULT hr = idoc->QueryInterface( IID_IHTMLDocument2, (void**)&htmldoc );
	if( FAILED(hr) )
		return false;

	// get the jscript
	hr = htmldoc->get_Script( &spDisp );
	ATLASSERT( SUCCEEDED(hr) );
	return SUCCEEDED(hr);
}
