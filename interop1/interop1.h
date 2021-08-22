
// interop1.h : main header file for the interop1 application
//
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"       // main symbols

// std includes
#include <memory>

// Cinterop1App:
// See interop1.cpp for the implementation of this class
//

class IInteropDispatch;
class CInteropOccManager;

class Cinterop1App : public CWinAppEx
{
public:
	Cinterop1App() noexcept;


// Overrides
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// Implementation
	UINT  m_nAppLook;
	BOOL  m_bHiColorIcons;

	virtual void PreLoadState();
	virtual void LoadCustomState();
	virtual void SaveCustomState();

	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()

public:
	IInteropDispatch *getidispatch( void ) const { return m_spIDispatch; }

protected:

	// ole support
	std::shared_ptr<CInteropOccManager> m_spOleCtrlContainerMgr;
	CComPtr<IInteropDispatch> m_spIDispatch;
};

extern Cinterop1App theApp;
