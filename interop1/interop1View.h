
// interop1View.h : interface of the Cinterop1View class
//

#pragma once

#include <memory>
#include "IDispatchImpl.h"

class CWebBrowser2;

class Cinterop1View : public CView, public IInteropDispatchMemberTarget
{
protected: // create from serialization only
	Cinterop1View() noexcept;
	DECLARE_DYNCREATE(Cinterop1View)

// Attributes
public:
	Cinterop1Doc* GetDocument() const;

// Operations
public:

// Overrides
public:
	virtual bool invoke( const std::vector<CComVariant>& vArgs, CComVariant& retVal ) override;
	
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~Cinterop1View();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	
	afx_msg int OnCreate(LPCREATESTRUCT p);
	afx_msg void OnDestroy();
	afx_msg void OnSize(UINT nType, int cx, int cy );
	afx_msg BOOL OnEraseBkgnd( CDC *pDC );
	afx_msg void OnTimer(UINT_PTR nID);
	DECLARE_MESSAGE_MAP()

protected:
	CString m_csPageID;
	bool m_bLoadedHTML;
	std::shared_ptr<CWebBrowser2> m_spBrowser;

	std::wstring gethtmlpath( void ) const;
};

#ifndef _DEBUG  // debug version in interop1View.cpp
inline Cinterop1Doc* Cinterop1View::GetDocument() const
   { return reinterpret_cast<Cinterop1Doc*>(m_pDocument); }
#endif

