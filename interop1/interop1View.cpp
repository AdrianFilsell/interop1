
// interop1View.cpp : implementation of the Cinterop1View class
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS can be defined in an ATL project implementing preview, thumbnail
// and search filter handlers and allows sharing of document code with that project.
#ifndef SHARED_HANDLERS
#include "interop1.h"
#endif

#include "interop1Doc.h"
#include "interop1View.h"

#include "webbrowser.h"
#include "TextFieldDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Cinterop1View

IMPLEMENT_DYNCREATE(Cinterop1View, CView)

BEGIN_MESSAGE_MAP(Cinterop1View, CView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &Cinterop1View::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()

	ON_WM_CREATE()
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_WM_ERASEBKGND()
	ON_WM_TIMER()
END_MESSAGE_MAP()

// Cinterop1View construction/destruction

Cinterop1View::Cinterop1View() noexcept
{
	// TODO: add construction code here
	m_bLoadedHTML = false;
	const LONGLONG ll( reinterpret_cast<LONGLONG>(this) );
	m_csPageID.Format( _T("%lli"), ll );
}

Cinterop1View::~Cinterop1View()
{
}

BOOL Cinterop1View::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

// Cinterop1View drawing

void Cinterop1View::OnDraw(CDC* /*pDC*/)
{
	Cinterop1Doc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;

	// TODO: add draw code for native data here
}


// Cinterop1View printing


void Cinterop1View::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL Cinterop1View::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void Cinterop1View::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void Cinterop1View::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

void Cinterop1View::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void Cinterop1View::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// Cinterop1View diagnostics

#ifdef _DEBUG
void Cinterop1View::AssertValid() const
{
	CView::AssertValid();
}

void Cinterop1View::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

Cinterop1Doc* Cinterop1View::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(Cinterop1Doc)));
	return (Cinterop1Doc*)m_pDocument;
}
#endif //_DEBUG

std::wstring Cinterop1View::gethtmlpath( void ) const
{
	std::wstring sz;

	WCHAR filename[_MAX_PATH];
	HMODULE hModule = AfxGetInstanceHandle();
	const DWORD len = hModule ? GetModuleFileNameW( hModule, filename, sizeof(filename)/sizeof(filename[0] ) ) : 0;
	if( len < 1 )
	{
		ASSERT( false );
		return sz;
	}

	WCHAR drive[_MAX_DRIVE];
	WCHAR dir[_MAX_DIR];
	_wsplitpath_s( filename,drive,_MAX_DRIVE,dir,_MAX_DIR,nullptr,0,nullptr,0);

	sz += drive;
	sz += dir;
	sz += L"interop1.html";
	
	return sz;
}

// Cinterop1View message handlers
int Cinterop1View::OnCreate(LPCREATESTRUCT p)
{
	// call the base class
	const int n = CView::OnCreate(p);
	ASSERT( n == 0 );
	if( n != 0 )
		return n;

	// create the browser and register this as a valid target for external member called "InteropMsg"
	m_spBrowser = std::shared_ptr<CWebBrowser2>(new CWebBrowser2);
	if( m_spBrowser->Create( nullptr, nullptr, WS_VISIBLE|WS_CHILD, CRect(0,0,0,0), this, AFX_IDW_PANE_FIRST, nullptr ) )
	{
		IInteropDispatch *pIDispatch = theApp.getidispatch();
		ASSERT( pIDispatch );
		if( pIDispatch )
			pIDispatch->pushback_target(L"InteropMsg",this);

		m_spBrowser->Navigate( gethtmlpath().c_str(), nullptr, nullptr, nullptr, nullptr );
		SetTimer( IDC_HTML_LOADED, 100, nullptr );
	}

	return n;
}

void Cinterop1View::OnDestroy()
{
	// erase this as a target
	IInteropDispatch *pIDispatch = theApp.getidispatch();
	ASSERT( pIDispatch );
	if( pIDispatch )
		pIDispatch->erase_target(this);

	// call the base class
	CView::OnDestroy();	
}

void Cinterop1View::OnSize(UINT nType, int cx, int cy )
{
	// call the base class
	CView::OnSize(nType,cx,cy);

	// position the browser
	if( m_spBrowser && m_spBrowser->GetSafeHwnd() )
	{
		CRect rcClient;
		GetClientRect(rcClient);
		m_spBrowser->MoveWindow(rcClient);
	}
}

BOOL Cinterop1View::OnEraseBkgnd( CDC *pDC )
{
	return TRUE;
}

void Cinterop1View::OnTimer(UINT_PTR nID)
{
	// call the base class
	CView::OnTimer(nID);

	// check for loaded html page
	switch( nID )
	{
		case IDC_HTML_LOADED:
		{
			CComVariant retVal;
			if( m_spBrowser->calljscript( L"GetPageLoaded", {}, retVal ) )
			{
				ASSERT( retVal.vt == VT_BOOL );
				const bool bLoaded = retVal.boolVal;
				if( bLoaded )
				{
					KillTimer( nID );
					m_bLoadedHTML = true;

					const std::vector<CComVariant> vArgs = { CComVariant(m_csPageID) };
					m_spBrowser->calljscript( L"SetPageID", vArgs, retVal );
				}
			}
		}
		break;
		default:ASSERT( false );
	}
}

bool Cinterop1View::invoke( const std::vector<CComVariant>& vArgs, CComVariant& retVal )
{
	// we at least expect name and page id
	if( vArgs.size() < 2 )
		return false;
	if( vArgs[0].vt != VT_BSTR )
		return false;
	const CComBSTR szName(vArgs[0].bstrVal);
	if( vArgs[1].vt != VT_BSTR )
		return false;
	const CComBSTR szID = vArgs[1].bstrVal;
	if( CString( szID ) != m_csPageID )
		return false; // not meant for us

	// check name
	if( szName == L"btnclick")
	{
		if( m_bLoadedHTML )
		{
			TextFieldDlg dlg(this);
			if( m_spBrowser->calljscript( L"GetBtnText", vArgs, retVal ) &&
				retVal.vt == VT_BSTR )
				dlg.m_csBtnText = retVal.bstrVal;
			if( m_spBrowser->calljscript( L"GetParaText", vArgs, retVal ) &&
				retVal.vt == VT_BSTR )
				dlg.m_csParaText = retVal.bstrVal;
			if( m_spBrowser->calljscript( L"GetForenameText", vArgs, retVal ) &&
				retVal.vt == VT_BSTR )
				dlg.m_csForenameText = retVal.bstrVal;
			if( m_spBrowser->calljscript( L"GetSurnameText", vArgs, retVal ) &&
				retVal.vt == VT_BSTR )
				dlg.m_csSurnameText = retVal.bstrVal;
			
			if( dlg.DoModal() )
			{
				{
					const std::vector<CComVariant> vArgs = { CComVariant(dlg.m_csBtnText) };
					m_spBrowser->calljscript( L"SetBtnText", vArgs, retVal );
				}
				{
					const std::vector<CComVariant> vArgs = { CComVariant(dlg.m_csParaText) };
					m_spBrowser->calljscript( L"SetParaText", vArgs, retVal );
				}
				{
					const std::vector<CComVariant> vArgs = { CComVariant(dlg.m_csForenameText) };
					m_spBrowser->calljscript( L"SetForenameText", vArgs, retVal );
				}
				{
					const std::vector<CComVariant> vArgs = { CComVariant(dlg.m_csSurnameText) };
					m_spBrowser->calljscript( L"SetSurnameText", vArgs, retVal );
				}
			}
		}
		return true;
	}

	return false;
}
