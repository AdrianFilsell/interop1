
#include "pch.h"
#include "IDispatchImpl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

bool IInteropDispatchMember::getnames(DISPID* rgDispId,const UINT cNames) const
{
	// name + args
	if( !rgDispId || cNames < ( 1 + m_vParams.size() ) )
		return false;
	if( !m_spName )
		return false;
	rgDispId[0] = m_spName->getid();
	auto i = m_vParams.cbegin(), end = m_vParams.cend();
	for( int nParam = 0 ; i != end ; ++i, ++nParam )
		rgDispId[1+nParam] = (*i)->getid();
	return true;
}

IInteropDispatch::IInteropDispatch()
{
    m_n = 0;
	{
		// members
		std::shared_ptr<const IInteropDispatchMemberElement> spName { new IInteropDispatchMemberElement( L"InteropMsg", 0 ) };
		std::shared_ptr<const IInteropDispatchMember> spMember { new IInteropDispatchMember( spName, {} ) };
		m_vMembers.push_back( spMember );
	}
	sort();
}

IInteropDispatch::~IInteropDispatch( void )
{
	ASSERT( m_n == 0 );
}

STDMETHODIMP IInteropDispatch::QueryInterface( REFIID riid, void **ppv )
{
    *ppv = NULL;
    if ( IID_IDispatch == riid )
	{
        *ppv = this;
	}
	if ( NULL != *ppv )
    {
        ((LPUNKNOWN)*ppv)->AddRef();
        return NOERROR;
    }
	return E_NOINTERFACE;
}


STDMETHODIMP_(ULONG) IInteropDispatch::AddRef(void)
{
	ASSERT( m_n >= 0 );
    return ++m_n;
}

STDMETHODIMP_(ULONG) IInteropDispatch::Release(void)
{
	ASSERT( m_n > 0 );
    const long lResult = --m_n;
	if(lResult == 0)
		delete this;
	return lResult;
}

STDMETHODIMP IInteropDispatch::GetTypeInfoCount(UINT* /*pctinfo*/)
{
	return E_NOTIMPL;
}

STDMETHODIMP IInteropDispatch::GetTypeInfo(/* [in] */ UINT /*iTInfo*/,
            /* [in] */ LCID /*lcid*/,
            /* [out] */ ITypeInfo** /*ppTInfo*/)
{
	return E_NOTIMPL;
}

STDMETHODIMP IInteropDispatch::GetIDsOfNames(
            /* [in] */ REFIID riid,
            /* [size_is][in] */ OLECHAR** rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID* rgDispId)
{
	HRESULT hr = ResultFromScode(DISP_E_UNKNOWNNAME);

	// attempt to get the id of the member name and the id's of its params
	std::shared_ptr<const IInteropDispatchMember> spMember = findmember(rgszNames,cNames);
	if( spMember && spMember->getnames(rgDispId,cNames) )
		hr = S_OK;

	return hr;
}

STDMETHODIMP IInteropDispatch::Invoke(
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID /*riid*/,
            /* [in] */ LCID /*lcid*/,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS* pDispParams,
            /* [out] */ VARIANT* pVarResult,
            /* [out] */ EXCEPINFO* /*pExcepInfo*/,
            /* [out] */ UINT* puArgErr)
{
	HRESULT hr = ResultFromScode(DISP_E_MEMBERNOTFOUND);

	// attempt to get the member from name id
	std::shared_ptr<const IInteropDispatchMember> spMember = findmember(dispIdMember);
	if( !spMember )
		return hr;

	hr = S_OK;
		
	if(!( wFlags & DISPATCH_METHOD ))
		return hr;
	
	// targets
	auto m = m_mTargets.find( spMember );
	if( m != m_mTargets.cend() )
	{
		// args for member
		std::vector<CComVariant> args;
		for( int nArg = 0 ; nArg < int( pDispParams->cArgs ) ; ++nArg )
			args.insert( args.begin(), CComVariant( pDispParams->rgvarg[nArg] ) );

		// find the 'correct' target
		CComVariant retVal;
		auto i = (*m).second.begin();
		auto end = (*m).second.cend();
		for( ; i != end ; ++i )
		{
			if( (*i)->invoke( args, retVal ) )
			{
				if( pVarResult )
				{
					HRESULT hr = VariantCopy( pVarResult, &retVal );
				}
				break;
			}
		}
	}
	
	return hr;
}

std::shared_ptr<const IInteropDispatchMember> IInteropDispatch::findmember(const LPOLESTR *rgszNames,const UINT cNames) const
{
	if(!rgszNames || cNames < 1)
		return nullptr;

	// first name is member name then subsequent names are parameters
	const std::wstring szName = rgszNames[0];
	return findmember( szName );
}

std::shared_ptr<const IInteropDispatchMember> IInteropDispatch::findmember(const std::wstring& szName) const
{
	// get first non before iterator using member name
	const int nFrom = 0;
	std::vector<std::shared_ptr<const IInteropDispatchMember>>::const_iterator i;
	if( nFrom >= m_vMembers.size() )
		i = m_vMembers.cend();
	else
		i = std::lower_bound(m_vMembers.cbegin() + nFrom, m_vMembers.cend(), nullptr, vectorptrnamesortpred<std::shared_ptr<const IInteropDispatchMember>>(szName));
	return i != m_vMembers.cend() && (*i)->getname() && (*i)->getname()->getname() == szName ? (*i) : nullptr;
}

std::shared_ptr<const IInteropDispatchMember> IInteropDispatch::findmember(const DISPID dispIdMember) const
{
	// get first non before iterator using member name
	const int nFrom = 0;
	if( nFrom >= m_vMembers.size() )
		return nullptr;
	std::vector<std::shared_ptr<const IInteropDispatchMember>>::const_iterator i = m_vMembers.cbegin() + nFrom, end = m_vMembers.cend();
	for( ; i != end ; ++i )
		if( (*i)->getname() && (*i)->getname()->getid() == dispIdMember )
			return (*i);
	return nullptr;
}

bool IInteropDispatch::pushback_target( const std::wstring& szName, IInteropDispatchMemberTarget* p )
{
	// associate target with member
	std::shared_ptr<const IInteropDispatchMember> spMember = findmember(szName);
	if( !p || !spMember )
		return false;
	auto i = m_mTargets.find( spMember );
	if( i == m_mTargets.cend() )
		i = m_mTargets.insert( std::make_pair(spMember,std::vector<IInteropDispatchMemberTarget*>() ) ).first;
	(*i).second.push_back( p );
	return true;
}

void IInteropDispatch::erase_target( IInteropDispatchMemberTarget* p )
{
	// erase this target from every member
	auto m = m_mTargets.begin();
	for( ; m != m_mTargets.cend() ; )
	{
		auto nPreVectorSize = (*m).second.size();
		auto i = (*m).second.begin(); 
		for( ; i != (*m).second.cend() ; )
			if( (*i) == p )
			{
				ASSERT( (*m).second.size() == nPreVectorSize ); // should not be multiple instances of the same target registered for the same member but carry on with the subsequent erase for completeness
				i = (*m).second.erase(i);
			}
			else
				++i;
		if( (*m).second.size() == 0 )
			m = m_mTargets.erase(m);
		else
			++m;
	}
}

PROCESS_LOCAL(CInteropOccManager,_afxOccManager)

PROCESS_LOCAL(CControlSiteFactoryMgr,_afxControlFactoryMgr)

BEGIN_INTERFACE_MAP(CInteropControlSite, COleControlSite)
	INTERFACE_PART(CInteropControlSite, IID_IDocHostUIHandler, DocHostUIHandler)
END_INTERFACE_MAP()

ULONG FAR EXPORT  CInteropControlSite::XDocHostUIHandler::AddRef()
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)	// METHOD_PROLOGUE ensures pThis correct
	return pThis->ExternalAddRef();
}

ULONG FAR EXPORT  CInteropControlSite::XDocHostUIHandler::Release()
{                            
    METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)	// METHOD_PROLOGUE ensures pThis correct
	return pThis->ExternalRelease();
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::QueryInterface(REFIID riid, void **ppvObj)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)	// METHOD_PROLOGUE ensures pThis correct
    HRESULT hr = (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
	return hr;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::GetHostInfo( DOCHOSTUIINFO* pInfo )
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)	// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::ShowUI(
				DWORD dwID, 
				IOleInPlaceActiveObject * /*pActiveObject*/,
				IOleCommandTarget * pCommandTarget,
				IOleInPlaceFrame * /*pFrame*/,
				IOleInPlaceUIWindow * /*pDoc*/)
{

	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::HideUI(void)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::UpdateUI(void)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::EnableModeless(BOOL /*fEnable*/)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::OnDocWindowActivate(BOOL /*fActivate*/)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::OnFrameWindowActivate(BOOL /*fActivate*/)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::ResizeBorder(
				LPCRECT /*prcBorder*/, 
				IOleInPlaceUIWindow* /*pUIWindow*/,
				BOOL /*fRameWindow*/)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::ShowContextMenu(
				DWORD /*dwID*/, 
				POINT* /*pptPosition*/,
				IUnknown* /*pCommandTarget*/,
				IDispatch* /*pDispatchObjectHit*/)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::TranslateAccelerator(LPMSG lpMsg,
            /* [in] */ const GUID __RPC_FAR *pguidCmdGroup,
            /* [in] */ DWORD nCmdID)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

HRESULT FAR EXPORT  CInteropControlSite::XDocHostUIHandler::GetOptionKeyPath(BSTR* pbstrKey, DWORD)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}

STDMETHODIMP CInteropControlSite::XDocHostUIHandler::GetDropTarget( 
            /* [in] */ IDropTarget __RPC_FAR *pDropTarget,
            /* [out] */ IDropTarget __RPC_FAR *__RPC_FAR *ppDropTarget)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)
	return E_NOTIMPL;
}

STDMETHODIMP CInteropControlSite::XDocHostUIHandler::GetExternal( 
            /* [out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDispatch)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
	if (!ppDispatch)
		return E_POINTER;
	IDispatch* pDisp = pThis->m_pIDispatch;// our IDispatch allows the javascript in our html page to call 'external.InteropMsg'
	pDisp->AddRef();
	*ppDispatch = pDisp;
	return S_OK;
}
        
STDMETHODIMP CInteropControlSite::XDocHostUIHandler::TranslateUrl( 
            /* [in] */ DWORD dwTranslate,
            /* [in] */ OLECHAR __RPC_FAR *pchURLIn,
            /* [out] */ OLECHAR __RPC_FAR *__RPC_FAR *ppchURLOut)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}
        
STDMETHODIMP CInteropControlSite::XDocHostUIHandler::FilterDataObject( 
            /* [in] */ IDataObject __RPC_FAR *pDO,
            /* [out] */ IDataObject __RPC_FAR *__RPC_FAR *ppDORet)
{
	METHOD_PROLOGUE(CInteropControlSite, DocHostUIHandler)// METHOD_PROLOGUE ensures pThis correct
    return E_NOTIMPL;
}
