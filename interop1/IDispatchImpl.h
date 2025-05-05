
#pragma once

#include <oaidl.h>
#include <mshtml.h>
#include <mshtmhst.h>
#include <map>
#include <vector>
#include <string>
#include <memory>
#include <algorithm>
#include <functional>

// T - data type
template <typename T> struct vectorptrsortpred : public std::binary_function<T, T, bool>
{
public:
	bool operator()(const T& spLhs, const T& spRhs) const { return *spLhs < *spRhs; }
};
template <typename T> struct vectorptrnamesortpred : public std::binary_function<T, T, bool>
{
public:
	vectorptrnamesortpred(const std::wstring& sz):m_szName(sz){}
	bool operator()(const T& spLhs, const T& spRhs) const { ASSERT( !spRhs ); return *spLhs < m_szName; }
protected:
	const std::wstring& m_szName;
};

class IInteropDispatchMemberElement
{
public:
	IInteropDispatchMemberElement():m_nID(-1){}
	IInteropDispatchMemberElement(const std::wstring& szName,const long nID):m_nID(nID),m_szName(szName){}
	IInteropDispatchMemberElement(const IInteropDispatchMemberElement& other):IInteropDispatchMemberElement(){*this=other;}
	~IInteropDispatchMemberElement(){}

	long getid( void ) const { return m_nID; }
	const std::wstring& getname( void ) const { return m_szName; }

	bool operator <( const IInteropDispatchMemberElement& other ) const { return m_szName < other.m_szName; }
	bool operator <( const std::wstring& other ) const { return m_szName < other; }
	IInteropDispatchMemberElement& operator =( const IInteropDispatchMemberElement& other ) { m_nID = other.m_nID; m_szName = other.m_szName; return *this; }
protected:
	long m_nID;
	std::wstring m_szName;
};

class IInteropDispatchMemberTarget
{
public:
	IInteropDispatchMemberTarget(){}
	~IInteropDispatchMemberTarget(){}

	virtual bool invoke( const std::vector<CComVariant>& vArgs, CComVariant& retVal ) = 0;
};	

class IInteropDispatchMember
{
public:
	IInteropDispatchMember(){}
	IInteropDispatchMember(std::shared_ptr<const IInteropDispatchMemberElement> spName,const std::vector<std::shared_ptr<const IInteropDispatchMemberElement>>& vParams):m_spName(spName),m_vParams(vParams){sort();}
	IInteropDispatchMember(const IInteropDispatchMember& other):IInteropDispatchMember(){*this=other;}
	~IInteropDispatchMember(){}

	const IInteropDispatchMemberElement *getname( void ) const { return m_spName.get(); }
    bool getnames(DISPID* rgDispId,const UINT cNames) const;
	
	IInteropDispatchMember& operator =( const IInteropDispatchMember& other ) { m_spName = other.m_spName; m_vParams = other.m_vParams; return *this; }
	bool operator <( const IInteropDispatchMember& other ) const { return m_spName ? ( other.m_spName ? *m_spName < *other.m_spName : false ) : true; }
	bool operator <( const std::wstring& other ) const { return m_spName ? ( *m_spName < other ) : true; }
protected:
	std::shared_ptr<const IInteropDispatchMemberElement> m_spName;
	std::vector<std::shared_ptr<const IInteropDispatchMemberElement>> m_vParams;

	void sort( void ) {	if(m_vParams.size()>1)std::sort(m_vParams.begin(), m_vParams.end(), vectorptrsortpred<std::shared_ptr<const IInteropDispatchMemberElement>>());}
};

class IInteropDispatch : public IDispatch
{

public:

	// Constructor/Destructor
	IInteropDispatch();
	~IInteropDispatch();

	// IUnknown
	STDMETHODIMP QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	// IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT* pctinfo);			// provide number of type info interfaces
	STDMETHODIMP GetTypeInfo(/* [in] */ UINT iTInfo,
	/* [in] */ LCID lcid,
	/* [out] */ ITypeInfo** ppTInfo);						// provide type info for an object
	STDMETHODIMP GetIDsOfNames(
	/* [in] */ REFIID riid,
	/* [size_is][in] */ LPOLESTR *rgszNames,
	/* [in] */ UINT cNames,
	/* [in] */ LCID lcid,
	/* [size_is][out] */ DISPID *rgDispId);					// provide int ids for member ( and optionally arguments )
	STDMETHODIMP Invoke(
	/* [in] */ DISPID dispIdMember,
	/* [in] */ REFIID riid,
	/* [in] */ LCID lcid,
	/* [in] */ WORD wFlags,
	/* [out][in] */ DISPPARAMS  *pDispParams,
	/* [out] */ VARIANT  *pVarResult,
	/* [out] */ EXCEPINFO *pExcepInfo,
	/* [out] */ UINT *puArgErr);							// provide access to exposed methods/properties

	// Targets
	bool pushback_target( const std::wstring& szName, IInteropDispatchMemberTarget *p );
	void erase_target( IInteropDispatchMemberTarget* p );
protected:
	ULONG m_n;
	std::vector<std::shared_ptr<const IInteropDispatchMember>> m_vMembers;
	std::map<std::shared_ptr<const IInteropDispatchMember>,std::vector<IInteropDispatchMemberTarget*>> m_mTargets;
	
	std::shared_ptr<const IInteropDispatchMember> findmember(const std::wstring& szName) const;
	std::shared_ptr<const IInteropDispatchMember> findmember(const LPOLESTR *rgszNames,const UINT cNames) const;
	std::shared_ptr<const IInteropDispatchMember> findmember(const DISPID dispIdMember) const;
	
	void sort( void ) {	if(m_vMembers.size()>1)std::sort(m_vMembers.begin(), m_vMembers.end(), vectorptrsortpred<std::shared_ptr<const IInteropDispatchMember>>());}
};

class CInteropControlSite : public COleControlSite
{
public:
	CInteropControlSite(COleControlContainer *pCnt,IInteropDispatch *pIDispatch):COleControlSite(pCnt),m_pIDispatch(pIDispatch){}
protected:

	DECLARE_INTERFACE_MAP();
	BEGIN_INTERFACE_PART(DocHostUIHandler, IDocHostUIHandler)
		STDMETHOD(ShowContextMenu)(/* [in] */ DWORD dwID,
				/* [in] */ POINT __RPC_FAR *ppt,
				/* [in] */ IUnknown __RPC_FAR *pcmdtReserved,
				/* [in] */ IDispatch __RPC_FAR *pdispReserved);
		STDMETHOD(GetHostInfo)( 
				/* [out][in] */ DOCHOSTUIINFO __RPC_FAR *pInfo);
		STDMETHOD(ShowUI)( 
				/* [in] */ DWORD dwID,
				/* [in] */ IOleInPlaceActiveObject __RPC_FAR *pActiveObject,
				/* [in] */ IOleCommandTarget __RPC_FAR *pCommandTarget,
				/* [in] */ IOleInPlaceFrame __RPC_FAR *pFrame,
				/* [in] */ IOleInPlaceUIWindow __RPC_FAR *pDoc);
		STDMETHOD(HideUI)(void);
		STDMETHOD(UpdateUI)(void);
		STDMETHOD(EnableModeless)(/* [in] */ BOOL fEnable);
		STDMETHOD(OnDocWindowActivate)(/* [in] */ BOOL fEnable);
		STDMETHOD(OnFrameWindowActivate)(/* [in] */ BOOL fEnable);
		STDMETHOD(ResizeBorder)( 
				/* [in] */ LPCRECT prcBorder,
				/* [in] */ IOleInPlaceUIWindow __RPC_FAR *pUIWindow,
				/* [in] */ BOOL fRameWindow);
		STDMETHOD(TranslateAccelerator)( 
				/* [in] */ LPMSG lpMsg,
				/* [in] */ const GUID __RPC_FAR *pguidCmdGroup,
				/* [in] */ DWORD nCmdID);
		STDMETHOD(GetOptionKeyPath)( 
				/* [out] */ LPOLESTR __RPC_FAR *pchKey,
				/* [in] */ DWORD dw);
		STDMETHOD(GetDropTarget)(
				/* [in] */ IDropTarget __RPC_FAR *pDropTarget,
				/* [out] */ IDropTarget __RPC_FAR *__RPC_FAR *ppDropTarget);
		STDMETHOD(GetExternal)( 
				/* [out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDispatch);
		STDMETHOD(TranslateUrl)( 
				/* [in] */ DWORD dwTranslate,
				/* [in] */ OLECHAR __RPC_FAR *pchURLIn,
				/* [out] */ OLECHAR __RPC_FAR *__RPC_FAR *ppchURLOut);
		STDMETHOD(FilterDataObject)( 
				/* [in] */ IDataObject __RPC_FAR *pDO,
				/* [out] */ IDataObject __RPC_FAR *__RPC_FAR *ppDORet);
	END_INTERFACE_PART(DocHostUIHandler)

protected:
	IInteropDispatch *m_pIDispatch;
};

class CInteropOccManager : public COccManager
{
public:
	CInteropOccManager(IInteropDispatch * pIDispatch):m_pIDispatch(pIDispatch){}
	virtual COleControlSite* CreateSite(COleControlContainer* pCtrlCont, const CControlCreationInfo& creationInfo) override
	{
		if( creationInfo.m_clsid == CLSID_WebBrowser )
		{
			// our IDispatch implementation allows the javascript in the html to make 'external' calls ( we are aware of )
			CInteropControlSite *pSite = new CInteropControlSite(pCtrlCont,m_pIDispatch);
			return pSite;
		}
		return COccManager::CreateSite(pCtrlCont,creationInfo);
	}
protected:
	IInteropDispatch *m_pIDispatch;

};
