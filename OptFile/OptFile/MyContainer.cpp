#include "StdAfx.h"
#include "MyContainer.h"
#include "MySite.h"
#include "MyGuids.h"
#include "dispids.h"
#include "Server.h"
#include "MyEnumUnknown.h"

CMyContainer::CMyContainer(void) :
	m_pImpIOleContainer(NULL),
	m_pImpIOleInPlaceFrame(NULL),		// frame manager
	m_pImpIMyContainer(NULL),
	m_cRefs(0),							// object reference count
	m_pStart(NULL),
	m_pEnd(NULL),
	m_cObjects(0)
{
}


CMyContainer::~CMyContainer(void)
{
	this->CloseAllObjects();
	// clear the implemented classes
	Utils_DELETE_POINTER(m_pImpIOleContainer);
	Utils_DELETE_POINTER(m_pImpIOleInPlaceFrame);	// frame manager
	Utils_DELETE_POINTER(m_pImpIMyContainer);
}

// close all objects
void CMyContainer::CloseAllObjects()
{
	// clear the object list
	CLink		*	pLink		= this->m_pStart;
	CLink		*	pNext;
	while (NULL != pLink)
	{
		pNext	= pLink->GetNext();
		delete pLink;
		pLink	= pNext;
	}
	// make it official
	this->m_pStart	= NULL;
	this->m_pEnd	= NULL;
	this->m_cObjects= 0;
}

// IUnknown methods
STDMETHODIMP CMyContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IParseDisplayName == riid || IID_IOleContainer == riid)
		*ppv = this->m_pImpIOleContainer;
	else if (IID_IOleWindow == riid || IID_IOleInPlaceUIWindow == riid || IID_IOleInPlaceFrame == riid)
		*ppv = this->m_pImpIOleInPlaceFrame;
	else if (IID_IDispatch == riid || IID_IMyContainer == riid)
		*ppv = this->m_pImpIMyContainer;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyContainer::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyContainer::Release()
{
	ULONG				cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// Initialization
HRESULT CMyContainer::Init()
{
	m_pImpIOleContainer			= new CImpIOleContainer(this);
	m_pImpIOleInPlaceFrame		= new CImpIOleInPlaceFrame(this);		// frame manager
	m_pImpIMyContainer			= new CImpIMyContainer(this);
	return (NULL != this->m_pImpIOleContainer		&&
			NULL != this->m_pImpIOleInPlaceFrame	&&
			NULL != this->m_pImpIMyContainer) ? S_OK : E_OUTOFMEMORY;
}

// virtuals

// OnObjectCreated base class must be called, since it creates the linked list of objects
void CMyContainer::OnObjectCreated(
									LPCTSTR			szObject,
									IUnknown	*	pUnkSite)
{
	CLink			*	pLink		= new CLink(pUnkSite);			// create the new link
	// add to the end of the list
	if (NULL == this->m_pStart)
	{
		// first list element
		this->m_pStart		= pLink;
		this->m_pEnd		= pLink;
		this->m_cObjects	= 1;
	}
	else
	{
		// add to the end of the list
		this->m_pEnd->SetNext(pLink);
		pLink->SetPrev(this->m_pEnd);
		this->m_pEnd		= pLink;
		// increment the list count
		this->m_cObjects++;
	}
}

BOOL CMyContainer::OnPropRequestEdit(
									LPCTSTR			szObject,
									DISPID			dispid)
{
	return TRUE;
}

void CMyContainer::OnPropChanged(
									LPCTSTR			szObject,
									DISPID			dispid)
{
}

void CMyContainer::OnSinkEvent(
									LPCTSTR			ObjectName,
									long			Dispid,
									long			nParams,
									VARIANT		*	Parameters)
{
}

HRESULT CMyContainer::MySetActiveObject(
									IOleInPlaceActiveObject *pActiveObject)
{
	return E_NOTIMPL;
}

// get a named object
BOOL CMyContainer::GetNamedObject(
									LPCTSTR			szObject,
									REFIID			riid,
									LPVOID		*	ppv)
{
	HRESULT				hr;
	CLink			*	pLink;
	IUnknown		*	pUnkSite;
	BOOL				fSuccess	= FALSE;
	*ppv = NULL;
	if (this->GetLinkElement(szObject, &pLink))
	{
		hr = pLink->GetMySite(&pUnkSite);
		if (SUCCEEDED(hr))
		{
			hr = pUnkSite->QueryInterface(riid, ppv);
			fSuccess = SUCCEEDED(hr);
			pUnkSite->Release();
		}
	}
	return fSuccess;
}

// get extended object
HRESULT CMyContainer::GetExtendedObject(
	LPCTSTR			szObject,
	IDispatch	**	ppdispExtended)
{
	return this->GetNamedObject(szObject, IID_IMyExtended, (LPVOID*)ppdispExtended);
}

// get active x object
HRESULT CMyContainer::GetActiveX(
	LPCTSTR			szObject,
	IDispatch	**	ppdispActiveX)
{
	HRESULT				hr = E_FAIL;
	IDispatch		*	pdispExtended = NULL;
	VARIANT				varResult;
	*ppdispActiveX = NULL;
	if (this->GetExtendedObject(szObject, &pdispExtended))
	{
		hr = Utils_InvokePropertyGet(pdispExtended, DISPID_Object, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			if (VT_UNKNOWN == varResult.vt && NULL != varResult.punkVal)
			{
				hr = varResult.punkVal->QueryInterface(IID_IDispatch, (LPVOID*)ppdispActiveX);
			}
			else hr = E_FAIL;
			VariantClear(&varResult);
		}
		pdispExtended->Release();
	}
	return hr;
}


BOOL CMyContainer::GetLinkElement(
									LPCTSTR			szObject,
									CMyContainer::CLink		**	ppLink)
{
	LPTSTR				szName;
	BOOL				fDone		= FALSE;		// completed flag
	CLink			*	pLink		= this->m_pStart;

	*ppLink = NULL;
	while (!fDone && NULL != pLink)
	{
		szName	= NULL;
		pLink->GetObjectName(&szName);
		if (NULL != szName)
		{
			fDone = 0 == lstrcmpi(szName, szObject);
			if (fDone)
			{
				*ppLink = pLink;
			}
			CoTaskMemFree((LPVOID) szName);
		}
		if (!fDone)
			pLink = pLink->GetNext();
	}
	return fDone;
}

CMyContainer::CImpIOleContainer::CImpIOleContainer(CMyContainer * pBackObj) :
	m_pBackObj(pBackObj)
{
}

CMyContainer::CImpIOleContainer::~CImpIOleContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyContainer::CImpIOleContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_pBackObj->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIOleContainer::AddRef()
{
	return this->m_pBackObj->AddRef();
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIOleContainer::Release()
{
	return this->m_pBackObj->Release();
}

// IParseDisplayName Method 
STDMETHODIMP CMyContainer::CImpIOleContainer::ParseDisplayName(
									IBindCtx	*	pbc,     
									LPOLESTR		pszDisplayName, 
									ULONG		*	pchEaten,							
									IMoniker	**	ppmkOut)
{
	return E_NOTIMPL;
}

// IOleContainer Methods
STDMETHODIMP CMyContainer::CImpIOleContainer::EnumObjects(
									DWORD			grfFlags,  //Value specifying what is to be enumerated
									IEnumUnknown **	ppenum)
{
	HRESULT					hr;
	IUnknown			**	ppUnk			= NULL;
	ULONG					i;
	ULONG					cObjects		= this->m_pBackObj->m_cObjects;		// number of contained objects
	CLink				*	pLink			= this->m_pBackObj->m_pStart;
	CMyEnumUnknown		*	pEnum = NULL;

	*ppenum		= NULL;
	if (0 == cObjects || NULL == pLink) return E_FAIL;
	// create local array
	ppUnk	= new IUnknown * [cObjects];
	for (i=0; i<cObjects; i++)
		ppUnk[i]	= NULL;
	i = 0;
	while (i < cObjects && NULL != pLink)
	{
		pLink->GetMySite(&(ppUnk[i]));
		// increment the indices
		i++;
		pLink = pLink->GetNext();
	}
	pEnum = new CMyEnumUnknown(this);
	hr = pEnum->Init(cObjects, ppUnk, 0);
	if (SUCCEEDED(hr)) hr = pEnum->QueryInterface(IID_IEnumUnknown, (LPVOID*)ppenum);
	if (FAILED(hr)) delete pEnum;
//	hr = Utils_CreateEnumUnknown(this, i, ppUnk, ppenum);
	// cleanup
	for (i=0; i<cObjects; i++)
		Utils_RELEASE_INTERFACE(ppUnk[i]);
	delete [] ppUnk;
	return hr;
}

STDMETHODIMP CMyContainer::CImpIOleContainer::LockContainer(
									BOOL			fLock)
{
	return E_NOTIMPL;
}

CMyContainer::CImpIOleInPlaceFrame::CImpIOleInPlaceFrame(CMyContainer * pBackObj) :
	m_pBackObj(pBackObj)
{}

CMyContainer::CImpIOleInPlaceFrame::~CImpIOleInPlaceFrame()
{}

// IUnknown methods
STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return m_pBackObj->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIOleInPlaceFrame::AddRef()
{
	return m_pBackObj->AddRef();
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIOleInPlaceFrame::Release()
{
	return m_pBackObj->Release();
}

// IOleWindow Methods
STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::GetWindow(
									HWND		*	phwnd)
{
	*phwnd = m_pBackObj->GetMainWindow();
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::ContextSensitiveHelp(
									BOOL			fEnterMode)
{
	return S_OK;
}

// IOleInPlaceUIWindow Methods
STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::GetBorder(
									LPRECT			lprectBorder)
{
	HWND			hwndMain;
	RECT			rect;

	hwndMain = m_pBackObj->GetMainWindow();
	// get the rect for the entire frame.
	GetClientRect(hwndMain, &rect);
	CopyRect(lprectBorder, &rect);
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::RequestBorderSpace(
									LPCBORDERWIDTHS pborderwidths)
{
	// always approve the request
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::SetBorderSpace(
									LPCBORDERWIDTHS pborderwidths)
{
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::SetActiveObject(
									IOleInPlaceActiveObject *pActiveObject,
									LPCOLESTR		pszObjName)
{
	this->m_pBackObj->MySetActiveObject(pActiveObject);
	return S_OK;
}

		// IOleInPlaceFrame Methods
STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::InsertMenus(
									HMENU			hmenuShared,                 
									LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::SetMenu(
									HMENU			hmenuShared,     
									HOLEMENU		holemenu,     
									HWND			hwndActiveObject)
{
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::RemoveMenus(
									HMENU			hmenuShared)
{
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::SetStatusText(
									LPCOLESTR		pszStatusText)
{
	return E_FAIL;					// no status bar yet
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::EnableModeless(
									BOOL			fEnable)
{
	return S_OK;					// no modeless dialogs yet
}

STDMETHODIMP CMyContainer::CImpIOleInPlaceFrame::TranslateAccelerator(
									LPMSG			lpmsg,  //Pointer to structure
									WORD			wID)
{
	return S_FALSE;
}

CMyContainer::CImpIMyContainer::CImpIMyContainer(CMyContainer * pBackObj) :
	m_pBackObj(pBackObj)
{
}

CMyContainer::CImpIMyContainer::~CImpIMyContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyContainer::CImpIMyContainer::QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv)
{
	return this->m_pBackObj->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIMyContainer::AddRef()
{
	return this->m_pBackObj->AddRef();
}

STDMETHODIMP_(ULONG) CMyContainer::CImpIMyContainer::Release()
{
	return this->m_pBackObj->Release();
}

// IDispatch methods
STDMETHODIMP CMyContainer::CImpIMyContainer::GetTypeInfoCount(  
									PUINT				pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyContainer::CImpIMyContainer::GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo)
{
	HRESULT				hr;
	ITypeLib		*	pTypeLib;
	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(IID_IMyContainer, ppTInfo);
		pTypeLib->Release();
	}
	return hr;
}

STDMETHODIMP CMyContainer::CImpIMyContainer::GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	hr = this->GetTypeInfo(0, lcid, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyContainer::CImpIMyContainer::Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr)
{
	switch (dispIdMember)
	{
	case DISPID_Container_CreateObject:
		return this->CreateObject(pDispParams, pVarResult);
	case DISPID_Container_OnPropRequestEdit:
		return this->OnPropRequestEdit(pDispParams, pVarResult);
	case DISPID_Container_OnPropChanged:
		return this->OnPropChanged(pDispParams);
	case DISPID_Container_OnSinkEvent:
		return this->OnSinkEvent(pDispParams);
	case DISPID_Container_CloseObject:
		return this->CloseObject(pDispParams);
	case DISPID_Container_AmbientProperty:
		return E_NOTIMPL;
	case DISPID_Container_GetNamedObject:
		return this->GetNamedObject(pDispParams, pVarResult);
	case DISPID_Container_CloseAllObjects:
		this->m_pBackObj->CloseAllObjects();
		return S_OK;
	case DISPID_Container_TryMNemonic:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->TryMNemonic(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyContainer::CImpIMyContainer::CreateObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	TCHAR				szProgID[MAX_PATH];
	TCHAR				szObjectName[MAX_PATH];
	TCHAR				szStorageFile[MAX_PATH];
	CMySite			*	pMySite			= NULL;
	RECT				rc;
	IDispatch		*	pdispExtended;				// extended object
	VARIANTARG			avarg[6];
	IUnknown		*	punkSite;
//	BOOL				fOK;
	
	if (NULL == pVarResult) pVarResult = &varResult;
	VariantInit(pVarResult);
	VariantInit(&varg);
	// retrieve the parameters
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szProgID, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szObjectName, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 2, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szStorageFile, MAX_PATH, varg.bstrVal);
	hr = DispGetParam(pDispParams, 3, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.left		= varg.lVal;
	hr = DispGetParam(pDispParams, 4, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.top		= varg.lVal;
	hr = DispGetParam(pDispParams, 5, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.right	= varg.lVal;
	hr = DispGetParam(pDispParams, 6, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.bottom	= varg.lVal;
	pMySite		= new CMySite(szProgID, this->m_pBackObj, szObjectName);
	if (NULL == pMySite) return E_OUTOFMEMORY;
	hr = pMySite->Init();
	if (FAILED(hr)) return hr;
	hr = pMySite->QueryInterface(IID_IMyExtended, (LPVOID*) &pdispExtended);
	if (SUCCEEDED(hr))
	{
		InitVariantFromInt32((long) this->m_pBackObj->GetMainWindow(), &avarg[5]);
		InitVariantFromString(szStorageFile, &avarg[4]);
		InitVariantFromInt32(rc.left, &avarg[3]);
		InitVariantFromInt32(rc.top, &avarg[2]);
		InitVariantFromInt32(rc.right, &avarg[1]);
		InitVariantFromInt32(rc.bottom, &avarg[0]);
		hr = Utils_InvokeMethod(pdispExtended, DISPID_Extended_DoCreateObject, avarg, 6, pVarResult);
		VariantClear(&avarg[4]);
		if (SUCCEEDED(hr))
		{
			VariantToBoolean(*pVarResult, &fSuccess);
//			fSuccess = Utils_VariantToBool(pVarResult, &fOK);
			if (fSuccess)
			{
				hr = pdispExtended->QueryInterface(IID_IUnknown, (LPVOID*) &punkSite);
				if (SUCCEEDED(hr))
				{
					this->m_pBackObj->OnObjectCreated(szObjectName, punkSite);
					punkSite->Release();
				}
			}
		}
		pdispExtended->Release();
	}
	return hr;
}

HRESULT CMyContainer::CImpIMyContainer::OnPropRequestEdit(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VARIANT				varResult;
	TCHAR				szObject[MAX_PATH];
	DISPID				dispid;
	BOOL				fAllow		= TRUE;
	VariantInit(&varg);
	if (NULL == pVarResult) pVarResult = &varResult;
	VariantInit(pVarResult);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szObject, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	dispid = (DISPID) varg.lVal;
	fAllow = this->m_pBackObj->OnPropRequestEdit(szObject, dispid);
	pVarResult->vt		= VT_BOOL;
	pVarResult->boolVal	= fAllow ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

HRESULT CMyContainer::CImpIMyContainer::OnPropChanged(
									DISPPARAMS		*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szObject[MAX_PATH];
	DISPID				dispid;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szObject, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	dispid = (DISPID) varg.lVal;
	this->m_pBackObj->OnPropChanged(szObject, dispid);
	return S_OK;
}

HRESULT CMyContainer::CImpIMyContainer::OnSinkEvent(
									DISPPARAMS		*	pDispParams)
{
	HRESULT					hr;
	VARIANTARG				varg;
	UINT					uArgErr;
	TCHAR					szObject[MAX_PATH];
	DISPID					dispid;
	long					uBound;					// safe array size
	VARIANT				*	params;					// array of parameters

	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szObject, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	dispid = (DISPID) varg.lVal;
	if ((VT_ARRAY | VT_VARIANT) == pDispParams->rgvarg[0].vt)
	{
		SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &uBound);
		SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**) &params);
		this->m_pBackObj->OnSinkEvent(szObject, dispid, uBound + 1, params);
		SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
	}
	else
	{
		// no parameter array
		this->m_pBackObj->OnSinkEvent(szObject, dispid, 0, NULL);
	}
	return S_OK;
}

HRESULT CMyContainer::CImpIMyContainer::CloseObject(
									DISPPARAMS		*	pDispParams)
{
	CLink				*	pLink;
	CLink				*	pPrev;				// previous link element
	CLink				*	pNext;				// next link element
	HRESULT					hr;
	VARIANTARG				varg;
	UINT					uArgErr;
	// get the object name
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// find the link element
	if (this->m_pBackObj->GetLinkElement(varg.bstrVal, &pLink))
	{
		// store the previous and next elements
		pPrev	= pLink->GetPrev();
		pNext	= pLink->GetNext();
		// delete the link and decrement the list count
		delete pLink;
		this->m_pBackObj->m_cObjects;
		// recreate the list - 4 cases
		if (NULL != pNext)
		{
			if (NULL != pPrev)
			{
				// somewhere inside the list
				pPrev->SetNext(pNext);
				pNext->SetPrev(pPrev);
			}
			else
			{
				// at the start of the list
				pNext->SetPrev(NULL);
				this->m_pBackObj->m_pStart = pNext;
			}
		}
		else
		{
			// pNext is null
			if (NULL != pPrev)
			{
				// at the end of the list
				pPrev->SetNext(NULL);
				this->m_pBackObj->m_pEnd = pPrev;
			}
			else
			{
				// there is no list, make it official
				this->m_pBackObj->m_pStart		= NULL;
				this->m_pBackObj->m_pEnd		= NULL;
				this->m_pBackObj->m_cObjects	= 0;
			}
		}
	}
	else hr = E_FAIL;
	// clear the variant
	VariantClear(&varg);
	return hr;
}

HRESULT CMyContainer::CImpIMyContainer::GetNamedObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult)
{
	HRESULT					hr;
	VARIANTARG				varg;
	UINT					uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;			// need a result
	VariantInit(pVarResult);
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	pVarResult->vt		= VT_DISPATCH;
	hr = this->m_pBackObj->GetNamedObject(varg.bstrVal, IID_IMyExtended, (LPVOID*) &pVarResult->pdispVal) ? S_OK :E_FAIL;
	VariantClear(&varg);
	return hr;
}

HRESULT CMyContainer::CImpIMyContainer::TryMNemonic(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT				hr;
	IUnknown		*	pUnkSite;
	IDispatch		*	pdispExtended;
	BOOL				fSuccess = FALSE;
	CLink			*	pLink = this->m_pBackObj->m_pStart;
	VARIANT				varResult;
	while (NULL != pLink && !fSuccess)
	{
		hr = pLink->GetMySite(&pUnkSite);
		if (SUCCEEDED(hr))
		{
			hr = pUnkSite->QueryInterface(IID_IMyExtended, (LPVOID*)&pdispExtended);
			if (SUCCEEDED(hr))
			{
				hr = pdispExtended->Invoke(DISPID_Extended_TryMNemonic, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, pDispParams, &varResult, NULL, NULL);
				if (SUCCEEDED(hr))VariantToBoolean(varResult, &fSuccess);
				pdispExtended->Release();
			}
			pUnkSite->Release();
		}
		if (!fSuccess) pLink = pLink->GetNext();
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}


CMyContainer::CLink::CLink(IUnknown * punkSite) :
	m_pUnkSite(NULL),
	m_pPrev(NULL),
	m_pNext(NULL)
{
	if (NULL !=punkSite)
	{
		this->m_pUnkSite	= punkSite;
		this->m_pUnkSite->AddRef();
	}
}

CMyContainer::CLink::~CLink()
{
	this->CloseObject();			// close the activeX
	Utils_RELEASE_INTERFACE(this->m_pUnkSite);	// clear the site
}

// close the object
void CMyContainer::CLink::CloseObject()
{
	HRESULT					hr;
	IUnknown			*	pUnkSite = NULL;
	IDispatch			*	pdispMyExtended	= NULL;
	hr = this->GetMySite(&pUnkSite);
	if (SUCCEEDED(hr))
	{
		hr = pUnkSite->QueryInterface(IID_IMyExtended, (LPVOID*) &pdispMyExtended);
		pUnkSite->Release();
	}
	if (SUCCEEDED(hr))
	{
		Utils_InvokeMethod(pdispMyExtended, DISPID_Extended_DoCloseObject, NULL, 0, NULL);
		pdispMyExtended->Release();
	}
}

HRESULT CMyContainer::CLink::GetMySite(
									IUnknown	**	ppUnkSite)
{
	if (NULL != this->m_pUnkSite)
	{
		*ppUnkSite	= this->m_pUnkSite;
		this->m_pUnkSite->AddRef();
		return S_OK;
	}
	else
	{
		*ppUnkSite	= NULL;
		return E_UNEXPECTED;
	}
}

// get the assigned object name
void CMyContainer::CLink::GetObjectName(
									LPTSTR		*	szName)
{
	HRESULT				hr;
	IUnknown		*	pUnkSite = NULL;					// site object
	IDispatch		*	pdispMyExtended = NULL;			// extended interface
	VARIANT				varResult;
	*szName		= NULL;
	hr = this->GetMySite(&pUnkSite);
	if (SUCCEEDED(hr))
	{
		hr = pUnkSite->QueryInterface(IID_IMyExtended, (LPVOID*) &pdispMyExtended);
		pUnkSite->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokePropertyGet(pdispMyExtended, DISPID_Name, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
//			Utils_VariantToString(&varResult, szName);
			VariantToStringAlloc(varResult, szName);
			VariantClear(&varResult);
		}
		pdispMyExtended->Release();
	}
}

CMyContainer::CLink * CMyContainer::CLink::GetPrev()
{
	return this->m_pPrev;
}

void CMyContainer::CLink::SetPrev(
									CMyContainer::CLink		*	pPrev)
{
	this->m_pPrev = pPrev;
}

CMyContainer::CLink* CMyContainer::CLink::GetNext()
{
	return this->m_pNext;
}

void CMyContainer::CLink::SetNext(
									CMyContainer::CLink		*	pNext)
{
	this->m_pNext = pNext;
}