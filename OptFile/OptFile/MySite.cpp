#include "StdAfx.h"
#include "MySite.h"
#include "MyGuids.h"
#include "dispids.h"
#include "Server.h"

CMySite::CMySite(LPCTSTR szProgID, IUnknown * pUnkContainer, LPCTSTR szName) :
	m_pImpIAmbients(NULL),
	m_pImpISink(NULL),
	m_pImpIExtended(NULL),
	m_pImpIPropertyNotifySink(NULL),
	m_pImpIOleClientSite(NULL),
	m_pImpIOleInPlaceSite(NULL),
	m_pImpIOleControlSite(NULL),
	// object reference count
	m_cRefs(0),
	// active X object
	m_pUnk(NULL),
	m_pStorage(NULL),
	// assigned object name
	m_szName(NULL),
	// container
	m_pUnkContainer(NULL),
	// sink interface id
	m_iidSink(IID_NULL),
	m_dwCookie(0),		// connection cookie
	m_dwPropNotify(0),		// property notify cookie
	// object class ID
	m_clsid(CLSID_NULL)

{
	HRESULT				hr;
	SetRectEmpty(&m_rcPos);
	ZeroMemory((PVOID) &m_controlInfo, sizeof(CONTROLINFO));
	// store reference to the container
	this->m_pUnkContainer	= pUnkContainer;
	if (NULL != this->m_pUnkContainer) this->m_pUnkContainer->AddRef();
	// store the object name
//	Utils_DoCopyString(&(this->m_szName), szName);
	SHStrDup(szName, &(this->m_szName));
	// obtain the clsid
	hr = CLSIDFromProgID(szProgID, &m_clsid);
}


CMySite::~CMySite(void)
{
	// disconnect the sinks
	if (NULL != this->m_pUnk)
	{
		Utils_ConnectToConnectionPoint(this->m_pUnk, NULL, this->m_iidSink, FALSE, &this->m_dwCookie);
		Utils_ConnectToConnectionPoint(this->m_pUnk, NULL, IID_IPropertyNotifySink, FALSE, &this->m_dwPropNotify);
	}
	Utils_DELETE_POINTER(m_pImpIAmbients);
	Utils_DELETE_POINTER(m_pImpISink);
	Utils_DELETE_POINTER(m_pImpIExtended);
	Utils_DELETE_POINTER(m_pImpIPropertyNotifySink);
	Utils_DELETE_POINTER(m_pImpIOleClientSite);
	Utils_DELETE_POINTER(m_pImpIOleInPlaceSite);
	Utils_DELETE_POINTER(m_pImpIOleControlSite);
	Utils_RELEASE_INTERFACE(this->m_pUnkContainer);
	Utils_RELEASE_INTERFACE(this->m_pUnk);
	Utils_RELEASE_INTERFACE(this->m_pStorage);
//	Utils_DoCopyString(&m_szName, NULL);
	CoTaskMemFree((LPVOID) this->m_szName);
	this->m_szName = NULL;
}

// IUnknown methods
STDMETHODIMP CMySite::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid)
		*ppv = this->m_pImpIAmbients;
	else if (riid == this->m_iidSink)
		*ppv = this->m_pImpISink;
	else if (IID_IMyExtended == riid)
		*ppv = this->m_pImpIExtended;
	else if (IID_IPropertyNotifySink == riid)
		*ppv = this->m_pImpIPropertyNotifySink;
	else if (IID_IOleClientSite == riid)
		*ppv = this->m_pImpIOleClientSite;
	else if (IID_IOleWindow == riid || IID_IOleInPlaceSite == riid)
		*ppv = this->m_pImpIOleInPlaceSite;
	else if (IID_IOleControlSite == riid)
		*ppv = this->m_pImpIOleControlSite;
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

STDMETHODIMP_(ULONG) CMySite::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMySite::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// initialization
HRESULT	CMySite::Init()
{
	m_pImpIAmbients				= new CImpIAmbients(this);
	m_pImpISink					= new CImpISink(this);
	m_pImpIExtended				= new CImpIExtended(this);
	m_pImpIPropertyNotifySink	= new CImpIPropertyNotifySink(this);
	m_pImpIOleClientSite		= new CImpIOleClientSite(this);
	m_pImpIOleInPlaceSite		= new CImpIOleInPlaceSite(this);
	m_pImpIOleControlSite		= new CImpIOleControlSite(this);
	return S_OK;
}

void CMySite::GetControlInfo()
{
	HRESULT			hr;
	IOleControl	*	pOleControl;

	hr = this->m_pUnk->QueryInterface(IID_IOleControl, (LPVOID*)&pOleControl);
	if (SUCCEEDED(hr))
	{
		pOleControl->GetControlInfo(&(this->m_controlInfo));
		pOleControl->Release();
	}
}

BOOL CMySite::TryNMemonic(
	LPMSG			pmsg)
{
	BYTE			fVirt = FVIRTKEY;
	UINT			i;
	LPACCEL			pAccel = NULL;;
	BOOL			fSuccess = FALSE;
	HRESULT			hr;
	IOleControl*	pOleControl;

	if (0 == this->m_controlInfo.cAccel) return FALSE;
	pAccel = new ACCEL[this->m_controlInfo.cAccel];
	if (CopyAcceleratorTable(this->m_controlInfo.hAccel, pAccel, this->m_controlInfo.cAccel) > 0)
	{
		hr = this->m_pUnk->QueryInterface(IID_IOleControl, (LPVOID*)&pOleControl);
		if (SUCCEEDED(hr))
		{
			// form virtual key code
			fVirt |= (WM_SYSKEYDOWN == pmsg->message) ? FALT : 0;
			fVirt |= (0x8000 & GetKeyState(VK_CONTROL)) ? FCONTROL : 0;
			fVirt |= (0x8000 & GetKeyState(VK_SHIFT)) ? FSHIFT : 0;
			i = 0;
			while (i < this->m_controlInfo.cAccel && !fSuccess)
			{
				if (pAccel[i].key == pmsg->wParam && pAccel[i].fVirt == fVirt)
				{
					fSuccess = TRUE;
					pOleControl->OnMnemonic(pmsg);
				}
				if (!fSuccess) i++;
			}
			pOleControl->Release();
		}
	}
	delete[] pAccel;
	return fSuccess;
}


CMySite::CImpIOleClientSite::CImpIOleClientSite(CMySite * pMySite) :
	m_pMySite(pMySite)
{}

CMySite::CImpIOleClientSite::~CImpIOleClientSite()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIOleClientSite::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleClientSite::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleClientSite::Release()
{
	return this->m_pMySite->Release();
}

// IOleClientSite methods
STDMETHODIMP CMySite::CImpIOleClientSite::SaveObject()
{
	IPersistStorage		*	pIPersistStorage;
	HRESULT					hr;

	// get a pointer to IPersistStorage
	hr = this->m_pMySite->m_pUnk->QueryInterface(IID_IPersistStorage, (LPVOID*) &pIPersistStorage);
	if (SUCCEEDED(hr))
	{
		hr = OleSave(pIPersistStorage, this->m_pMySite->m_pStorage, TRUE);
		pIPersistStorage->SaveCompleted(NULL);
		pIPersistStorage->Release();
	}
	else
		hr = E_FAIL;
	return hr;
}

STDMETHODIMP CMySite::CImpIOleClientSite::GetMoniker(
									DWORD			dwAssign,  
									DWORD			dwWhichMoniker,
									IMoniker	**	ppmk)
{
	*ppmk	= NULL;
    return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpIOleClientSite::GetContainer(
									LPOLECONTAINER* ppContainer)
{
	return this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleContainer, (LPVOID*) ppContainer);
}

STDMETHODIMP CMySite::CImpIOleClientSite::ShowObject()
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleClientSite::OnShowWindow(
									BOOL			fShow)
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleClientSite::RequestNewObjectLayout()
{
	return E_NOTIMPL;
}

CMySite::CImpIOleInPlaceSite::CImpIOleInPlaceSite(CMySite * pMySite) :
	m_pMySite(pMySite)
{}

CMySite::CImpIOleInPlaceSite::~CImpIOleInPlaceSite()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIOleInPlaceSite::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleInPlaceSite::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleInPlaceSite::Release()
{
	return this->m_pMySite->Release();
}

// IOleWindow Methods
STDMETHODIMP CMySite::CImpIOleInPlaceSite::GetWindow(
									HWND		*	phwnd)
{
	HRESULT					hr;
	IOleWindow			*	pOleWindow;

	*phwnd	= NULL;
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		hr = pOleWindow->GetWindow(phwnd);
		pOleWindow->Release();
	}
	return hr;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::ContextSensitiveHelp(
							BOOL			fEnterMode)
{
	HRESULT					hr;
	IOleWindow			*	pOleWindow;

	hr = this->m_pMySite->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		hr = pOleWindow->ContextSensitiveHelp(fEnterMode);
		pOleWindow->Release();
	}
	return hr;
}

// IOleInPlaceSite Methods
STDMETHODIMP CMySite::CImpIOleInPlaceSite::CanInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::OnInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::OnUIActivate()
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::GetWindowContext(
									IOleInPlaceFrame **ppFrame,
									IOleInPlaceUIWindow **ppDoc,
									LPRECT			lprcPosRect,  //Points to position of in-place object
									LPRECT			lprcClipRect, //Points to in-place object's position 
																	 // rectangle
									LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	HRESULT				hr;
	RECT				rc;					// object rectangle
	SIZEL				szl;				// control size
	HWND				hwndMain	= NULL;			// main window
	IOleObject		*	pOleObject;			// controls IOleObject
	IOleWindow		*	pOleWindow;

	// in place frame 
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleInPlaceFrame, (LPVOID*) ppFrame);
	*ppDoc		= NULL;						// using SDI
	// control rectangle
	if (IsRectEmpty(&(this->m_pMySite->m_rcPos)))
	{
		hr = this->m_pMySite->m_pUnk->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
		if (SUCCEEDED(hr))
		{
			hr = pOleObject->GetExtent(DVASPECT_CONTENT, &szl);
			pOleObject->Release();
		}
		if (FAILED(hr))
		{
			// the AcitveX control isn;t running - use the ViewObject to obtain the size
			IViewObject2			*	pIViewObject2;

			hr =  this->m_pMySite->m_pUnk->QueryInterface(IID_IViewObject2, (LPVOID*) &pIViewObject2);
			if (SUCCEEDED(hr))
			{
				hr = pIViewObject2->GetExtent(DVASPECT_CONTENT, 0, NULL, &szl);
				pIViewObject2->Release();
			}
		}
		if (FAILED(hr)) return hr;

		// convert it to pixels
		Utils_HimetricToPixels(&szl);
		lprcPosRect->left = 0;
		lprcPosRect->top = 0;
		lprcPosRect->right = szl.cx;
		lprcPosRect->bottom = szl.cy;
//		rc.left	= rc.top = 0;
//		rc.right	= szl.cx;
//		rc.bottom	= szl.cy;
//		Utils_HiMetricToPixels(&rc);
//
//		CopyMemory((PVOID) lprcPosRect, (const void*) &rc, sizeof(RECT));
		if (FAILED(hr)) return hr;
	}
	else
		CopyRect(lprcPosRect, &(this->m_pMySite->m_rcPos));
	// get the main window
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		pOleWindow->GetWindow(&hwndMain);
		pOleWindow->Release();
	}
	if (FAILED(hr)) return hr;
	GetClientRect(hwndMain, &rc);
	CopyMemory((PVOID) lprcClipRect, (const void*) &rc, sizeof(RECT));
	// fill the FRAMEINFO
	lpFrameInfo->fMDIApp		= FALSE;
	lpFrameInfo->hwndFrame		= hwndMain;
	lpFrameInfo->haccel			= NULL;
	lpFrameInfo->cAccelEntries	= 0;

	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::Scroll(
									SIZE			scrollExtent)
{
	return E_FAIL;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::OnUIDeactivate(
									BOOL			fUndoable)
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::OnInPlaceDeactivate()
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::DiscardUndoState()
{
	return E_FAIL;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::DeactivateAndUndo()
{
	return E_FAIL;
}

STDMETHODIMP CMySite::CImpIOleInPlaceSite::OnPosRectChange(
									LPCRECT			lprcPosRect)
{
	// update the size in the document object
	HWND					hwndMain		= NULL;
	IOleWindow			*	pOleWindow;
	RECT					rect;
	HRESULT					hr;
	IOleInPlaceObject	*	pIOleInPlaceObject;

	// get the main window
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		pOleWindow->GetWindow(&hwndMain);
		pOleWindow->Release();
	}
	if (FAILED(hr)) return hr;
	GetClientRect(hwndMain, &rect);
	// tell the object its new size
	hr =this->m_pMySite->m_pUnk->QueryInterface(IID_IOleInPlaceObject, (LPVOID*) &pIOleInPlaceObject);
	if (SUCCEEDED(hr))
	{
        pIOleInPlaceObject->SetObjectRects(lprcPosRect, &rect);
		pIOleInPlaceObject->Release();
	}
	return hr;
}

CMySite::CImpIOleControlSite::CImpIOleControlSite(CMySite * pMySite)  :
	m_pMySite(pMySite)
{}

CMySite::CImpIOleControlSite::~CImpIOleControlSite()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIOleControlSite::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleControlSite::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIOleControlSite::Release()
{
	return this->m_pMySite->Release();
}

// IOleControlSite Methods
STDMETHODIMP CMySite::CImpIOleControlSite::OnControlInfoChanged(void)
{
	this->m_pMySite->GetControlInfo();
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleControlSite::LockInPlaceActive(
									BOOL			fLock)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpIOleControlSite::GetExtendedControl(
									IDispatch	**	ppDisp)
{
	return this->m_pMySite->QueryInterface(IID_IMyExtended, (LPVOID*) ppDisp);
}

STDMETHODIMP CMySite::CImpIOleControlSite::TransformCoords(
									POINTL		*	pPtlHimetric ,  //Address of POINTL structure
									POINTF		*	pPtfContainer ,  //Address of POINTF structure
									DWORD			dwFlags)
{
	HRESULT					hr = S_OK;
	HWND					hwndMain	= NULL;				// main window
	HDC						hdc;					// device context
	POINT					rgptConvert[2];
	IOleWindow			*	pOleWindow;

	// get the main window
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		hr = pOleWindow->GetWindow(&hwndMain);
		pOleWindow->Release();
	}
	hdc			= GetDC(hwndMain);
	SetMapMode(hdc, MM_HIMETRIC);					// map mode to HIMETRIC
	rgptConvert[0].x = 0;
	rgptConvert[0].y = 0;
	if (dwFlags & XFORMCOORDS_HIMETRICTOCONTAINER)
	{
		rgptConvert[1].x = pPtlHimetric->x;
		rgptConvert[1].y = pPtlHimetric->y;
		LPtoDP(hdc, rgptConvert, 2);
		if (dwFlags & XFORMCOORDS_SIZE)
		{
			pPtfContainer->x = (float)(rgptConvert[1].x - rgptConvert[0].x);
			pPtfContainer->y = (float)(rgptConvert[0].y - rgptConvert[1].y);
		}
		else if (dwFlags & XFORMCOORDS_POSITION)
		{
			pPtfContainer->x = (float)rgptConvert[1].x;
			pPtfContainer->y = (float)rgptConvert[1].y;
		}
		else
		{
			hr = E_INVALIDARG;
		}
	}
	else if (dwFlags & XFORMCOORDS_CONTAINERTOHIMETRIC)
	{
		rgptConvert[1].x = (int)(pPtfContainer->x);
		rgptConvert[1].y = (int)(pPtfContainer->y);
		DPtoLP(hdc, rgptConvert, 2);
		if (dwFlags & XFORMCOORDS_SIZE)
		{
			pPtlHimetric->x = rgptConvert[1].x - rgptConvert[0].x;
			pPtlHimetric->y = rgptConvert[0].y - rgptConvert[1].y;
		}
		else if (dwFlags & XFORMCOORDS_POSITION)
		{
			pPtlHimetric->x = rgptConvert[1].x;
			pPtlHimetric->y = rgptConvert[1].y;
		}
		else
		{
			hr = E_INVALIDARG;
		}
	}
	else
	{
		hr = E_INVALIDARG;
	}
	// release the device context
	ReleaseDC(hwndMain, hdc);
	return hr;
}

STDMETHODIMP CMySite::CImpIOleControlSite::TranslateAccelerator(
									LPMSG			pMsg ,        //Pointer to the structure
									DWORD			grfModifiers)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpIOleControlSite::OnFocus(
									BOOL			fGotFocus)
{
	return S_OK;
}

STDMETHODIMP CMySite::CImpIOleControlSite::ShowPropertyFrame(void)
{
	return E_NOTIMPL;
}

CMySite::CImpIAmbients::CImpIAmbients(CMySite * pMySite) :
	m_pMySite(pMySite)
{
}

CMySite::CImpIAmbients::~CImpIAmbients()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIAmbients::QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIAmbients::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIAmbients::Release()
{
	return this->m_pMySite->Release();
}

// IDispatch methods
STDMETHODIMP CMySite::CImpIAmbients::GetTypeInfoCount(  
									PUINT				pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CMySite::CImpIAmbients::GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpIAmbients::GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpIAmbients::Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr)
{
	HRESULT					hr;
	IDispatch			*	pdisp;
	VARIANTARG				varg;
	// only allow property get
	if (DISPATCH_PROPERTYGET != wFlags) return E_INVALIDARG;
	if (NULL == this->m_pMySite->m_pUnkContainer) return E_UNEXPECTED;
	hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		InitVariantFromInt32(dispIdMember, &varg);
		hr = Utils_DoInvoke(pdisp, DISPID_Container_AmbientProperty, DISPATCH_PROPERTYGET, &varg, 1, pVarResult);
		pdisp->Release();
	}
	return hr;
}

CMySite::CImpISink::CImpISink(CMySite * pMySite) :
	m_pMySite(pMySite)
{
}

CMySite::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpISink::QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpISink::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpISink::Release()
{
	return this->m_pMySite->Release();
}

// IDispatch methods
STDMETHODIMP CMySite::CImpISink::GetTypeInfoCount(  
									PUINT				pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CMySite::CImpISink::GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpISink::GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMySite::CImpISink::Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr)
{
	HRESULT					hr;
	IDispatch			*	pdisp;
	VARIANTARG				avarg[3];
	long					i;
	SAFEARRAYBOUND			sab;
	if (NULL != this->m_pMySite->m_pUnkContainer)
	{
		hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
		if (SUCCEEDED(hr))
		{
			InitVariantFromString(this->m_pMySite->m_szName, &avarg[2]);
			InitVariantFromInt32(dispIdMember, &avarg[1]);
	//		Utils_StringToVariant(this->m_pMySite->m_szName, &avarg[2]);
	//		Utils_LongToVariant(dispIdMember, &avarg[1]);
			VariantInit(&avarg[0]);
			if (pDispParams->cArgs > 0)
			{
				avarg[0].vt		= VT_ARRAY | VT_VARIANT;
				sab.lLbound		= 0;
				sab.cElements	= pDispParams->cArgs;
				avarg[0].parray	= SafeArrayCreate(VT_VARIANT, 1, &sab);
				for (i=0; i<pDispParams->cArgs; i++)
					SafeArrayPutElement(avarg[0].parray, &i, (void*) &(pDispParams->rgvarg[i]));
			}
			else
			{
				avarg[0].vt		= VT_ERROR;
				avarg[0].scode	= DISP_E_MEMBERNOTFOUND;
			}
			Utils_InvokeMethod(pdisp, DISPID_Container_OnSinkEvent, avarg, 3, NULL);
			VariantClear(&avarg[2]);
			VariantClear(&avarg[0]);
			pdisp->Release();
		}
	}
	return hr;
}

CMySite::CImpIExtended::CImpIExtended(CMySite * pMySite) :
	m_pMySite(pMySite)
{
}

CMySite::CImpIExtended::~CImpIExtended()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIExtended::QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIExtended::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIExtended::Release()
{
	return this->m_pMySite->Release();
}

// IDispatch methods
STDMETHODIMP CMySite::CImpIExtended::GetTypeInfoCount(  
									PUINT				pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMySite::CImpIExtended::GetTypeInfo(  
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
		hr = pTypeLib->GetTypeInfoOfGuid(IID_IMyExtended, ppTInfo);
		pTypeLib->Release();
	}
	return hr;
}

STDMETHODIMP CMySite::CImpIExtended::GetIDsOfNames(  
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

STDMETHODIMP CMySite::CImpIExtended::Invoke(  
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
	case DISPID_Name:
		if (wFlags == DISPATCH_PROPERTYGET)
		{
			return this->GetName(pVarResult);
		}
		else return DISP_E_MEMBERNOTFOUND;
		break;
	case DISPID_Object:
		if (wFlags == DISPATCH_PROPERTYGET)
		{
			return this->GetObject(pVarResult);
		}
		else return DISP_E_MEMBERNOTFOUND;
		break;
	case DISPID_Extended_DoCreateObject:
		if (wFlags == DISPATCH_METHOD)
		{
			return this->DoCreateObject(pDispParams, pVarResult);
		}
		else return E_INVALIDARG;
		break;
	case DISPID_Extended_DoCloseObject:
		if (wFlags == DISPATCH_METHOD)
		{
			this->DoCloseObject();
		}
		else return E_INVALIDARG;
		break;
	case DISPID_Extended_TryMNemonic:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->TryNMemonic(pDispParams, pVarResult);
		else
			return DISP_E_MEMBERNOTFOUND;
	default:
		return DISP_E_MEMBERNOTFOUND;
	}
	return S_OK;
}

HRESULT	CMySite::CImpIExtended::GetName(
									VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
//	Utils_StringToVariant(this->m_pMySite->m_szName, pVarResult);
	InitVariantFromString(this->m_pMySite->m_szName, pVarResult);
	return S_OK;
}

HRESULT CMySite::CImpIExtended::GetObject(
									VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_UNKNOWN;
	pVarResult->punkVal	= this->m_pMySite->m_pUnk;
	if (NULL != pVarResult->punkVal)
		pVarResult->punkVal->AddRef();
	return S_OK;
}

HRESULT CMySite::CImpIExtended::DoCreateObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult)
{
	HRESULT					hr;
	UINT					uArgErr;
	VARIANTARG				varg;
	VARIANT					varResult;
	BOOL					fNewStorageFile = FALSE;	// is this a new storage file?
	IPersistStorage		*	pIPersistStorage;			// persistence on the LPS300D object
	IOleObject			*	pOleObject;
	IOleClientSite		*	pClientSite;				// our client site
	RECT					rc;
	BOOL					fSuccess		= FALSE;	// success flag
	HWND					hwndParent		= NULL;
	TCHAR					szStorageFile[MAX_PATH];
//	ITypeInfo			*	pTypeInfo;					// sink type info
//	TYPEATTR			*	pTypeAttr;					// type attribute
	IUnknown			*	punkSink;
	IProvideClassInfo2	*	pProvideClassInfo = NULL;

	// make sure that we have a return
	if (NULL == pVarResult) pVarResult = &varResult;
	VariantInit(pVarResult);
	// retrieve the parameters
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hwndParent = (HWND) varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szStorageFile, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr			= DispGetParam(pDispParams, 2, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.left		= varg.lVal;
	hr			= DispGetParam(pDispParams, 3, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.top		= varg.lVal;
	hr			= DispGetParam(pDispParams, 4, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.right	= varg.lVal;
	hr			= DispGetParam(pDispParams, 5, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	rc.bottom	= varg.lVal;
	// copy the window rectangle
	CopyRect(&(this->m_pMySite->m_rcPos), &rc);
	// open or create the storage object
	// create the storage file with null name
	hr = StgCreateDocfile(NULL,
		STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
		0, &(this->m_pMySite->m_pStorage));
	fNewStorageFile = TRUE;
	if (SUCCEEDED(hr))
	{
		// check if the punk is set
		if (NULL == this->m_pMySite->m_pUnk)
		{
			// create an instance
			hr = CoCreateInstance(this->m_pMySite->m_clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (LPVOID*)&(this->m_pMySite->m_pUnk));
		}
		else hr = S_OK;
		if (SUCCEEDED(hr))
		{
			hr = this->m_pMySite->m_pUnk->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
		}
		if (SUCCEEDED(hr))
		{
			// load the object
			hr = pOleObject->QueryInterface(IID_IPersistStorage, (LPVOID*) &pIPersistStorage);
			if (SUCCEEDED(hr))
			{
				if (fNewStorageFile)
				{
					hr = pIPersistStorage->InitNew(this->m_pMySite->m_pStorage);
				}
				else
				{
					hr = pIPersistStorage->Load(this->m_pMySite->m_pStorage);
				}
				pIPersistStorage->Release();
			}
			// show the object
			hr = this->m_pMySite->QueryInterface(IID_IOleClientSite, (LPVOID*) &pClientSite);
			if (SUCCEEDED(hr))
			{
				pOleObject->SetClientSite(pClientSite);
				// show the object
				hr = pOleObject->DoVerb(
                           OLEIVERB_SHOW,
                           NULL,
                           pClientSite,
                           -1,
                           hwndParent,
                           &rc);
				pClientSite->Release();
				fSuccess = SUCCEEDED(hr);
			}
			pOleObject->Release();
		}
		// get the sink interface id
		hr = this->m_pMySite->m_pUnk->QueryInterface(IID_IProvideClassInfo2, (LPVOID*)&pProvideClassInfo);
		if (SUCCEEDED(hr))
		{
			hr = pProvideClassInfo->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID, &(this->m_pMySite->m_iidSink));
			pProvideClassInfo->Release();
		}
	/*
		hr = Utils_GetSinkTypeInfo(this->m_pMySite->m_pUnk, &pTypeInfo);
		if (SUCCEEDED(hr))
		{
			hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
			if (SUCCEEDED(hr))
			{
				this->m_pMySite->m_iidSink = pTypeAttr->guid;
				pTypeInfo->ReleaseTypeAttr(pTypeAttr);
			}
			pTypeInfo->Release();
		}
	*/
		if (SUCCEEDED(hr))
		{
			// connect the sinks
			hr = this->m_pMySite->QueryInterface(this->m_pMySite->m_iidSink, (LPVOID*) &punkSink);
			if (SUCCEEDED(hr))
			{
		//		Utils_ConnectToConnectionPoint(punkSink, this->m_pMySite->m_iidSink, TRUE, this->m_pMySite->m_pUnk, &(this->m_pMySite->m_dwCookie), NULL);
				Utils_ConnectToConnectionPoint(this->m_pMySite->m_pUnk, punkSink, this->m_pMySite->m_iidSink, TRUE, &(this->m_pMySite->m_dwCookie));
				punkSink->Release();
			}
			hr = this->m_pMySite->QueryInterface(IID_IPropertyNotifySink, (LPVOID*) &punkSink);
			if (SUCCEEDED(hr))
			{
//				Utils_ConnectToConnectionPoint(punkSink, IID_IPropertyNotifySink, TRUE, this->m_pMySite->m_pUnk, &(this->m_pMySite->m_dwPropNotify), NULL);
				Utils_ConnectToConnectionPoint(this->m_pMySite->m_pUnk, punkSink, IID_IPropertyNotifySink, TRUE, &(this->m_pMySite->m_dwPropNotify));
				punkSink->Release();
			}
		}
	}
	pVarResult->vt		= VT_BOOL;
	pVarResult->boolVal	= fSuccess ? VARIANT_TRUE : VARIANT_FALSE;
	return hr;
}

void CMySite::CImpIExtended::DoCloseObject()
{
	HRESULT					hr;
	IOleObject			*	pOleObject;

	if (NULL == this->m_pMySite->m_pUnk) return;
	hr = this->m_pMySite->m_pUnk->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
	if (SUCCEEDED(hr))
	{
		pOleObject->Close(OLECLOSE_SAVEIFDIRTY);
		pOleObject->SetClientSite(NULL);
		pOleObject->Release();
	}
	// disconnect the sinks
	if (NULL != this->m_pMySite->m_pUnk)
	{
//		Utils_ConnectToConnectionPoint(NULL, this->m_pMySite->m_iidSink, FALSE,  this->m_pMySite->m_pUnk, &( this->m_pMySite->m_dwCookie), NULL);
//		Utils_ConnectToConnectionPoint(NULL, IID_IPropertyNotifySink, FALSE,  this->m_pMySite->m_pUnk, &( this->m_pMySite->m_dwPropNotify), NULL);
		Utils_ConnectToConnectionPoint(this->m_pMySite->m_pUnk, NULL, this->m_pMySite->m_iidSink, FALSE, &(this->m_pMySite->m_dwCookie));
		Utils_ConnectToConnectionPoint(this->m_pMySite->m_pUnk, NULL, IID_IPropertyNotifySink, FALSE, &(this->m_pMySite->m_dwPropNotify));
	}
}

HRESULT CMySite::CImpIExtended::TryNMemonic(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT			hr;
	BOOL			fSuccess = FALSE;
	MSG				msg;
	UINT			cb = sizeof(MSG);

	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = VariantToBuffer(pDispParams->rgvarg[0], (void*)&msg, cb);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pMySite->TryNMemonic(&msg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

CMySite::CImpIPropertyNotifySink::CImpIPropertyNotifySink(CMySite * pMySite) :
	m_pMySite(pMySite)
{
}

CMySite::CImpIPropertyNotifySink::~CImpIPropertyNotifySink()
{
}

// IUnknown methods
STDMETHODIMP CMySite::CImpIPropertyNotifySink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_pMySite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMySite::CImpIPropertyNotifySink::AddRef()
{
	return this->m_pMySite->AddRef();
}

STDMETHODIMP_(ULONG) CMySite::CImpIPropertyNotifySink::Release()
{
	return this->m_pMySite->Release();
}

// IPropertyNotifySink methods
STDMETHODIMP CMySite::CImpIPropertyNotifySink::OnChanged( 
									DISPID			dispID)
{
	HRESULT					hr;
	IDispatch			*	pdisp;
	VARIANTARG				avarg[2];
	if (NULL != this->m_pMySite->m_pUnkContainer)
	{
		hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
		if (SUCCEEDED(hr))
		{
			InitVariantFromString(this->m_pMySite->m_szName, &avarg[1]);
			InitVariantFromInt32(dispID, &avarg[0]);


//			Utils_StringToVariant(this->m_pMySite->m_szName, &avarg[1]);
	//		Utils_LongToVariant(dispID, &avarg[0]);
			Utils_InvokeMethod(pdisp, DISPID_Container_OnPropChanged, avarg, 2, NULL);
			VariantClear(&avarg[1]);
			pdisp->Release();
		}
	}
	return S_OK;
}

STDMETHODIMP CMySite::CImpIPropertyNotifySink::OnRequestEdit( 
									DISPID			dispID)
{
	HRESULT					hr;
	IDispatch			*	pdisp;
	VARIANTARG				avarg[2];
	VARIANT					varResult;
	BOOL					fAllow		= TRUE;			// default is allow the change
//	BOOL					fOK;
	if (NULL != this->m_pMySite->m_pUnkContainer)
	{
		hr = this->m_pMySite->m_pUnkContainer->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
		if (SUCCEEDED(hr))
		{
			InitVariantFromString(this->m_pMySite->m_szName, &avarg[1]);
			InitVariantFromInt32(dispID, &avarg[0]);

	//		Utils_StringToVariant(this->m_pMySite->m_szName, &avarg[1]);
		//	Utils_LongToVariant(dispID, &avarg[0]);
			hr = Utils_InvokeMethod(pdisp, DISPID_Container_OnPropRequestEdit, avarg, 2, &varResult);
			VariantClear(&avarg[1]);
			if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fAllow);
	//			fAllow = Utils_VariantToBool(&varResult, &fOK);
			pdisp->Release();
		}
	}
	return fAllow ? S_OK : S_FALSE;
}