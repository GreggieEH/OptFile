#pragma once
#include "unknwn.h"
class CMySite : public IUnknown
{
public:
	CMySite(LPCTSTR szProgID, IUnknown * pUnkContainer, LPCTSTR szName);
	~CMySite(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
protected:
	void						GetControlInfo();
	BOOL						TryNMemonic(
									LPMSG			pmsg);
private:
	class CImpIAmbients : public IDispatch
	{
	public:
		CImpIAmbients(CMySite * pMySite);
		~CImpIAmbients();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount(  
									PUINT				pctinfo);
		STDMETHODIMP			GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo);
		STDMETHODIMP			GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId);
		STDMETHODIMP			Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr);
	private:
		CMySite				*	m_pMySite;
	};
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CMySite * pMySite);
		~CImpISink();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount(  
									PUINT				pctinfo);
		STDMETHODIMP			GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo);
		STDMETHODIMP			GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId);
		STDMETHODIMP			Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr);
	private:
		CMySite				*	m_pMySite;
	};
	class CImpIExtended : public IDispatch
	{
	public:
		CImpIExtended(CMySite * pMySite);
		~CImpIExtended();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID				riid,
									LPVOID			*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount(  
									PUINT				pctinfo);
		STDMETHODIMP			GetTypeInfo(  
									UINT				iTInfo,         
									LCID				lcid,                   
									ITypeInfo		**  ppTInfo);
		STDMETHODIMP			GetIDsOfNames(  
									REFIID				riid,                  
									OLECHAR			**  rgszNames,  
									UINT				cNames,          
									LCID				lcid,                   
									DISPID			*	rgDispId);
		STDMETHODIMP			Invoke(  
									DISPID				dispIdMember,      
									REFIID				riid,              
									LCID				lcid,                
									WORD				wFlags,              
									DISPPARAMS		*	pDispParams,  
									VARIANT			*	pVarResult,  
									EXCEPINFO		*	pExcepInfo,  
									PUINT				puArgErr);
	protected:
		HRESULT					GetName(
									VARIANT			*	pVarResult);
		HRESULT					GetObject(
									VARIANT			*	pVarResult);
		HRESULT					DoCreateObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
		void					DoCloseObject();
		HRESULT					TryNMemonic(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
	private:
		CMySite				*	m_pMySite;
	};
	class CImpIPropertyNotifySink : public IPropertyNotifySink
	{
	public:
		CImpIPropertyNotifySink(CMySite * pMySite);
		~CImpIPropertyNotifySink();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotifySink methods
		STDMETHODIMP			OnChanged( 
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit( 
									DISPID			dispID);
	private:
		CMySite				*	m_pMySite;
	};
	class CImpIOleClientSite : public IOleClientSite		// client site
	{
	public:
		CImpIOleClientSite(CMySite * pMySite);
		~CImpIOleClientSite();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleClientSite methods
		STDMETHODIMP			SaveObject();
		STDMETHODIMP			GetMoniker(
									DWORD			dwAssign,  
									DWORD			dwWhichMoniker,
									IMoniker	**	ppmk);
		STDMETHODIMP			GetContainer(
									LPOLECONTAINER* ppContainer);
		STDMETHODIMP			ShowObject();
		STDMETHODIMP			OnShowWindow(
									BOOL			fShow);
		STDMETHODIMP			RequestNewObjectLayout();
	private:
		CMySite				*	m_pMySite;
	};
	class CImpIOleInPlaceSite : public IOleInPlaceSite		// in place site
	{
	public:
		CImpIOleInPlaceSite(CMySite * pMySite);
		~CImpIOleInPlaceSite();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleWindow Methods
		STDMETHODIMP			GetWindow(
									HWND		*	phwnd);
		STDMETHODIMP			ContextSensitiveHelp(
									BOOL			fEnterMode);
		// IOleInPlaceSite Methods
		STDMETHODIMP			CanInPlaceActivate();
		STDMETHODIMP			OnInPlaceActivate();
		STDMETHODIMP			OnUIActivate();
		STDMETHODIMP			GetWindowContext(
									IOleInPlaceFrame **ppFrame,
									IOleInPlaceUIWindow **ppDoc,
									LPRECT			lprcPosRect,  //Points to position of in-place object
									LPRECT			lprcClipRect, //Points to in-place object's position 
																	 // rectangle
									LPOLEINPLACEFRAMEINFO lpFrameInfo);
		STDMETHODIMP			Scroll(
									SIZE			scrollExtent);
		STDMETHODIMP			OnUIDeactivate(
									BOOL			fUndoable);
		STDMETHODIMP			OnInPlaceDeactivate();
		STDMETHODIMP			DiscardUndoState();
		STDMETHODIMP			DeactivateAndUndo();
		STDMETHODIMP			OnPosRectChange(
									LPCRECT			lprcPosRect); 
	private:
		CMySite				*	m_pMySite;
	};
	class CImpIOleControlSite : public IOleControlSite		// control site
	{
	public:
		CImpIOleControlSite(CMySite * pMySite);
		~CImpIOleControlSite();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleControlSite Methods
		STDMETHODIMP			OnControlInfoChanged(void);
		STDMETHODIMP			LockInPlaceActive(
									BOOL			fLock);
		STDMETHODIMP			GetExtendedControl(
									IDispatch	**	ppDisp);
		STDMETHODIMP			TransformCoords(
									POINTL		*	pPtlHimetric ,  //Address of POINTL structure
									POINTF		*	pPtfContainer ,  //Address of POINTF structure
									DWORD			dwFlags);
		STDMETHODIMP			TranslateAccelerator(
									LPMSG			pMsg ,        //Pointer to the structure
									DWORD			grfModifiers);
		STDMETHODIMP			OnFocus(
									BOOL			fGotFocus);
		STDMETHODIMP			ShowPropertyFrame(void);
	private:
		CMySite				*	m_pMySite;
	};
	friend CImpIAmbients;
	friend CImpISink;
	friend CImpIExtended;
	friend CImpIPropertyNotifySink;
	friend CImpIOleClientSite;
	friend CImpIOleInPlaceSite;
	friend CImpIOleControlSite;
	// data members
	CImpIAmbients			*	m_pImpIAmbients;
	CImpISink				*	m_pImpISink;
	CImpIExtended			*	m_pImpIExtended;
	CImpIPropertyNotifySink	*	m_pImpIPropertyNotifySink;
	CImpIOleClientSite		*	m_pImpIOleClientSite;
	CImpIOleInPlaceSite		*	m_pImpIOleInPlaceSite;
	CImpIOleControlSite		*	m_pImpIOleControlSite;
	// object reference count
	ULONG						m_cRefs;
	// active X object
	IUnknown				*	m_pUnk;
	IStorage				*	m_pStorage;
	// assigned object name
	LPTSTR						m_szName;
	// container
	IUnknown				*	m_pUnkContainer;
	// sink interface id
	IID							m_iidSink;
	DWORD						m_dwCookie;			// connection cookie
	DWORD						m_dwPropNotify;		// property notify cookie
	// window rectangle for the ActiveX control
	RECT						m_rcPos;
	// control info
	CONTROLINFO					m_controlInfo;
	// object class ID
	CLSID						m_clsid;
};

