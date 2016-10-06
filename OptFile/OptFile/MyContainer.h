#pragma once
#include "unknwn.h"
class CMyContainer : public IUnknown
{
public:
	CMyContainer(void);
	~CMyContainer(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// Initialization
	HRESULT						Init();
protected:
	virtual HWND				GetMainWindow()		= 0;				// pure virtual
	virtual void				OnObjectCreated(
									LPCTSTR			szObject,
									IUnknown	*	pUnkSite);
	virtual BOOL				OnPropRequestEdit(
									LPCTSTR			szObject,
									DISPID			dispid);
	virtual void				OnPropChanged(
									LPCTSTR			szObject,
									DISPID			dispid);
	virtual void				OnSinkEvent(
									LPCTSTR			ObjectName,
									long			Dispid,
									long			nParams,
									VARIANT		*	Parameters);
	virtual HRESULT				MySetActiveObject(
									IOleInPlaceActiveObject *pActiveObject);	
	// get a named object
	BOOL						GetNamedObject(
									LPCTSTR			szObject,
									REFIID			riid,
									LPVOID		*	ppv);
	// get extended object
	HRESULT						GetExtendedObject(
									LPCTSTR			szObject,
									IDispatch	**	ppdispExtended);
	// get active x object
	HRESULT						GetActiveX(
									LPCTSTR			szObject,
									IDispatch	**	ppdispActiveX);
	// close all objects
	void						CloseAllObjects();
private:
	class CLink;
	BOOL						GetLinkElement(
									LPCTSTR			szObject,
									CLink		**	ppLink);
	class CImpIOleContainer : public IOleContainer
	{
	public:
		CImpIOleContainer(CMyContainer * pBackObj);
		~CImpIOleContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IParseDisplayName Method 
		STDMETHODIMP			ParseDisplayName(
									IBindCtx	*	pbc,     
									LPOLESTR		pszDisplayName, 
									ULONG		*	pchEaten,							
									IMoniker	**	ppmkOut);
		// IOleContainer Methods
		STDMETHODIMP			EnumObjects(
									DWORD			grfFlags,  //Value specifying what is to be enumerated
									IEnumUnknown **	ppenum);
		STDMETHODIMP			LockContainer(
									BOOL			fLock);
	private:
		CMyContainer		*	m_pBackObj;
	};
	class CImpIOleInPlaceFrame : public IOleInPlaceFrame
	{
	public:
		CImpIOleInPlaceFrame(CMyContainer * pBackObj);
		~CImpIOleInPlaceFrame();
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
		// IOleInPlaceUIWindow Methods
		STDMETHODIMP			GetBorder(
									LPRECT			lprectBorder);
		STDMETHODIMP			RequestBorderSpace(
									LPCBORDERWIDTHS pborderwidths);
		STDMETHODIMP			SetBorderSpace(
									LPCBORDERWIDTHS pborderwidths);
		STDMETHODIMP			SetActiveObject(
									IOleInPlaceActiveObject *pActiveObject,
									LPCOLESTR		pszObjName);
		// IOleInPlaceFrame Methods
		STDMETHODIMP			InsertMenus(
									HMENU			hmenuShared,                 
									LPOLEMENUGROUPWIDTHS lpMenuWidths);
		STDMETHODIMP			SetMenu(
									HMENU			hmenuShared,     
									HOLEMENU		holemenu,     
									HWND			hwndActiveObject);
		STDMETHODIMP			RemoveMenus(
									HMENU			hmenuShared);
		STDMETHODIMP			SetStatusText(
									LPCOLESTR		pszStatusText);
		STDMETHODIMP			EnableModeless(
									BOOL			fEnable);
		STDMETHODIMP			TranslateAccelerator(
									LPMSG			lpmsg,  //Pointer to structure
									WORD			wID);
	private:
		CMyContainer		*	m_pBackObj;
	};
	class CImpIMyContainer : public IDispatch
	{
	public:
		CImpIMyContainer(CMyContainer * pBackObj);
		~CImpIMyContainer();
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
		HRESULT					CreateObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
		HRESULT					OnPropRequestEdit(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
		HRESULT					OnPropChanged(
									DISPPARAMS		*	pDispParams);
		HRESULT					OnSinkEvent(
									DISPPARAMS		*	pDispParams);
		HRESULT					CloseObject(
									DISPPARAMS		*	pDispParams);
		HRESULT					GetNamedObject(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
		HRESULT					TryMNemonic(
									DISPPARAMS		*	pDispParams,
									VARIANT			*	pVarResult);
	private:
		CMyContainer		*	m_pBackObj;
	};
	friend CImpIOleContainer;
	friend CImpIOleInPlaceFrame;
	friend CImpIMyContainer;

	// linked list element
	class CLink
	{
	public:
		CLink(IUnknown * punkSite);
		~CLink();
		HRESULT					GetMySite(
									IUnknown	**	ppUnkSite);
		void					GetObjectName(
									LPTSTR		*	szName);
		CLink				*	GetPrev();
		void					SetPrev(
									CLink		*	pPrev);
		CLink				*	GetNext();
		void					SetNext(
									CLink		*	pNext);
	protected:
		void					CloseObject();
	private:
		IUnknown			*	m_pUnkSite;
		CLink				*	m_pPrev;
		CLink				*	m_pNext;
	};
	CImpIOleContainer		*	m_pImpIOleContainer;
	CImpIOleInPlaceFrame	*	m_pImpIOleInPlaceFrame;		// frame manager
	CImpIMyContainer		*	m_pImpIMyContainer;
	ULONG						m_cRefs;					// main object reference count
	// linked list of contained objects
	CLink					*	m_pStart;
	CLink					*	m_pEnd;
	ULONG						m_cObjects;
};

