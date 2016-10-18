#pragma once

#include <ActivScp.h>

class CScriptHost;

class CScriptSite : public IUnknown
{
public:
	CScriptSite(CScriptHost * pScriptHost);
	~CScriptSite();
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
private:
	// nested classes
	class CImpIActiveScriptSite : public IActiveScriptSite
	{
	public:
		CImpIActiveScriptSite(CScriptSite * pScriptSite);
		~CImpIActiveScriptSite();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IActiveScriptSite methods
		STDMETHODIMP			GetLCID(
									LCID		*	plcid);
		STDMETHODIMP			GetItemInfo(
									LPCOLESTR		pstrName,
									DWORD			dwReturnMask,
									IUnknown	**	ppunkItem,
									ITypeInfo	**	ppTypeInfo);
		STDMETHODIMP			GetDocVersionString(
									BSTR		*	pbstrVersionString);
		STDMETHODIMP			OnScriptTerminate(
									const VARIANT*	pvarResult,
									const EXCEPINFO*	pexcepinfo);
		STDMETHODIMP			OnStateChange(
									SCRIPTSTATE		ssScriptState);
		STDMETHODIMP			OnScriptError(
									IActiveScriptError *pase);
		STDMETHODIMP			OnEnterScript(void);
		STDMETHODIMP			OnLeaveScript(void);
	private:
		CScriptSite			*	m_pScriptSite;
	};
	class CImpIActiveScriptSiteWindow : public IActiveScriptSiteWindow
	{
	public:
		CImpIActiveScriptSiteWindow(CScriptSite * pScriptSite);
		~CImpIActiveScriptSiteWindow();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IActiveScriptSiteWindow methods
		STDMETHODIMP			GetWindow(
									HWND		*	phwnd);
		STDMETHODIMP			EnableModeless(
									BOOL			fEnable);
	private:
		CScriptSite			*	m_pScriptSite;
	};
	friend CImpIActiveScriptSite;
	friend CImpIActiveScriptSiteWindow;
	CScriptHost				*	m_pScriptHost;				// script host
	// data members
	CImpIActiveScriptSite	*	m_pImpIActiveScriptSite;
	CImpIActiveScriptSiteWindow*	m_pImpIActiveScriptSiteWindow;
	// reference count
	ULONG						m_cRefs;
};

