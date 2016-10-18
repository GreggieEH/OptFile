#include "stdafx.h"
#include "ScriptSite.h"
#include "ScriptHost.h"

CScriptSite::CScriptSite(CScriptHost * pScriptHost) :
	m_pScriptHost(pScriptHost),
	m_pImpIActiveScriptSite(NULL),
	m_pImpIActiveScriptSiteWindow(NULL),
	// reference count
	m_cRefs(0)
{
}

CScriptSite::~CScriptSite()
{
	Utils_DELETE_POINTER(this->m_pImpIActiveScriptSite);
	Utils_DELETE_POINTER(this->m_pImpIActiveScriptSiteWindow);
}

// IUnknown methods
STDMETHODIMP CScriptSite::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
	{
		*ppv = this;
	}
	else if (IID_IActiveScriptSite == riid)
	{
		*ppv = this->m_pImpIActiveScriptSite;
	}
	else if (IID_IActiveScriptSiteWindow == riid)
	{
		*ppv = this->m_pImpIActiveScriptSiteWindow;
	}
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

STDMETHODIMP_(ULONG) CScriptSite::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CScriptSite::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CScriptSite::Init()
{
	this->m_pImpIActiveScriptSite = new CImpIActiveScriptSite(this);
	this->m_pImpIActiveScriptSiteWindow = new CImpIActiveScriptSiteWindow(this);
	return S_OK;
}

CScriptSite::CImpIActiveScriptSite::CImpIActiveScriptSite(CScriptSite * pScriptSite) :
	m_pScriptSite(pScriptSite)
{
}

CScriptSite::CImpIActiveScriptSite::~CImpIActiveScriptSite()
{
}

// IUnknown methods
STDMETHODIMP CScriptSite::CImpIActiveScriptSite::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_pScriptSite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CScriptSite::CImpIActiveScriptSite::AddRef()
{
	return this->m_pScriptSite->AddRef();
}

STDMETHODIMP_(ULONG) CScriptSite::CImpIActiveScriptSite::Release()
{
	return this->m_pScriptSite->Release();
}

// IActiveScriptSite methods
STDMETHODIMP CScriptSite::CImpIActiveScriptSite::GetLCID(
	LCID		*	plcid)
{
	return E_NOTIMPL;			// use default
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::GetItemInfo(
	LPCOLESTR		pstrName,
	DWORD			dwReturnMask,
	IUnknown	**	ppunkItem,
	ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	if (ppunkItem) *ppunkItem = NULL;			// don't have named item
	if (ppTypeInfo)*ppTypeInfo = NULL;
	if (0 != (SCRIPTINFO_IUNKNOWN & dwReturnMask))
	{
		hr = this->m_pScriptSite->m_pScriptHost->GetItemUnknown(pstrName, ppunkItem) ? S_OK : TYPE_E_ELEMENTNOTFOUND;
	}
	if (0 != (SCRIPTINFO_ITYPEINFO & dwReturnMask))
	{
		hr = this->m_pScriptSite->m_pScriptHost->GetItemTypeInfo(pstrName, ppTypeInfo) ? S_OK : TYPE_E_ELEMENTNOTFOUND;
	}
	return hr;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::GetDocVersionString(
	BSTR		*	pbstrVersionString)
{
	*pbstrVersionString = SysAllocString(L"SciInput");
	return S_OK;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::OnScriptTerminate(
	const VARIANT*	pvarResult,
	const EXCEPINFO*	pexcepinfo)
{
	return S_OK;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::OnStateChange(
	SCRIPTSTATE		ssScriptState)
{
	return S_OK;
}

#define				STRING_BUFFER			1024

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::OnScriptError(
	IActiveScriptError *pase)
{
	DWORD		dwCookie;
	LONG		nChar;
	ULONG		nLine;
	BSTR		bstr = 0;
	EXCEPINFO	ei;
	//	LPOLESTR	pOleStr;
	TCHAR		szMessage[STRING_BUFFER];
	LPTSTR		szString;
	TCHAR		szTemp[MAX_PATH];

	ZeroMemory(&ei, sizeof(ei));
/*
	pase->GetSourcePosition(&dwCookie, &nLine, &nChar);
	pase->GetSourceLineText(&bstr);
	pase->GetExceptionInfo(&ei);
	// form the message string
	szString = NULL;
//	SHStrDup(ei.bstrSource, &szString);
	if (WideCharToMultiByte(CP_ACP, 0, ei.bstrSource, -1,szTemp, MAX_PATH, NULL, NULL) > 0)
//	if (NULL != szString)
	{
		StringCchCopy(szMessage, STRING_BUFFER, szTemp);
	}
	StringCchCat(szMessage, STRING_BUFFER, TEXT("\n[Line :d]"));
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%d"), nLine);
	StringCchCat(szMessage, STRING_BUFFER, szTemp);
	StringCchCat(szMessage, STRING_BUFFER, TEXT("\n"));
	if (NULL != bstr)
	{
		if (WideCharToMultiByte(CP_ACP, 0, bstr, -1, szTemp, MAX_PATH, NULL, NULL) > 0)
		{
			StringCchCat(szMessage, STRING_BUFFER, szTemp);
		}
		SysFreeString(bstr);
		bstr = NULL;
	}
	// error description
	if (NULL != ei.bstrDescription)
	{
		szString = NULL;
		if (WideCharToMultiByte(CP_ACP, 0, ei.bstrDescription, -1, szTemp, MAX_PATH, NULL, NULL) > 0)
		{
			StringCchCat(szMessage, STRING_BUFFER, TEXT("\nDescription ="));
			StringCchCat(szMessage, STRING_BUFFER, szTemp);
		}
	}
	// error notification
	this->m_pScriptSite->m_pScriptHost->OnScriptError(szMessage);
*/
	SysFreeString(ei.bstrSource);
	SysFreeString(ei.bstrDescription);
	SysFreeString(ei.bstrHelpFile);
	return S_OK;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::OnEnterScript(void)
{
	return S_OK;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSite::OnLeaveScript(void)
{
	return S_OK;
}

CScriptSite::CImpIActiveScriptSiteWindow::CImpIActiveScriptSiteWindow(CScriptSite * pScriptSite)
{

}

CScriptSite::CImpIActiveScriptSiteWindow::~CImpIActiveScriptSiteWindow()
{

}

// IUnknown methods
STDMETHODIMP CScriptSite::CImpIActiveScriptSiteWindow::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_pScriptSite->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CScriptSite::CImpIActiveScriptSiteWindow::AddRef()
{
	return this->m_pScriptSite->AddRef();
}

STDMETHODIMP_(ULONG) CScriptSite::CImpIActiveScriptSiteWindow::Release()
{
	return this->m_pScriptSite->Release();
}

// IActiveScriptSiteWindow methods
STDMETHODIMP CScriptSite::CImpIActiveScriptSiteWindow::GetWindow(
	HWND		*	phwnd)
{
	*phwnd = this->m_pScriptSite->m_pScriptHost->GetMainWindow();
	return S_OK;
}

STDMETHODIMP CScriptSite::CImpIActiveScriptSiteWindow::EnableModeless(
	BOOL			fEnable)
{
	return S_OK;
}
