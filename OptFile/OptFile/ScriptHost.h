#pragma once

#include <ActivScp.h>

//class CMyObject;
//class CMyLoadWSC;
class CScriptSite;

class CScriptHost
{
public:
	CScriptHost();
	~CScriptHost();
	// get main window
	HWND			GetMainWindow();
	// script error
	void			OnScriptError(
						LPCTSTR			szError);
	// script loaded flag
	BOOL			GetScriptLoaded();
	HRESULT			LoadScript(
						LPCTSTR			szScriptText,
						IDispatch	*	pdispHost,
						LPCTSTR			szHost,
						LPCTSTR			szScriptLet);
	BOOL			LoadScript(
						LPCTSTR			szFilePath);
	void			UnloadScript();
	// get script dispatch
	BOOL			GetScriptDispatch(
						IDispatch	**	ppdispScript);
	// item information
	BOOL			GetItemUnknown(
						LPCOLESTR		pItem,
						IUnknown	**	ppUnk);
	BOOL			GetItemTypeInfo(
						LPCOLESTR		pItem,
						ITypeInfo	**	ppTypeInfo);
protected:
	// get the script engine
	HRESULT			GetScriptEngine(
						LPCTSTR		szLanguage,
						CLSID	*	pClsid);
	BOOL			LoadScriptFile(
						LPCTSTR		szFilePath,
						LPOLESTR*	pszScript);
private:
//	CMyLoadWSC	*	m_pObject;
	IActiveScript*	m_pActiveScript;
	TCHAR			m_szHost[MAX_PATH];		// host name
	IDispatch	*	m_pdispHost;			// host object
};

