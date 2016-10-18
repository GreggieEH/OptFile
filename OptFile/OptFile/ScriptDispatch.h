#pragma once

class CScriptHost;

class CScriptDispatch
{
public:
	CScriptDispatch(LPCTSTR szScriptFile);
	virtual ~CScriptDispatch();
	HRESULT		InvokeMethod(
					LPCTSTR			szMethod,
					VARIANTARG	*	pvarg,
					int				cArgs,
					VARIANT		*	pVarResult);
	BOOL		GetScript(
					IDispatch	**	ppdisp);
private:
	CScriptHost*	m_pScriptHost;
};

