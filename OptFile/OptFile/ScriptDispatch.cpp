#include "stdafx.h"
#include "ScriptDispatch.h"
#include "ScriptHost.h"
#include "Server.h"

CScriptDispatch::CScriptDispatch(LPCTSTR szScriptFile) :
	m_pScriptHost(NULL)
{
	TCHAR			szFilePath[MAX_PATH];
	IDispatch	*	pdisp;

	GetModuleFileName(GetServer()->GetInstance(), szFilePath, MAX_PATH);
	PathRemoveFileSpec(szFilePath);
	PathAppend(szFilePath, szScriptFile);

	this->m_pScriptHost = new CScriptHost;
	this->m_pScriptHost->LoadScript(szFilePath);
}

CScriptDispatch::~CScriptDispatch()
{
	Utils_DELETE_POINTER(this->m_pScriptHost);
}


HRESULT CScriptDispatch::InvokeMethod(
	LPCTSTR			szMethod,
	VARIANTARG	*	pvarg,
	int				cArgs,
	VARIANT		*	pVarResult)
{
	HRESULT				hr = E_FAIL;
	DISPID				dispid;
	IDispatch		*	pdisp;

	if (this->GetScript(&pdisp))
	{
		Utils_GetMemid(pdisp, szMethod, &dispid);
		hr = Utils_InvokeMethod(pdisp, dispid, pvarg, cArgs, pVarResult);
		pdisp->Release();
	}
	return hr;
}

BOOL CScriptDispatch::GetScript(
	IDispatch	**	ppdisp)
{
	if (NULL != this->m_pScriptHost)
	{
		return this->m_pScriptHost->GetScriptDispatch(ppdisp);
	}
	else
	{
		*ppdisp = NULL;
		return FALSE;
	}
}
