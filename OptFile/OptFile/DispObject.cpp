#include "stdafx.h"
#include "DispObject.h"

CDispObject::CDispObject(void) :
	m_pdisp(NULL)
{
}

CDispObject::~CDispObject(void)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

BOOL CDispObject::doInit(
						LPCTSTR			szProgID)
{
	HRESULT			hr;
//	LPOLESTR		ProgID		= NULL;
	CLSID			clsid;
	BOOL			fSuccess	= FALSE;
	IDispatch	*	pdisp = NULL;

//	Utils_AnsiToUnicode(szProgID, &ProgID);
	hr = CLSIDFromProgID(szProgID, &clsid);
//	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch,
			(LPVOID*) &pdisp);
	}
	if (SUCCEEDED(hr))
	{
		fSuccess = this->doInit(pdisp);
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CDispObject::doInit(
						IDispatch	*	pdisp)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
	if (NULL != pdisp)
	{
		this->m_pdisp	= pdisp;
		this->m_pdisp->AddRef();
		return TRUE;
	}
	return FALSE;
}

BOOL CDispObject::getMyObject(
						IDispatch	**	ppdisp)
{
	if (NULL != this->m_pdisp)
	{
		*ppdisp = this->m_pdisp;
		this->m_pdisp->AddRef();
		return TRUE;
	}
	else
	{
		*ppdisp	= NULL;
		return FALSE;
	}
}

// properties
HRESULT CDispObject::getProperty(
						LPCTSTR			szProp,
						VARIANT		*	Value)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, szProp, &dispid);
	return Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, Value);
}

HRESULT CDispObject::setProperty(
						LPCTSTR			szProp,
						VARIANTARG	*	Value)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, szProp, &dispid);
	return Utils_InvokePropertyPut(this->m_pdisp, dispid, Value, 1);
}

BOOL CDispObject::getBoolProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			bValue		= FALSE;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &bValue);
	return bValue;
}

void CDispObject::setBoolProperty(
						LPCTSTR			szProp,
						BOOL			bValue)
{
	VARIANTARG			varg;
	InitVariantFromBoolean(bValue, &varg);
	this->setProperty(szProp, &varg);
}

long CDispObject::getLongProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	long			lval		= 0;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult,&lval);
	return lval;
}

void CDispObject::setLongProperty(
						LPCTSTR			szProp,
						long			lval)
{
	VARIANTARG		varg;
	InitVariantFromInt32(lval, &varg);
	this->setProperty(szProp, &varg);
}

double CDispObject::getDoubleProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	double			dval		= 0.0;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dval);
	return dval;
}

void CDispObject::setDoubleProperty(
						LPCTSTR			szProp,
						double			dval)
{
	VARIANTARG			varg;
	InitVariantFromDouble(dval, &varg);
	this->setProperty(szProp, &varg);
}

BOOL CDispObject::getStringProperty(
						LPCTSTR			szProp,
						LPTSTR			szString,
						UINT			nBufferSize)
{
	BOOL		fSuccess	= FALSE;
	LPTSTR		szTemp		= NULL;

	szString[0]	= '\0';
	if (this->getStringProperty(szProp, &szTemp))
	{
		StringCchCopy(szString, nBufferSize, szTemp);
		fSuccess = TRUE;
		CoTaskMemFree((LPVOID) szTemp);
	}
	return fSuccess;
}

BOOL CDispObject::getStringProperty(
						LPCTSTR			szProp,
						LPTSTR		*	szString)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*szString		= NULL;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, szString);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CDispObject::setStringProperty(
						LPCTSTR			szProp,
						LPCTSTR			szString)
{
	VARIANTARG			varg;
	InitVariantFromString(szString, &varg);
	this->setProperty(szProp, &varg);
	VariantClear(&varg);
}

HRESULT	CDispObject::InvokeMethod(
						LPCTSTR			szMethod,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult)
{
	DISPID		dispid;
	Utils_GetMemid(this->m_pdisp, szMethod, &dispid);
	return Utils_InvokeMethod(this->m_pdisp, dispid, pvarg, cArgs, pVarResult);
}

DISPID CDispObject::GetDispid(
						LPCTSTR			szMethod)
{
	DISPID		dispid;
	Utils_GetMemid(this->m_pdisp, szMethod, &dispid);
	return dispid;
}

HRESULT CDispObject::InvokeMethod(
						DISPID			dispid,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult)
{
	return Utils_InvokeMethod(this->m_pdisp, dispid, pvarg, cArgs, pVarResult);
}

HRESULT CDispObject::DoInvoke(
	DISPID			dispid,
	WORD			wFlags,
	VARIANTARG	*	pvarg,
	int				cArgs,
	VARIANT		*	pVarResult)
{
	return Utils_DoInvoke(this->m_pdisp, dispid, wFlags, pvarg, cArgs, pVarResult);
}
