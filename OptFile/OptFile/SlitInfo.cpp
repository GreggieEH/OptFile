#include "stdafx.h"
#include "SlitInfo.h"

CSlitInfo::CSlitInfo(void)
{
}

CSlitInfo::~CSlitInfo(void)
{
}

BOOL CSlitInfo::AmInit()
{
	IDispatch	*	pdisp;
	BOOL			amInit = FALSE;
	if (this->getMyObject(&pdisp))
	{
		amInit = TRUE;
		pdisp->Release();
	}
	return amInit;
}


BOOL CSlitInfo::InitNew()
{
	return this->doInit(TEXT("Sciencetech.SlitInfo.wsc.1.00"));
}

BOOL CSlitInfo::loadFromXML(
					IDispatch	*	xml)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;
	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= xml;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("loadFromXML"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CSlitInfo::saveToXML(
					IDispatch	*	xml)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;
	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= xml;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("saveToXML"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CSlitInfo::GetamReadOnly()
{
	return this->getBoolProperty(TEXT("amReadOnly"));
}

void CSlitInfo::SetamReadOnly(
				BOOL			amReadOnly)
{
	this->setBoolProperty(TEXT("amReadOnly"), FALSE);
}

BOOL CSlitInfo::GetamDirty()
{
	return this->getBoolProperty(TEXT("amDirty"));
}

void CSlitInfo::clearDirty()
{
	this->InvokeMethod(TEXT("clearDirty"), NULL, 0, NULL);
}

void CSlitInfo::addSlit(
					LPCTSTR			SlitTitle, 
					LPCTSTR			Name, 
					double			width)
{
	HRESULT				hr;
	VARIANTARG			avarg[3];
	InitVariantFromString(SlitTitle, &avarg[2]);
	InitVariantFromString(Name, &avarg[1]);
	InitVariantFromDouble(width, &avarg[0]);
	hr = this->InvokeMethod(TEXT("addSlit"), avarg, 3, NULL);
	VariantClear(&avarg[2]);
	VariantClear(&avarg[1]);
}

BOOL CSlitInfo::getSlitName(
					LPCTSTR			SlitTitle,
					LPTSTR			szName,
					UINT			nBufferSize)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	LPTSTR			szString		= NULL;
	BOOL			fSuccess		= FALSE;
	InitVariantFromString(SlitTitle, &varg);
	hr = this->InvokeMethod(TEXT("getSlitName"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szName, nBufferSize);
		fSuccess = TRUE;
/*
		if (NULL != szString)
		{
			fSuccess	= TRUE;
			StringCchCopy(szName, nBufferSize, szString);
			CoTaskMemFree((LPVOID) szString);
		}
*/
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CSlitInfo::setSlitName(
					LPCTSTR			SlitTitle,
					LPCTSTR			szName)
{
	VARIANTARG				avarg[2];
	InitVariantFromString(SlitTitle, &avarg[1]);
	InitVariantFromString(szName, &avarg[0]);
	this->InvokeMethod(TEXT("setSlitName"), avarg, 2, NULL);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
}

double CSlitInfo::getSlitWidth(
					LPCTSTR			SlitTitle)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	double				slitWidth	= -1.0;
	InitVariantFromString(SlitTitle, &varg);
	hr = this->InvokeMethod(TEXT("getSlitWidth"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &slitWidth);
	return slitWidth;
}

void CSlitInfo::setSlitWidth(
					LPCTSTR			SlitTitle,
					double			slitWidth)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(SlitTitle, &avarg[1]);
	InitVariantFromDouble(slitWidth, &avarg[0]);
	this->InvokeMethod(TEXT("setSlitWidth"), avarg, 2, NULL);
	VariantClear(&avarg[1]);
}

BOOL CSlitInfo::haveSlit(
					LPCTSTR			SlitTitle)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fHaveSlit		= FALSE;
	InitVariantFromString(SlitTitle, &varg);
	hr = this->InvokeMethod(TEXT("haveSlit"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fHaveSlit);
	return fHaveSlit;
}

// output slit width
BOOL CSlitInfo::GetOutputSlitWidth(
					double		*	slitWidth)
{
	BOOL			fSuccess	= FALSE;
	*slitWidth = 1.0;			// default slit width in mm
	if (this->haveSlit(TEXT("outputSlit")))
	{
		*slitWidth	= this->getSlitWidth(TEXT("outputSlit"));
		fSuccess	= TRUE;
	}
	return fSuccess;
}
