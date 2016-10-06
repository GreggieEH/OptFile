#include "stdafx.h"
#include "InputInfo.h"

CInputInfo::CInputInfo(void) :
	CDispObject()
{
}

CInputInfo::~CInputInfo(void)
{
}

BOOL CInputInfo::InitNew()
{
	if (this->doInit(TEXT("Sciencetech.InputInfo.wsc.1.00")))
	{
		this->InvokeMethod(TEXT("initNew"), NULL, 0, NULL);
		return TRUE;
	}
	return FALSE;
}

BOOL CInputInfo::loadFromXML(
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
	if (fSuccess)
	{
		double		setInput = -1.0;
		if (this->GetuserSetinput(&setInput))
		{
			BOOL		stat = TRUE;
		}
	}
	return fSuccess;
}

BOOL CInputInfo::saveToXML(
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

BOOL CInputInfo::GetamReadOnly()
{
	return this->getBoolProperty(TEXT("amReadOnly"));
}

void CInputInfo::SetamReadOnly(
						BOOL			amReadOnly)
{
	this->setBoolProperty(TEXT("amReadOnly"), FALSE);
}

double CInputInfo::GetlampDistance()
{
	return this->getDoubleProperty(TEXT("lampDistance"));
}

void CInputInfo::SetlampDistance(
						double			lampDistance)
{
	this->setDoubleProperty(TEXT("lampDistance"), lampDistance);
}

BOOL CInputInfo::GetinputOpticConfig(
						LPTSTR			szinputOpticConfig,
						UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("inputOpticConfig"), szinputOpticConfig, nBufferSize);
}

void CInputInfo::SetinputOpticConfig(
						LPCTSTR			szinputOpticConfig)
{
	this->setStringProperty(TEXT("inputOpticConfig"), szinputOpticConfig);
}

BOOL CInputInfo::GetamDirty()
{
	return this->getBoolProperty(TEXT("amDirty"));
}

void CInputInfo::clearDirty()
{
	this->InvokeMethod(TEXT("clearDirty"), NULL, 0, NULL);
}

long CInputInfo::GetnumInputConfigs()
{
	return this->getLongProperty(TEXT("numInputConfigs"));
}

BOOL CInputInfo::getInputConfig(
						long			index,
						LPTSTR			szInputConfig,
						UINT			nBufferSize)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	LPTSTR				szString		= NULL;
	BOOL				fSuccess		= FALSE;
	szInputConfig[0]	= '\0';
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(TEXT("getInputConfig"), &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szInputConfig, nBufferSize);
		fSuccess = TRUE;
	/*
		if (NULL != szString)
		{
			fSuccess	= TRUE;
			StringCchCopy(szInputConfig, nBufferSize, szString);
			CoTaskMemFree((LPVOID) szString);
		}
	*/
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CInputInfo::getRadianceAvailable(double * FOV)
{
	TCHAR			szInputOptic[MAX_PATH];
	float			f1;
	BOOL			fAvail		= FALSE;

	szInputOptic[0]	= '\0';
	*FOV	= 0.0017;
	this->GetinputOpticConfig(szInputOptic, MAX_PATH);
	if (NULL != StrStrI(szInputOptic, TEXT("FOV")))
	{
		fAvail	= TRUE;
		if (1 == _stscanf_s(szInputOptic, TEXT("FOV %f"), &f1))
		{
			*FOV	= f1;
		}
	}
	return fAvail;
}

//double CInputInfo::GetuserSetinput()
//{
//	return this->getDoubleProperty(L"userSetinput");
//}

void CInputInfo::SetuserSetinput(
	double			userSetinput)
{
	this->setDoubleProperty(L"userSetinput", userSetinput);

	TCHAR			szString[MAX_PATH];
	this->GetinputOpticConfig(szString, MAX_PATH);
	BOOL		stat = TRUE;
}

BOOL CInputInfo::GetuserSetinput(
	double		*	userSetInput)
{
	HRESULT				hr;
	VARIANT				varResult;
	BOOL				fUserInput = FALSE;
	TCHAR				szString[MAX_PATH];
	*userSetInput = 0.0;
	hr = this->getProperty(L"userSetinput", &varResult);
	if (SUCCEEDED(hr))
	{
		if (VT_BSTR == varResult.vt)
		{
			hr = VariantToString(varResult, szString, MAX_PATH);
			VariantClear(&varResult);
		}
		else
		{
			hr = VariantToDouble(varResult, userSetInput);
			fUserInput = SUCCEEDED(hr);
		}
	}
	return fUserInput;
}

void CInputInfo::clearUserSetInput()
{
	this->InvokeMethod(L"clearUserSetInput", NULL, 0, NULL);
}
