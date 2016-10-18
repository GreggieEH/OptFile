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
//	if (this->doInit(TEXT("Sciencetech.InputInfo.wsc.1.00")))
//	{
	this->InvokeMethod(TEXT("initNew"), NULL, 0, NULL);
	return TRUE;
//	}
//	return FALSE;
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
//	return this->getBoolProperty(TEXT("amReadOnly"));
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fReadonly = TRUE;
	hr = this->InvokeMethod(L"get_amReadOnly", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fReadonly);
	return fReadonly;
}

void CInputInfo::SetamReadOnly(
						BOOL			amReadOnly)
{
//	this->setBoolProperty(TEXT("amReadOnly"), FALSE);
//	VARIANTARG			varg;
//	InitVariantFromBoolean(FALSE, &varg);
//	this->InvokeMethod(L"")
}

double CInputInfo::GetlampDistance()
{
	HRESULT			hr;
	VARIANT			varResult;
	double			lampDistance = 0.0;
	hr = this->InvokeMethod(L"get_lampDistance", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &lampDistance);
	return lampDistance;
//	return this->getDoubleProperty(TEXT("lampDistance"));
}

void CInputInfo::SetlampDistance(
						double			lampDistance)
{
//	this->setDoubleProperty(TEXT("lampDistance"), lampDistance);
	VARIANTARG			varg;
	InitVariantFromDouble(lampDistance, &varg);
	this->InvokeMethod(L"put_lampDistance", &varg, 1, NULL);
}

BOOL CInputInfo::GetinputOpticConfig(
						LPTSTR			szinputOpticConfig,
						UINT			nBufferSize)
{
//	return this->getStringProperty(TEXT("inputOpticConfig"), szinputOpticConfig, nBufferSize);
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	szinputOpticConfig[0] = '\0';
	hr = this->InvokeMethod(L"get_inputOpticConfig", NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szinputOpticConfig, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CInputInfo::SetinputOpticConfig(
						LPCTSTR			szinputOpticConfig)
{
//	this->setStringProperty(TEXT("inputOpticConfig"), szinputOpticConfig);
	VARIANTARG			varg;
	InitVariantFromString(szinputOpticConfig, &varg);
	this->InvokeMethod(L"put_inputOpticConfig", &varg, 1, NULL);
	VariantClear(&varg);
}

BOOL CInputInfo::GetamDirty()
{
//	return this->getBoolProperty(TEXT("amDirty"));
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			amDirty = TRUE;
	hr = this->InvokeMethod(L"get_amDirty", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &amDirty);
	return amDirty;
}

void CInputInfo::clearDirty()
{
	this->InvokeMethod(TEXT("clearDirty"), NULL, 0, NULL);
}

long CInputInfo::GetnumInputConfigs()
{
//	return this->getLongProperty(TEXT("numInputConfigs"));
	HRESULT			hr;
	VARIANT			varResult;
	long			nConfig = 0;
	hr = this->InvokeMethod(L"get_numInputConfigs", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &nConfig);
	return nConfig;
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
//	this->setDoubleProperty(L"userSetinput", userSetinput);
	VARIANTARG			varg;
	InitVariantFromDouble(userSetinput, &varg);
	this->InvokeMethod(L"put_userSetinput", &varg, 1, NULL);

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
//	hr = this->getProperty(L"userSetinput", &varResult);
	hr = this->InvokeMethod(L"get_userSetinput", NULL, 0, &varResult);
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
