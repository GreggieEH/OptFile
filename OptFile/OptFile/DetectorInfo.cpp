#include "stdafx.h"
#include "DetectorInfo.h"

CDetectorInfo::CDetectorInfo(void) :
	CDispObject()
{
}

CDetectorInfo::~CDetectorInfo(void)
{
}

BOOL CDetectorInfo::InitNew()
{
	return this->doInit(TEXT("Sciencetech.DetectorInfo.wsc.1.00"));
}

BOOL CDetectorInfo::loadFromXML(
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

BOOL CDetectorInfo::saveToXML(
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

BOOL CDetectorInfo::GetamReadOnly()
{
	return this->getBoolProperty(TEXT("amReadOnly"));
}

void CDetectorInfo::SetamReadOnly(
						BOOL			famReadOnly)
{
	this->setBoolProperty(TEXT("amReadOnly"), FALSE);
}

BOOL CDetectorInfo::GetdetectorModel(
						LPTSTR			szDetector,
						UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("detectorModel"), szDetector, nBufferSize);
}

void CDetectorInfo::SetdetectorModel(
						LPCTSTR			szDetector)
{
	this->setStringProperty(TEXT("detectorModel"), szDetector);
}

BOOL CDetectorInfo::GetportName(
						LPTSTR			szPortName,
						UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("portName"), szPortName, nBufferSize);
}

void CDetectorInfo::SetportName(
						LPCTSTR			szPortName)
{
	this->setStringProperty(TEXT("portName"), szPortName);
}

BOOL CDetectorInfo::GetgainSetting(
						LPTSTR			szgainSetting,
						UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("gainSetting"), szgainSetting, nBufferSize);
}

void CDetectorInfo::SetgainSetting(
						LPCTSTR			szgainSetting)
{
	this->setStringProperty(TEXT("gainSetting"), szgainSetting);
}

double CDetectorInfo::Gettemperature()
{
	return this->getDoubleProperty(TEXT("temperature"));
}

void CDetectorInfo::Settemperature(
						double			temperature)
{
	this->setDoubleProperty(TEXT("temperature"), temperature);
}

BOOL CDetectorInfo::GetamDirty()
{
	return this->getBoolProperty(TEXT("amDirty"));
}

void CDetectorInfo::clearDirty()
{
	this->InvokeMethod(TEXT("clearDirty"), NULL, 0, NULL);
}

long CDetectorInfo::GetnumADChannels()
{
	return this->getLongProperty(TEXT("numADChannels"));
}

long CDetectorInfo::getADChannel(
						long			index)
						
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			adChannel	= -1;
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(TEXT("getADChannel"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &adChannel);
	return adChannel;
}

double CDetectorInfo::getLowLimit(
						long			index)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	double			lowLimit		= -1.0;
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(TEXT("getLowLimit"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &lowLimit);
	return lowLimit;
}

double CDetectorInfo::getHiLimit(
						long			index)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	double			hiLimit			= -1.0;
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(TEXT("getHiLimit"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &hiLimit);
	return hiLimit;
}

long CDetectorInfo::findADChannel(
						double			wavelength)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	long				channel	= -1;
	InitVariantFromDouble(wavelength, &varg);
	hr = this->InvokeMethod(TEXT("findADChannel"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &channel);
	return channel;
}

void CDetectorInfo::addADChannel(
						long			ADChannel, 
						double			lowLimit, 
						double			hiLimit)
{
	VARIANTARG			avarg[3];
	InitVariantFromInt32(ADChannel, &avarg[2]);
	InitVariantFromDouble(lowLimit, &avarg[1]);
	InitVariantFromDouble(hiLimit, &avarg[0]);
	this->InvokeMethod(TEXT("addADChannel"), avarg, 3, NULL);
}

void CDetectorInfo::clearADChannels()
{
	this->InvokeMethod(TEXT("clearADChannels"), NULL, 0, NULL);
}

long CDetectorInfo::FindIndexFromChannel(
						long			channel)
{
	long		index		= -1;
	long		i			= 0;
	BOOL		fDone		= FALSE;
	long		nChannels	= this->GetnumADChannels();
	while (i < nChannels && !fDone)
	{
		if (this->getADChannel(i) == channel)
		{
			fDone = TRUE;
			index = i;
		}
		if (!fDone) i++;
	}
	return index;
}

BOOL CDetectorInfo::GetADInfoString(
	LPTSTR		*	szInfo)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	*szInfo = NULL;
	hr = this->getProperty(L"ADInfoString", &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, szInfo);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CDetectorInfo::SetADInfoString(
	LPCTSTR			szInfo)
{
	VARIANTARG			varg;
	InitVariantFromString(szInfo, &varg);
	this->setProperty(L"ADInfoString", &varg);
	VariantClear(&varg);
}

BOOL CDetectorInfo::GetPulseIntegrator()
{
	LPTSTR			szADInfo = NULL;
	BOOL			pulseIntegrator = FALSE;
	if (this->GetADInfoString(&szADInfo))
	{
		pulseIntegrator = NULL != StrStrI(szADInfo, L"Pulse Integrator");
		CoTaskMemFree((LPVOID)szADInfo);
	}
	return pulseIntegrator;
}
