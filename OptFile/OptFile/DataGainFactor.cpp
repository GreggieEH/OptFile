#include "stdafx.h"
#include "DataGainFactor.h"

DataGainFactor::DataGainFactor() :
	CScriptDispatch(L"DataGainFactor.jvs.txt")
{
/*
	TCHAR			szFilePath[MAX_PATH];
//	CScriptHost		scriptHost;
	IDispatch	*	pdisp;

	GetModuleFileName(GetServer()->GetInstance(), szFilePath, MAX_PATH);
	PathRemoveFileSpec(szFilePath);
	PathAppend(szFilePath, L"DataGainFactor.jvs.txt");

//	this->doInit(L"Sci.DataGainFactor.wsc.1.00");
	this->m_pScriptHost = new CScriptHost;
	this->m_pScriptHost->LoadScript(szFilePath);
*/
}

DataGainFactor::~DataGainFactor()
{
//	Utils_DELETE_POINTER(this->m_pScriptHost);
}

void DataGainFactor::SetacquireMode(
	long		acquireMode)
{
//	this->setLongProperty(L"acquireMode", acquireMode);
//
//	long		aMode = this->getLongProperty(L"acquireMode");
//	BOOL		stat = FALSE;
	HRESULT			hr;
	VARIANTARG		varg;
	InitVariantFromInt32(acquireMode, &varg);
	hr = this->InvokeMethod(L"put_acquireMode", &varg, 1, NULL);
	BOOL			stat = FALSE;
}

void DataGainFactor::SetlockinGain(
	long		lockinGain)
{
//	this->setLongProperty(L"lockinGain", lockinGain);
	HRESULT			hr;
	VARIANTARG		varg;
	InitVariantFromInt32(lockinGain, &varg);
	hr = this->InvokeMethod(L"put_lockinGain", &varg, 1, NULL);
	BOOL			stat = FALSE;
}

void DataGainFactor::Setdetector(
	long		detector)
{
//	this->setLongProperty(L"detector", detector);
	HRESULT			hr;
	VARIANTARG		varg;
	InitVariantFromInt32(detector, &varg);
	hr = this->InvokeMethod(L"put_detector", &varg, 1, NULL);
	BOOL			stat = FALSE;
}

void DataGainFactor::SetdetectorGain(
	LPCTSTR		szdetectorGain)
{
//	this->setStringProperty(L"detectorGain", szdetectorGain);
	VARIANTARG			varg;
	HRESULT				hr;
	InitVariantFromString(szdetectorGain, &varg);
	hr = this->InvokeMethod(L"put_detectorGain", &varg, 1, NULL);
	VariantClear(&varg);
}

double DataGainFactor::GetgainFactor()
{
//	return this->getDoubleProperty(L"gainFactor");
	HRESULT			hr;
	VARIANT			varResult;
	double			gainFactor = 1.0;
	hr = this->InvokeMethod(L"get_gainFactor", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &gainFactor);
	return gainFactor;
}

double DataGainFactor::GetminRawDataValue()
{
//	return this->getDoubleProperty(L"minRawDataValue");
	HRESULT			hr;
	VARIANT			varResult;
	double			minValue = 0.0;
	hr = this->InvokeMethod(L"get_minRawDataValue", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &minValue);
	return minValue;
}

double DataGainFactor::GetDCOffset()
{
//	return this->getDoubleProperty(L"DCOffset");
	HRESULT			hr;
	VARIANT			varResult;
	double			dcOffset = 0.0;
	hr = this->InvokeMethod(L"get_DCOffset", NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dcOffset);
	return dcOffset;
}

/*
HRESULT DataGainFactor::InvokeMethod(
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

BOOL DataGainFactor::GetScript(
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
*/
