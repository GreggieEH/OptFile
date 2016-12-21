#include "stdafx.h"
#include "GratingScanInfo.h"
#include "DynamicArray.h"
#include "ImpGratingScanInfo.h"
#include "InterpolateValues.h"

CGratingScanInfo::CGratingScanInfo(void) :
	CDispObject()
{
}

CGratingScanInfo::~CGratingScanInfo(void)
{
}

BOOL CGratingScanInfo::InitNew()
{
//	return this->doInit(TEXT("Sciencetech.GratingScanInfo.wsc.1.00"));
	HRESULT			hr;
	CImpGratingScanInfo *pGratingScan;
	IDispatch	*	pdisp = NULL;
	BOOL			fSuccess = FALSE;

	pGratingScan = new CImpGratingScanInfo();
	hr = pGratingScan->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		fSuccess = this->doInit(pdisp);
		pdisp->Release();
	}
	else delete pGratingScan;
	return fSuccess;
}

BOOL CGratingScanInfo::loadFromXML(
					IDispatch	*	xml)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;

/*
	// set logging on
	this->SetLogging();
	// want to load spectral data
	this->SetLoadSpectralData();
*/
	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= xml;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("loadFromXML"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	// finished with logging
//	this->ClearLogging();
//	// finished load spectral data
//	this->ClearLoadSpectralData();
	return fSuccess;
}

BOOL CGratingScanInfo::saveToXML(
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

BOOL CGratingScanInfo::GetamReadOnly()
{
	return FALSE;
//	return this->getBoolProperty(TEXT("amReadOnly"));
}

void CGratingScanInfo::SetamReadOnly(
					BOOL			amReadOnly)
{
//	this->setBoolProperty(TEXT("amReadOnly"), FALSE);
}

long CGratingScanInfo::Getgrating()
{
	return this->getLongProperty(TEXT("grating"));
}

void CGratingScanInfo::Setgrating(
					long			grating)
{
	this->setLongProperty(TEXT("grating"), grating);
}

BOOL CGratingScanInfo::Getfilter(
					LPTSTR			szFilter,
					UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("Filter"), szFilter, nBufferSize);
}

void CGratingScanInfo::Setfilter(
					LPCTSTR			szFilter)
{
	this->setStringProperty(TEXT("Filter"), szFilter);
}

long CGratingScanInfo::GetDetector()
{
	return this->getLongProperty(L"Detector");
}
void CGratingScanInfo::SetDetector(
	long			detector)
{
	this->setLongProperty(L"Detector", detector);
}

double CGratingScanInfo::Getdispersion()
{
	return this->getDoubleProperty(TEXT("dispersion"));
}

void CGratingScanInfo::Setdispersion(
					double			dispersion)
{
	this->setDoubleProperty(TEXT("dispersion"), dispersion);
}

BOOL CGratingScanInfo::Getwavelengths(
					long		*	nValues,
					double		**	ppWaves)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	long			ub;
	long			i;
	double		*	arr;
	*nValues		= 0;
	*ppWaves		= NULL;
	hr = this->getProperty(TEXT("wavelengths"), &varResult);
	if (SUCCEEDED(hr))
	{
//		fSuccess = TranslateJScriptArray(&varResult, nValues, ppWaves);
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nValues = ub + 1;
			*ppWaves = new double[*nValues];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nValues); i++)
				(*ppWaves)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CGratingScanInfo::Getsignals(
					long		*	nValues,
					double		**	ppSignals)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	long			ub;
	long			i;
	double		*	arr;
	*nValues		= 0;
	*ppSignals		= NULL;
	hr = this->getProperty(TEXT("signals"), &varResult);
	if (SUCCEEDED(hr))
	{
//		fSuccess = TranslateJScriptArray(&varResult, nValues, ppSignals);
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nValues = ub + 1;
			*ppSignals = new double[*nValues];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nValues); i++)
				(*ppSignals)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CGratingScanInfo::addData(
					double			wavelength,
					double			signal)
{
	VARIANTARG			avarg[2];
	InitVariantFromDouble(wavelength, &avarg[1]);
	InitVariantFromDouble(signal, &avarg[0]);
	this->InvokeMethod(TEXT("addData"), avarg, 2, NULL);
//	long			index = this->findIndex(wavelength, 0.01);
//	BOOL			stat = FALSE;
}

BOOL CGratingScanInfo::GetnodeName(
					LPTSTR			sznodeName,
					UINT			nBufferSize)
{
	return this->getStringProperty(TEXT("NodeName"), sznodeName, nBufferSize);
}

void CGratingScanInfo::SetnodeName(
					LPCTSTR			sznodeName)
{
	this->setStringProperty(TEXT("NodeName"), sznodeName);
}

BOOL CGratingScanInfo::GetamDirty()
{
	return this->getBoolProperty(TEXT("amDirty"));
}

void CGratingScanInfo::clearDirty()
{
	this->InvokeMethod(TEXT("clearDirty"), NULL, 0, NULL);
}

BOOL CGratingScanInfo::GetGratingScan(
					long		*	nValues,
					double		**	ppWaves,
					double		**	ppSignals)
{
	long			nX	= 0;
	long			nY	= 0;
	BOOL			fSuccess	= FALSE;
	*nValues	= 0;
	*ppWaves	= NULL;
	*ppSignals	= NULL;

	if (this->Getwavelengths(&nX, ppWaves))
	{
		if (this->Getsignals(&nY, ppSignals))
		{
			if (nX == nY)
			{
				fSuccess = TRUE;
				*nValues = nX;
			}
			else
			{
				delete [] *ppSignals;
				*ppSignals = NULL;
			}
		}
		if (!fSuccess)
		{
			delete [] *ppWaves;
			*ppWaves = NULL;
		}
	}
	return fSuccess;
}


long CGratingScanInfo::GetarraySize()
{
	return this->getLongProperty(TEXT("arraySize"));
}

BOOL CGratingScanInfo::CheckGratingAndFilter(
					long			grating, 
					LPCTSTR			szFilter,
					long			detector)
{
	if (this->Getgrating() != grating) return FALSE;
	TCHAR			szString[MAX_PATH];
	if (detector >= 0 && this->GetDetector() >= 0)
	{
		if (this->GetDetector() != detector) return FALSE;
	}
	if (this->Getfilter(szString, MAX_PATH))
		return 0 == lstrcmpi(szFilter, szString);
	return FALSE;
}

void CGratingScanInfo::SetLogging()
{
/*
	VARIANTARG			varg;
	CImpIDispatch	*	pDispatch;

	pDispatch	= new CImpIDispatch(this, L"onLogString");
	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	pDispatch->QueryInterface(IID_IDispatch, (LPVOID*) &(varg.pdispVal));
	this->setProperty(TEXT("onLogString"), &varg);
	VariantClear(&varg);
*/
}

void CGratingScanInfo::ClearLogging()
{
//	VARIANTARG			varg;
//	VariantInit(&varg);
//	this->setProperty(TEXT("onLogString"), &varg);
}

void CGratingScanInfo::OnLogString()
{
/*
	LPTSTR			szString	= NULL;
	this->getStringProperty(TEXT("logString"), &szString);
	if (NULL != szString)
	{
		CoTaskMemFree((LPVOID) szString);
	}
*/
}

void CGratingScanInfo::SetLoadSpectralData()
{
/*
	VARIANTARG			varg;
	CImpIDispatch	*	pDispatch;

	pDispatch = new CImpIDispatch(this, L"onHaveSpectralData");
	VariantInit(&varg);
	varg.vt = VT_DISPATCH;
	pDispatch->QueryInterface(IID_IDispatch, (LPVOID*)&(varg.pdispVal));
	this->setProperty(TEXT("onHaveSpectralData"), &varg);
	VariantClear(&varg);
*/
}

void CGratingScanInfo::ClearLoadSpectralData()
{
//	VARIANTARG			varg;
//	VariantInit(&varg);
//	this->setProperty(TEXT("onHaveSpectralData"), &varg);
}

void CGratingScanInfo::OnHaveSpectralData()
{
/*
	HRESULT				hr;
	VARIANT				varResult;
	LPTSTR				szSpectralData = NULL;
	hr = this->getProperty(L"SpectralData", &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, &szSpectralData);
		if (SUCCEEDED(hr))
		{
			if (this->LoadSpectralData(szSpectralData))
			{
				this->setBoolProperty(L"translatedSpectralData", TRUE);
			}
			CoTaskMemFree((LPVOID)szSpectralData);
		}
		VariantClear(&varResult);
	}
*/
}

//BOOL CGratingScanInfo::LoadSpectralData(
//	LPTSTR			szString)
//{
/*
	BOOL			fSuccess = FALSE;
	size_t			slen1, slen2;
	LPTSTR			szRem;			// remainder string
	LPTSTR			szTemp;
	TCHAR			szOneLine[MAX_PATH];
	float			f1, f2;
	CDynamicArray	dynArray;				// dynamic array
	long			nWaves;
	double		*	pWaves;
	long			nSignals;
	double		*	pSignals;

	szTemp = szString;
	while (NULL != szTemp)
	{
		szRem = StrStr(szTemp, L"\n");
		if (NULL != szRem)
		{
			StringCchLength(szTemp, STRSAFE_MAX_CCH, &slen1);
			StringCchLength(szRem, STRSAFE_MAX_CCH, &slen2);
			StringCchCopyN(szOneLine, MAX_PATH, szTemp, slen1 - slen2);
			if (2 == _stscanf_s(szOneLine, L"%f\t%f", &f1, &f2))
			{
				dynArray.AddValues(f1, f2);
			}
		}
		else
		{
			StringCchCopy(szOneLine, MAX_PATH, szTemp);
			if (2 == _stscanf_s(szOneLine, L"%f\t%f", &f1, &f2))
			{
				dynArray.AddValues(f1, f2);
			}
			slen2 = 0;
		}
		if (slen2 > 2)
		{
			szTemp = &(szRem[1]);
		}
		else
		{
			szTemp = NULL;
		}
	}
	if (dynArray.GetXArray(&nWaves, &pWaves))
	{
		if (dynArray.GetYArray(&nSignals, &pSignals))
		{
			if (nWaves == nSignals)
			{
				fSuccess = this->SetDataArray(nWaves, pWaves, pSignals);
			}
			delete[] pSignals;
		}
		delete[] pWaves;
	}

	return fSuccess;
*/
//}


BOOL CGratingScanInfo::SetDataArray(
					long			nValues,
					double		*	pWaves,
					double		*	pSignals)
{
	// clear the current spectral data
//	this->InvokeMethod(TEXT("clearSpectralData"), NULL, 0, NULL);
	this->InvokeMethod(L"ClearData", NULL, 0, NULL);
	long			i;
	VARIANTARG		avarg[2];
	for (i = 0; i < nValues; i++)
	{
		InitVariantFromDouble(pWaves[i], &avarg[1]);
		InitVariantFromDouble(pSignals[i], &avarg[0]);
		this->InvokeMethod(L"addData", avarg, 2, NULL);
	}
	return TRUE;
}

/*
long CGratingScanInfo::findClosestIndex(
					double			wavelength)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			index	= 10000;
	InitVariantFromDouble(wavelength, &varg);
	hr = this->InvokeMethod(TEXT("findClosestIndex"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &index);
	return index;
}
*/

double CGratingScanInfo::getValue(
					long			index)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	double			signal	= -1.0;
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(TEXT("getValue"), &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &signal);
	return signal;
}


void CGratingScanInfo::setValue(
					long			index,
					double			signal)
{
	VARIANTARG		avarg[2];
	InitVariantFromInt32(index, &avarg[1]);
	InitVariantFromDouble(signal, &avarg[0]);
	this->InvokeMethod(TEXT("setValue"), avarg, 2, NULL);
}

BOOL CGratingScanInfo::GetWavelengthRange(
	double		*	minWave,
	double		*	maxWave)
{
	double		*	pWaves;
	long			nWaves;
	BOOL			fSuccess = FALSE;
	*minWave = 0.0;
	*maxWave = 0.0;
	if (this->Getwavelengths(&nWaves, &pWaves))
	{
		fSuccess = TRUE;
		*minWave = pWaves[0];
		*maxWave = pWaves[nWaves - 1];
	}
	return fSuccess;
}

long CGratingScanInfo::findIndex(
	double			wavelength,
	double			tolerance)
{
	HRESULT				hr;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	long				index = -1;
	InitVariantFromDouble(wavelength, &avarg[1]);
	InitVariantFromDouble(tolerance, &avarg[0]);
	hr = this->InvokeMethod(L"findIndex", avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &index);
	return index;
}

BOOL CGratingScanInfo::GethaveExtraValue()
{
	return this->getBoolProperty(L"haveExtraValue");
}

void CGratingScanInfo::SethaveExtraValue(
	BOOL			fhaveExtraValue)
{
//	this->setBoolProperty(L"haveExtraValue", fhaveExtraValue);
}

void CGratingScanInfo::setExtraValue(
		LPCTSTR			szTitle,
		double			wavelength,
		double			extraValue)
{
	VARIANTARG			avarg[3];
	InitVariantFromString(szTitle, &avarg[2]);
	InitVariantFromDouble(wavelength, &avarg[1]);
	InitVariantFromDouble(extraValue, &avarg[0]);
	this->InvokeMethod(L"setExtraValue", avarg, 3, NULL);
	VariantClear(&avarg[2]);
}

BOOL CGratingScanInfo::GetextraValues(
	LPCTSTR			szTitle,
	long		*	nValues,
	double		**	ppextraValues)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	long			ub;
	long			i;
	double		*	arr;
	*nValues = 0;
	*ppextraValues = NULL;

	InitVariantFromString(szTitle, &varg);
//	hr = this->getProperty(TEXT("extraValues"), &varResult);
	hr = this->DoInvoke(this->GetDispid(L"ExtraValueComponent"), DISPATCH_PROPERTYGET, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nValues = ub + 1;
			*ppextraValues = new double[*nValues];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nValues); i++)
				(*ppextraValues)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CGratingScanInfo::GetextraValueWavelengths(
	LPCTSTR			szTitle,
	long		*	nValues,
	double		**	ppWaves)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	long			ub;
	long			i;
	double		*	arr;
	*nValues = 0;
	*ppWaves = NULL;

	InitVariantFromString(szTitle, &varg);
	//	hr = this->getProperty(TEXT("extraValues"), &varResult);
	hr = this->DoInvoke(this->GetDispid(L"ExtraValueWavelengths"), DISPATCH_PROPERTYGET, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nValues = ub + 1;
			*ppWaves = new double[*nValues];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nValues); i++)
				(*ppWaves)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

long CGratingScanInfo::GetNumberExtraValues(
	LPCTSTR			szTitle)
{
	long		nValues = 0;
	double	*	pWaves = NULL;
	if (this->GetextraValueWavelengths(szTitle, &nValues, &pWaves))
	{
		delete[] pWaves;
	}
	return nValues;
}


/*
long CGratingScanInfo::findIndex(
	double			wavelength,
	double			tolerance)
{
	HRESULT				hr;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	long				index = -1;
	InitVariantFromDouble(wavelength, &avarg[1]);
	InitVariantFromDouble(tolerance, &avarg[0]);
	hr = this->InvokeMethod(L"findIndex", avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &index);
	return index;
}
*/

/*
CGratingScanInfo::CImpIDispatch::CImpIDispatch(CGratingScanInfo* pGratingScanInfo, LPCTSTR szName) :
	m_pGratingScanInfo(pGratingScanInfo),
	m_cRefs(0)
{
	StringCchCopy(this->m_szName, MAX_PATH, szName);
}

CGratingScanInfo::CImpIDispatch::~CImpIDispatch()
{
}

// IUnknown methods
STDMETHODIMP CGratingScanInfo::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CGratingScanInfo::CImpIDispatch::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CGratingScanInfo::CImpIDispatch::Release()
{
	ULONG			cRefs	= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CGratingScanInfo::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CGratingScanInfo::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CGratingScanInfo::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CGratingScanInfo::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	if (DISPID_VALUE == dispIdMember)
	{
		if (0 == lstrcmpi(this->m_szName, L"onLogString"))
			this->m_pGratingScanInfo->OnLogString();
		else if (0 == lstrcmpi(this->m_szName, L"onHaveSpectralData"))
			this->m_pGratingScanInfo->OnHaveSpectralData();
	}
	return S_OK;
}
*/

BOOL CGratingScanInfo::GetExtraValueTitles(
	LPTSTR		**	ppszTitles,
	ULONG		*	nStrings)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	*ppszTitles = NULL;
	*nStrings = 0;
	hr = this->getProperty(L"ExtraValueTitles", &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringArrayAlloc(varResult, ppszTitles, nStrings);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}
// signal units
BOOL CGratingScanInfo::GetSignalUnits(
	LPTSTR			szUnits,
	UINT			nBufferSize)
{
	return this->getStringProperty(L"SignalUnits", szUnits, nBufferSize);
}
void CGratingScanInfo::SetSignalUnits(
	LPCTSTR			szUnits)
{
	this->setStringProperty(L"SignalUnits", szUnits);
}

// interpolated data to 0.1 nm
BOOL CGratingScanInfo::GetNMInterpolatedData(
	long		*	nValues,
	double		**	ppWaves,
	double		**	ppSignals)
{
	CInterpolateValues	intValues;
	BOOL				fSuccess = FALSE;
	long				nX;
	double			*	pX;
	long				nY;
	double			*	pY;
	long				i;
//	double				wave;
	double				startWave;
	double				endWave;

	*nValues = 0;
	*ppWaves = NULL;
	*ppSignals = NULL;
	// read our data
	if (this->Getwavelengths(&nX, &pX))
	{
		if (this->Getsignals(&nY, &pY))
		{
			if (nX == nY)
			{
				// wavelength range
				startWave = pX[0];
				endWave = pX[nX - 1];
				*nValues = (((long)floor(endWave + 0.5) - (long)floor(startWave + 0.5)) * 10) + 1;
				*ppWaves = new double[*nValues];
				*ppSignals = new double[*nValues];
				// interpolation
				intValues.SetXArray(nX, pX);
				intValues.SetYArray(nX, pY);
				for (i = 0; i<(*nValues); i++)
				{
					//					(*ppWaves)[i]	= floor(startWave) + (i * 1.0);
					(*ppWaves)[i] = floor(startWave + 0.5) + (i * 0.1);
					(*ppSignals)[i] = intValues.interpolateValue((*ppWaves)[i]);
				}
				fSuccess = TRUE;
			}
			delete[] pY;
		}
		delete[] pX;
	}
	return fSuccess;
}
