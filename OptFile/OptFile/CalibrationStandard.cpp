#include "stdafx.h"
#include "CalibrationStandard.h"
#include "ImpICalibrationStandard.h"

CCalibrationStandard::CCalibrationStandard(void) :
	CDispObject()
{
}

CCalibrationStandard::~CCalibrationStandard(void)
{
}


BOOL CCalibrationStandard::InitNew()
{
	HRESULT						hr;
	CImpICalibrationStandard*	pCalStan;
	IDispatch				*	pdisp;
	BOOL						fSuccess = FALSE;

	pCalStan = new CImpICalibrationStandard();
	hr = pCalStan->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		fSuccess = this->doInit(pdisp);
		pdisp->Release();
	}
	else delete pCalStan;
	return fSuccess;
}

BOOL CCalibrationStandard::loadFromXML(
						IDispatch	*	pdispxml)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;

	TCHAR			szFileName[MAX_PATH];

	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= pdispxml;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("loadFromXML"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		this->GetfileName(szFileName, MAX_PATH);
		if (PathFileExists(szFileName))
		{
			fSuccess = TRUE;
		}

//		VariantToBoolean(varResult, &fSuccess);
	//	this->GetfileName(szFileName, MAX_PATH);
	}
	return fSuccess;
}

BOOL CCalibrationStandard::saveToXML(
						IDispatch	*	pdispxml)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;
	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= pdispxml;
	varg.pdispVal->AddRef();
	hr = this->InvokeMethod(TEXT("saveToXML"), &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CCalibrationStandard::GetfileName(
						LPTSTR			szFileName,
						UINT			nBufferSize)
{
	return this->getStringProperty(L"fileName", szFileName, nBufferSize);
}

void CCalibrationStandard::SetfileName(
						LPCTSTR			szfileName)
{
	this->setStringProperty(L"fileName", szfileName);
}

BOOL CCalibrationStandard::GetamLoaded()
{
	return this->getBoolProperty(L"amLoaded");
}

void CCalibrationStandard::SetamLoaded(
						BOOL			amLoaded)
{
	this->setBoolProperty(L"amLoaded", amLoaded);
}

BOOL CCalibrationStandard::Getwavelengths(
						long		*	nValues,
						double		**	ppWaves)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*nValues		= 0;
	*ppWaves		= NULL;
	hr = this->getProperty(TEXT("wavelengths"), &varResult);
	if (SUCCEEDED(hr))
	{
//		fSuccess = TranslateJScriptArray(&varResult, nValues, ppWaves);
		fSuccess = this->VariantToDoubleArray(&varResult, nValues, ppWaves);
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CCalibrationStandard::Getcalibration(
						long		*	nValues,
						double		**	ppcalibration)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*nValues		= 0;
	*ppcalibration	= NULL;
	hr = this->getProperty(TEXT("calibration"), &varResult);
	if (SUCCEEDED(hr))
	{
//		fSuccess = TranslateJScriptArray(&varResult, nValues, ppcalibration);
		fSuccess = this->VariantToDoubleArray(&varResult, nValues, ppcalibration);
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CCalibrationStandard::loadExcelFile(
						LPCTSTR			szFileName)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	InitVariantFromString(szFileName, &varg);
	hr = this->InvokeMethod(L"loadExcelFile", &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CCalibrationStandard::checkExcelFile(
						LPCTSTR			szFileName)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fExcel = FALSE;
	InitVariantFromString(szFileName, &varg);
	hr = this->InvokeMethod(L"checkExcelFile", &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fExcel);
	return fExcel;
	return 0 == lstrcmpi(PathFindExtension(szFileName), L".xls");
}

BOOL CCalibrationStandard::VariantToDoubleArray(
	VARIANT			*	Value,
	long			*	nValues,
	double			**	ppValues)
{
	long		i;
	long		ub;
	BOOL		fSuccess = FALSE;
	double	*	arr;
	*nValues = 0;
	*ppValues = NULL;
	if ((VT_ARRAY | VT_R8) != Value->vt || NULL == Value->parray) return FALSE;
	SafeArrayGetUBound(Value->parray, 1, &ub);
	*nValues = ub + 1;
	*ppValues = new double[*nValues];
	SafeArrayAccessData(Value->parray, (void**)&arr);
	for (i = 0; i < (*nValues); i++)
	{
		(*ppValues)[i] = arr[i];
	}
	SafeArrayUnaccessData(Value->parray);
	return TRUE;
}
