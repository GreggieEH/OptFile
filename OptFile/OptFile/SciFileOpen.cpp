#include "stdafx.h"
#include "SciFileOpen.h"

CSciFileOpen::CSciFileOpen(void) :
	CDispObject()
{
	this->doInit(TEXT("Sciencetech.SciFileOpen.1"));
}

CSciFileOpen::~CSciFileOpen(void)
{
}

/*
void CSciFileOpen::AddFileExtension(
						LPCTSTR		fileExtension,
						LPCTSTR		fileDescription)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(fileExtension, &avarg[1]);
	InitVariantFromString(fileDescription, &avarg[0]);
	this->InvokeMethod(TEXT("AddFileExtension"), avarg, 2, NULL);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
}
*/

/*
BOOL CSciFileOpen::SelectDataFile(
						HWND		hwnd,
						LPTSTR		szDataFile,
						UINT		nBufferSize)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	LPTSTR			szString		= NULL;
	BOOL			fSuccess		= FALSE;
	szDataFile[0]	= '\0';
	InitVariantFromInt32((long) hwnd, &varg);
	hr = this->InvokeMethod(TEXT("SelectDataFile"), &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szDataFile, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}
*/

void CSciFileOpen::SetBaseWorkingDirectory(
	LPCTSTR		szBaseWorkingDirectory)
{
	this->setStringProperty(L"BaseWorkingDirectory", szBaseWorkingDirectory);
}


BOOL CSciFileOpen::SelectCalibrationFile(
	HWND		hwndParent,
	LPTSTR		szCalibration,
	UINT		nBufferSize)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	szCalibration[0] = '\0';
	InitVariantFromInt32((long)hwndParent, &varg);
//	InitVariantFromString(folder, &avarg[0]);
	hr = this->InvokeMethod(L"SelectCalibrationFile", &varg, 1, &varResult);
//	VariantClear(&avarg[0]);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szCalibration, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

