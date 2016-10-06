#include "stdafx.h"
#include "SciOutFile.h"


CSciOutFile::CSciOutFile()
{
	this->doInit(L"Sciencetech.SciOutFile.1");
}


CSciOutFile::~CSciOutFile()
{
}

void CSciOutFile::SetfileName(
	LPCTSTR			szfileName)
{
	this->setStringProperty(L"fileName", szfileName);
}
void CSciOutFile::loadFile()
{
	this->InvokeMethod(L"loadFile", NULL, 0, NULL);
}
BOOL CSciOutFile::GetxValues(
	long		*	nX,
	double		**	ppX)
{
	HRESULT			hr;
	VARIANT			varResult;
	long			ub;
	long			i;
	double		*	arr;
	BOOL			fSuccess = FALSE;
	*nX = 0;
	*ppX = NULL;
	hr = this->getProperty(L"xValues", &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nX = ub + 1;
			*ppX = new double[*nX];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nX); i++)
				(*ppX)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}
BOOL CSciOutFile::GetyValues(
	long		*	nX,
	double		**	ppY)
{
	HRESULT			hr;
	VARIANT			varResult;
	long			ub;
	long			i;
	double		*	arr;
	BOOL			fSuccess = FALSE;
	*nX = 0;
	*ppY = NULL;
	hr = this->getProperty(L"yValues", &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt)
		{
			fSuccess = TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nX = ub + 1;
			*ppY = new double[*nX];
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nX); i++)
				(*ppY)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}
