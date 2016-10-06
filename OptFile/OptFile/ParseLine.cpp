#include "stdafx.h"
#include "ParseLine.h"


CParseLine::CParseLine()
{
//	this->doInit(L"Sciencetech.ParseLine.wsc.1.00");
	this->doInit(L"Sciencetech.ParseLine.1");
}


CParseLine::~CParseLine()
{
}

void CParseLine::doParseLine(
	LPCTSTR			oneLine)
{
	VARIANTARG			varg;
	InitVariantFromString(oneLine, &varg);
	this->InvokeMethod(L"doParseLine", &varg, 1, NULL);
	VariantClear(&varg);
}
long CParseLine::GetnumSubStrings()
{
	return this->getLongProperty(L"numSubStrings");

}
BOOL CParseLine::getSubString(
	long			index,
	LPTSTR			szString,
	UINT			nBufferSize)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	szString[0] = '\0';
	InitVariantFromInt32(index, &varg);
	hr = this->InvokeMethod(L"getSubString", &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szString, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}
