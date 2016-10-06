#include "stdafx.h"
#include "ImpGratingScanInfo.h"
#include "dispids.h"
#include "MyGuids.h"
#include "Server.h"
#include "SciXML.h"
#include "DynamicArray.h"
#include "ParseLine.h"

CImpGratingScanInfo::CImpGratingScanInfo() :
	m_cRefs(0),					// object reference count
	//TCHAR					m_szNodeName[MAX_PATH];		// object node name
	m_grating(-1),
	//TCHAR					m_szFilter[MAX_PATH];
	m_dispersion(2.0),
	m_detector(1),
	// spectral data
	m_nWaves(0),
	m_pWavelengths(NULL),
	m_nSignals(0),
	m_pSignals(NULL),
	// extra values
	m_nExtraValues(0),
	m_ppExtraValues(NULL),
	// dirty flag
	m_amDirty(FALSE)
{
	this->m_szNodeName[0] = '\0';
	this->m_szFilter[0] = '\0';
	// default signal title
	StringCchCopy(this->m_szSignal, MAX_PATH, L"Volts");
}

CImpGratingScanInfo::~CImpGratingScanInfo()
{
	this->ClearData();
	this->ClearExtraValues();
}

// IUnknown methods
STDMETHODIMP CImpGratingScanInfo::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || IID_IGratingScanInfo == riid)
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

STDMETHODIMP_(ULONG) CImpGratingScanInfo::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CImpGratingScanInfo::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CImpGratingScanInfo::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

STDMETHODIMP CImpGratingScanInfo::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	HRESULT			hr;
	ITypeLib	*	pTypeLib	= NULL;
	*ppTInfo = NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(IID_IGratingScanInfo, ppTInfo);
		pTypeLib->Release();
	}
	return hr;
}

STDMETHODIMP CImpGratingScanInfo::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo = NULL;
	hr = this->GetTypeInfo(0, lcid, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CImpGratingScanInfo::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	switch (dispIdMember)
	{
	case DISPID_GratingScanInfo_NodeName:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNodeName(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetNodeName(pDispParams);
		break;
	case DISPID_GratingScanInfo_grating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetGrating(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetGrating(pDispParams);
		break;
	case DISPID_GratingScanInfo_Filter:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetFilter(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetFilter(pDispParams);
		break;
	case DISPID_GratingScanInfo_dispersion:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDispersion(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDispersion(pDispParams);
		break;
	case DISPID_GratingScanInfo_Wavelengths:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetWavelengths(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetWavelengths(pDispParams);
		break;
	case DISPID_GratingScanInfo_Signals:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSignals(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSignals(pDispParams);
		break;
	case DISPID_GratingScanInfo_amDirty:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAmDirty(pVarResult);
		break;
	case DISPID_GratingScanInfo_loadFromXML:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->loadFromXML(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_saveToXML:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->saveToXML(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_addData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->addData(pDispParams);
		break;
	case DISPID_GratingScanInfo_ClearData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ClearData();
		break;
	case DISPID_GratingScanInfo_clearDirty:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->clearDirty();
		break;
	case DISPID_GratingScanInfo_findIndex:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->findIndex(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_arraySize:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetarraySize(pVarResult);
		break;
	case DISPID_GratingScanInfo_getValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->getValue(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_setValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->setValue(pDispParams);
		break;
//	case DISPID_GratingScanInfo_extraValues:
//		if (0 != (wFlags & DISPATCH_PROPERTYGET))
//			return this->GetextraValues(pVarResult);
//		break;
	case DISPID_GratingScanInfo_setExtraValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->setExtraValue(pDispParams);
		break;
	case DISPID_GratingScanInfo_haveExtraValue:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GethaveExtraValue(pVarResult);
		break;
	case DISPID_GratingScanInfo_ExtraValuesTitles:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetExtraValuesTitles(pVarResult);
		break;
	case DISPID_GratingScanInfo_SignalUnits:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSignalUnits(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSignalUnits(pDispParams);
		break;
	case DISPID_GratingScanInfo_ExtraValueComponent:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetExtraValueComponent(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_ExtraValueWavelengths:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetExtraValueWavelengths(pDispParams, pVarResult);
		break;
	case DISPID_GratingScanInfo_Detector:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDetector(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDetector(pDispParams);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CImpGratingScanInfo::loadFromXML(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	BOOL			fSuccess = FALSE;
	CSciXML			xml;

	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_DISPATCH != pDispParams->rgvarg[0].vt ||
		NULL == pDispParams->rgvarg[0].pdispVal) return DISP_E_TYPEMISMATCH;
	if (xml.doInit(pDispParams->rgvarg[0].pdispVal))
	{
		fSuccess = this->loadFromXML(&xml);
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

BOOL CImpGratingScanInfo::loadFromXML(
	CSciXML		*	xml)
{
	IDispatch	*	pdispRoot = NULL;
	BOOL			fSuccess = FALSE;
	IDispatch	*	pdispNode = NULL;
	IDispatch	*	pdispChild = NULL;
	TCHAR			szNodeValue[MAX_PATH];
	long			lval;
	float			fval;
	LPTSTR			szSpectralData = NULL;
	LPTSTR			szExtraValues = NULL;

	if (xml->GetrootNode(&pdispRoot))
	{
		if (xml->findNamedChild(pdispRoot, this->m_szNodeName, &pdispNode))
		{
			if (xml->findNamedChild(pdispNode, L"grating", &pdispChild))
			{
				if (xml->getNodeValue(pdispChild, szNodeValue, MAX_PATH))
				{
					if (1 == _stscanf_s(szNodeValue, L"%d", &lval))
					{
						this->m_grating = lval;
					}
				}
				pdispChild->Release();
			}
			if (xml->findNamedChild(pdispNode, L"filter", &pdispChild))
			{
				xml->getNodeValue(pdispChild, this->m_szFilter, MAX_PATH);
				pdispChild->Release();
			}
			if (xml->findNamedChild(pdispNode, L"Dispersion", &pdispChild))
			{
				if (xml->getNodeValue(pdispChild, szNodeValue, MAX_PATH))
				{
					if (1 == _stscanf_s(szNodeValue, L"%f", &fval))
					{
						this->m_dispersion = fval;
					}
				}
				pdispChild->Release();
			}
			if (xml->findNamedChild(pdispNode, L"Detector", &pdispChild))
			{
				if (xml->getNodeValue(pdispChild, szNodeValue, MAX_PATH))
				{
					if (1 == _stscanf_s(szNodeValue, L"%d", &lval))
					{
						this->m_detector = lval;
					}
				}
				pdispChild->Release();
			}
			if (xml->findNamedChild(pdispNode, L"SpectralData", &pdispChild))
			{
				if (xml->getNodeValue(pdispChild, &szSpectralData))
				{
					this->LoadSpectralData(szSpectralData);
					CoTaskMemFree((LPVOID)szSpectralData);
				}
				pdispChild->Release();
			}
/*
			if (xml->findNamedChild(pdispNode, L"ExtraValues", &pdispChild))
			{
				if (xml->getNodeValue(pdispChild, &szExtraValues))
				{
					this->LoadExtraValues(szExtraValues);
					CoTaskMemFree((LPVOID)szExtraValues);
				}
				pdispChild->Release();
			}
*/
			pdispNode->Release();
		}
		pdispRoot->Release();
	}
	return fSuccess;
}

BOOL CImpGratingScanInfo::LoadSpectralData(
	LPTSTR			szString)
{
	BOOL			fSuccess = FALSE;
	LPTSTR			szTemp;
	TCHAR			szOneLine[MAX_PATH];
	double			x, y;
	CDynamicArray	dynArray;				// dynamic array
	BOOL			fHaveTitles = FALSE;	// have titles flag

	this->ClearData();
	this->ClearExtraValues();
	szTemp = szString;
	if (this->FormLine(&szTemp, szOneLine, MAX_PATH))
	{
		if (this->ReadOneLine(szOneLine, &x, &y))
		{
			dynArray.AddValues(x, y);
		}
		else
		{
			fHaveTitles = ReadTitles(szOneLine);
		}
	}
	while (NULL != szTemp)
	{
		if (this->FormLine(&szTemp, szOneLine, MAX_PATH))
		{
			if (this->ReadOneLine(szOneLine, &x, &y))
			{
				dynArray.AddValues(x, y);
			}
		}
	}
	if (dynArray.GetXArray(&this->m_nWaves, &this->m_pWavelengths) &&
		dynArray.GetYArray(&this->m_nSignals, &this->m_pSignals))
	{
		fSuccess = TRUE;
	}
	return fSuccess;
}

BOOL CImpGratingScanInfo::FormLine(
	LPTSTR		*	szString,
	LPTSTR			szOneLine,
	UINT			nBufferSize)
{
	UINT			c;
	UINT			i;
	BOOL			fDone;
	LPTSTR			szTemp = *szString;

	i = 0;
	c = 0;
	szOneLine[0] = '\0';
	fDone = FALSE;
	while (!fDone)
	{
		if ('\n' == szTemp[i])
		{
			if (c > 0)
			{
				fDone = TRUE;
				*szString = &(szTemp[i]);
			}
		}
		else if ('\0' == szTemp[i])
		{
			fDone = TRUE;
			*szString = NULL;
		}
		else
		{
			szOneLine[c] = szTemp[i];
			c++;
		}
		i++;
	}
	szOneLine[c] = '\0';
	return c > 0;
}

BOOL CImpGratingScanInfo::ReadTitles(
	LPCTSTR			szOneLine)
{
	CParseLine		parseLine;
	BOOL			fSuccess = FALSE;
	long			nParse;
	long			i;
	TCHAR			szSubString[MAX_PATH];
	CExtraValues*	pExtraValues;
	parseLine.doParseLine(szOneLine);
	nParse = parseLine.GetnumSubStrings();
	if (nParse >= 2)
	{
		// ignore first
		// second gives signal units
		if (parseLine.getSubString(1, this->m_szSignal, MAX_PATH))
		{
			fSuccess = TRUE;
		}
		// remainder give extra values
		for (i = 2; i < nParse; i++)
		{
			if (parseLine.getSubString(i, szSubString, MAX_PATH))
			{
				pExtraValues = new CExtraValues;
				pExtraValues->SetTitle(szSubString);
				// add to the end array
				if (NULL == this->m_ppExtraValues)
				{
					this->m_ppExtraValues = new CExtraValues *[1];
					this->m_ppExtraValues[0] = pExtraValues;
					this->m_nExtraValues = 1;
				}
				else
				{
					long			nAlloc = this->m_nExtraValues + 1;
					CExtraValues**	ppExtraValues = new CExtraValues *[nAlloc];
					long			i;
					for (i = 0; i < this->m_nExtraValues; i++)
						ppExtraValues[i] = this->m_ppExtraValues[i];
					ppExtraValues[nAlloc - 1] = pExtraValues;
					delete[] this->m_ppExtraValues;
					this->m_ppExtraValues = ppExtraValues;
					this->m_nExtraValues = nAlloc;
				}
			}
		}
	}
	return fSuccess;
}

BOOL CImpGratingScanInfo::ReadOneLine(
	LPTSTR			szOneLine,
	double		*	x,
	double		*	y)
{
//	LPTSTR			szRem;
//	LPTSTR			szTemp;
	float			f1, f2, f3, f4, f5, f6;
	BOOL			fHaveExtraValues = FALSE;
	BOOL			fSuccess = FALSE;

	// see if there are extra values on the line
	if (this->m_nExtraValues > 0)
	{
		if (1 == this->m_nExtraValues)
		{ 
			if (3 == _stscanf_s(szOneLine, L"%f\t%f\t%f", &f1, &f2, &f3))
			{
				fHaveExtraValues = TRUE;
				fSuccess = TRUE;
				*x = f1;
				*y = f2;
				this->m_ppExtraValues[0]->AddValue(f1, f3);
			}
		}
		else if (2 == this->m_nExtraValues)
		{
			if (4 == _stscanf_s(szOneLine, L"%f\t%f\t%f\t%f", &f1, &f2, &f3, &f4))
			{
				fHaveExtraValues = TRUE;
				fSuccess = TRUE;
				*x = f1;
				*y = f2;
				this->m_ppExtraValues[0]->AddValue(f1, f3);
				this->m_ppExtraValues[1]->AddValue(f1, f4);
			}
		}
		else if (3 == this->m_nExtraValues)
		{
			if (5 == _stscanf_s(szOneLine, L"%f\t%f\t%f\t%f\t%f", &f1, &f2, &f3, &f4, &f5))
			{
				fHaveExtraValues = TRUE;
				fSuccess = TRUE;
				*x = f1;
				*y = f2;
				this->m_ppExtraValues[0]->AddValue(f1, f3);
				this->m_ppExtraValues[1]->AddValue(f1, f4);
				this->m_ppExtraValues[2]->AddValue(f1, f5);
			}
		}
		else if (4 == this->m_nExtraValues)
		{
			if (6 == _stscanf_s(szOneLine, L"%f\t%f\t%f\t%f\t%f\t%f", &f1, &f2, &f3, &f4, &f5, &f6))
			{
				fHaveExtraValues = TRUE;
				fSuccess = TRUE;
				*x = f1;
				*y = f2;
				this->m_ppExtraValues[0]->AddValue(f1, f3);
				this->m_ppExtraValues[1]->AddValue(f1, f4);
				this->m_ppExtraValues[2]->AddValue(f1, f5);
				this->m_ppExtraValues[3]->AddValue(f1, f6);
			}
		}
	}
	if (!fHaveExtraValues)
	{
		if (2 == _stscanf_s(szOneLine, L"%f\t%f", &f1, &f2))
		{
			fSuccess = TRUE;
			*x = f1;
			*y = f2;
		}
	}
	return fSuccess;
}



/*
void CImpGratingScanInfo::LoadExtraValues(
	LPTSTR			szExtraValues)
{
	size_t			slen1, slen2;
	LPTSTR			szRem;			// remainder string
	LPTSTR			szTemp;
	TCHAR			szOneLine[MAX_PATH];
	float			f1;

	// clear the current values
	if (NULL != this->m_pExtraValues)
	{
		delete[] this->m_pExtraValues;
		this->m_pExtraValues = NULL;
	}
	this->m_nExtraValues = 0;
	szTemp = szExtraValues;
	while (NULL != szTemp)
	{
		szRem = StrStr(szTemp, L"\n");
		if (NULL != szRem)
		{
			StringCchLength(szTemp, STRSAFE_MAX_CCH, &slen1);
			StringCchLength(szRem, STRSAFE_MAX_CCH, &slen2);
			StringCchCopyN(szOneLine, MAX_PATH, szTemp, slen1 - slen2);
			if (1 == _stscanf_s(szOneLine, L"%f", &f1))
			{
				this->AddExtraValue(f1);
			}
		}
		else
		{
			StringCchCopy(szOneLine, MAX_PATH, szTemp);
			if (1 == _stscanf_s(szOneLine, L"%f", &f1))
			{
				this->AddExtraValue(f1);
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
}
*/


HRESULT CImpGratingScanInfo::saveToXML(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	BOOL			fSuccess = FALSE;
	CSciXML			xml;

	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_DISPATCH != pDispParams->rgvarg[0].vt ||
		NULL == pDispParams->rgvarg[0].pdispVal) return DISP_E_TYPEMISMATCH;
	if (xml.doInit(pDispParams->rgvarg[0].pdispVal))
	{
		fSuccess = this->saveToXML(&xml);
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

BOOL CImpGratingScanInfo::saveToXML(
	CSciXML		*	xml)
{
	IDispatch	*	pdispRoot = NULL;
	IDispatch	*	pdispNode = NULL;
	IDispatch	*	pdispChild = NULL;
	TCHAR			szValue[MAX_PATH];
	BOOL			fSuccess = FALSE;
	LPTSTR			szSpectralData = NULL;
	LPTSTR			szExtraValues = NULL;

	if (xml->GetrootNode(&pdispRoot))
	{
		if (xml->createNode(this->m_szNodeName, &pdispNode))
		{
			if (xml->createNode(L"grating", &pdispChild))
			{
				StringCchPrintf(szValue, MAX_PATH, L"%1d", this->m_grating);
				xml->setNodeValue(pdispChild, szValue);
				xml->appendChildNode(pdispNode, pdispChild);
				pdispChild->Release();
			}
			if (xml->createNode(L"filter", &pdispChild))
			{
				xml->setNodeValue(pdispChild, this->m_szFilter);
				xml->appendChildNode(pdispNode, pdispChild);
				pdispChild->Release();
			}
			if (xml->createNode(L"Detector", &pdispChild))
			{
				StringCchPrintf(szValue, MAX_PATH, L"%1d", this->m_detector);
				xml->setNodeValue(pdispChild, szValue);
				xml->appendChildNode(pdispNode, pdispChild);
				pdispChild->Release();
			}
			if (xml->createNode(L"Dispersion", &pdispChild))
			{
				_stprintf_s(szValue, L"%8.3f", this->m_dispersion);
				xml->setNodeValue(pdispChild, szValue);
				xml->appendChildNode(pdispNode, pdispChild);
				pdispChild->Release();
			}
			if (xml->createNode(L"SpectralData", &pdispChild))
			{
				if (this->writeSpectralData(&szSpectralData))
				{
					xml->setNodeValue(pdispChild, szSpectralData);
					xml->appendChildNode(pdispNode, pdispChild);
					CoTaskMemFree((LPVOID)szSpectralData);
				}
				pdispChild->Release();
			}
/*
			// extra values
			if (NULL != this->m_pExtraValues)
			{
				if (xml->createNode(L"ExtraValues", &pdispChild))
				{
					if (this->WriteExtraValues(&szExtraValues))
					{
						xml->setNodeValue(pdispChild, szExtraValues);
						xml->appendChildNode(pdispNode, pdispChild);
						CoTaskMemFree((LPVOID)szExtraValues);
					}
					pdispChild->Release();
				}
			}
*/
			// append node to root
			xml->appendChildNode(pdispRoot, pdispNode);
			fSuccess = TRUE;
			pdispNode->Release();
		}
		pdispRoot->Release();
	}
	return fSuccess;
}

BOOL CImpGratingScanInfo::writeSpectralData(
	LPTSTR		*	szSpectralData)
{
	*szSpectralData = NULL;
	if (this->m_nWaves != this->m_nSignals) return FALSE;
	if (0 == this->m_nWaves) return FALSE;
	HRESULT		hr;
	long		nAlloc = this->m_nWaves * MAX_PATH;
	LPTSTR		szTemp = (LPTSTR)CoTaskMemAlloc(nAlloc * sizeof(TCHAR));
	long		i;
	TCHAR		szOneLine[MAX_PATH];
	size_t		slen;
	BOOL		fSuccess = FALSE;

	szTemp[0] = '\0';
	StringCchCat(szTemp, nAlloc, L"\r\n");
//	StringCchCat(szTemp, nAlloc, L"\t  wave(nm)\t    Volts\r\n");
	this->WriteTitle(szOneLine, MAX_PATH);
	StringCchCat(szTemp, nAlloc, szOneLine);
	StringCchCat(szTemp, nAlloc, L"\r\n");
	for (i = 0; i < this->m_nWaves; i++)
	{
//		_stprintf_s(szOneLine, L"\t%8.3f\t%10.5g\r\n", 
//			this->m_pWavelengths[i], this->m_pSignals[i]);
		this->WriteData(i, szOneLine, MAX_PATH);
		StringCchLength(szTemp, nAlloc, &slen);
		if (slen > 0)
		{
			StringCchCat(szTemp, nAlloc, szOneLine);
		}
		else
		{
			StringCchCopy(szTemp, nAlloc, szOneLine);
		}
	}
	StringCchLength(szTemp, nAlloc, &slen);
	if (slen > 0)
	{
		*szSpectralData = (LPTSTR)CoTaskMemAlloc((slen + 1)* sizeof(TCHAR));
		hr = StringCchCopy(*szSpectralData, slen + 1, szTemp);
		fSuccess = SUCCEEDED(hr);
	}
	CoTaskMemFree((LPVOID)szTemp);
	return fSuccess;
}

BOOL CImpGratingScanInfo::WriteTitle(
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
	if (NULL == this->m_ppExtraValues)
	{
		StringCchPrintf(szTitle, nBufferSize, L"\t  wave(nm)\t    %s", this->m_szSignal);
		return TRUE;
	}
	else
	{
		long			i;
		TCHAR			szString[MAX_PATH];
		StringCchPrintf(szTitle, nBufferSize, L"\t  wave(nm)\t    %s", this->m_szSignal);
		for (i = 0; i < this->m_nExtraValues; i++)
		{
			this->m_ppExtraValues[i]->GetTitle(szString, MAX_PATH);
			StringCchCat(szTitle, nBufferSize, L"\t");
			StringCchCat(szTitle, nBufferSize, szString);
		}
		return TRUE;
	}
}

void CImpGratingScanInfo::WriteData(
	long			index,
	LPTSTR			szOneLine,
	UINT			nBufferSize)
{
	BOOL			fHaveExtraValues = FALSE;
	double			value;
	double			wavelength;
	double			tolerance = 0.01;
	TCHAR			szValue[MAX_PATH];
	long			i;
	if (this->m_nExtraValues > 0 && NULL != this->m_ppExtraValues)
	{
		wavelength = this->m_pWavelengths[index];
		if (this->m_ppExtraValues[0]->doFindValue(wavelength, tolerance, &value))
		{
			fHaveExtraValues = TRUE;
			_stprintf_s(szOneLine, nBufferSize, L"\t%8.3f\t%10.5g", this->m_pWavelengths[index], this->m_pSignals[index]);
			_stprintf_s(szValue, L"%10.5g", value);
			StringCchCat(szOneLine, nBufferSize, L"\t");
			StringCchCat(szOneLine, nBufferSize, szValue);
			for (i = 1; i < this->m_nExtraValues; i++)
			{
				if (!this->m_ppExtraValues[i]->doFindValue(wavelength, tolerance, &value))
				{
					value = 0.0;
				}
				_stprintf_s(szValue, L"%10.5g", value);
				StringCchCat(szOneLine, nBufferSize, L"\t");
				StringCchCat(szOneLine, nBufferSize, szValue);
			}
			StringCchCat(szOneLine, nBufferSize, L"\r\n");
		}
	}
	if (!fHaveExtraValues)
	{
		_stprintf_s(szOneLine, nBufferSize, L"\t%8.3f\t%10.5g\r\n", this->m_pWavelengths[index], this->m_pSignals[index]);
	}
}



/*
BOOL CImpGratingScanInfo::WriteExtraValues(
	LPTSTR		*	szExtraValues)
{
	*szExtraValues = NULL;
	if (NULL == this->m_pExtraValues || 0 == this->m_nExtraValues) return FALSE;
	HRESULT		hr;
	long		nAlloc = this->m_nExtraValues * MAX_PATH;
	LPTSTR		szTemp = (LPTSTR)CoTaskMemAlloc(nAlloc * sizeof(TCHAR));
	long		i;
	TCHAR		szOneLine[MAX_PATH];
	size_t		slen;
	BOOL		fSuccess = FALSE;

	szTemp[0] = '\0';
	for (i = 0; i < this->m_nExtraValues; i++)
	{
		_stprintf_s(szOneLine, L"%10.5g\r\n", this->m_pExtraValues[i]);
		StringCchLength(szTemp, nAlloc, &slen);
		if (slen > 0)
		{
			StringCchCat(szTemp, nAlloc, szOneLine);
		}
		else
		{
			StringCchCopy(szTemp, nAlloc, szOneLine);
		}
	}
	StringCchLength(szTemp, nAlloc, &slen);
	if (slen > 0)
	{
		*szExtraValues = (LPTSTR)CoTaskMemAlloc((slen + 1)* sizeof(TCHAR));
		hr = StringCchCopy(*szExtraValues, slen + 1, szTemp);
		fSuccess = SUCCEEDED(hr);
	}
	CoTaskMemFree((LPVOID)szTemp);
	return fSuccess;
}
*/

HRESULT CImpGratingScanInfo::addData(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	double				wave;
	double				signal;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	wave = varg.dblVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	signal = varg.dblVal;

	if (NULL == this->m_pWavelengths)
	{
		this->m_pWavelengths = new double[1];
		this->m_pSignals = new double[1];
		this->m_pWavelengths[0] = wave;
		this->m_nWaves = 1;
		this->m_pSignals[0] = signal;
		this->m_nSignals = 1;
	}
	else
	{
		if (this->m_nWaves == this->m_nSignals)
		{
			long		nArray = this->m_nWaves + 1;
			double	*	pX = new double[nArray];
			double	*	pY = new double[nArray];
			long		i;
			for (i = 0; i < this->m_nWaves; i++)
			{
				pX[i] = this->m_pWavelengths[i];
				pY[i] = this->m_pSignals[i];
			}
			pX[i] = wave;
			pY[i] = signal;
			delete[] this->m_pWavelengths;
			delete[] this->m_pSignals;
			this->m_pWavelengths = pX;
			this->m_pSignals = pY;
			this->m_nWaves = nArray;
			this->m_nSignals = nArray;
		}
	}
	return S_OK;
}

HRESULT CImpGratingScanInfo::ClearData()
{
	if (NULL != this->m_pWavelengths)
	{
		delete[] this->m_pWavelengths;
		this->m_pWavelengths = NULL;
	}
	this->m_nWaves = 0;
	if (NULL != this->m_pSignals)
	{
		delete[] this->m_pSignals;
		this->m_pSignals = NULL;
	}
	this->m_nSignals = 0;
	return S_OK;
}

HRESULT CImpGratingScanInfo::clearDirty()
{
	this->m_amDirty = FALSE;
	return S_OK;
}

HRESULT CImpGratingScanInfo::findIndex(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	double			wavelength;		// input wavelength
	double			tolerance;		// wavelength tolerance
	double			tMin, tMax;		// test range for wavelength value
	long			i;
	long			index = -1;		// initialize with failure value
	BOOL			fDone;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	wavelength = varg.dblVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	tolerance = varg.dblVal;
	tMin = wavelength - tolerance;
	tMax = wavelength + tolerance;
	i = 0;
	fDone = FALSE;
	while (i < this->m_nWaves && !fDone)
	{
		if (this->m_pWavelengths[i] >= tMin && this->m_pWavelengths[i] <= tMax)
		{
			fDone = TRUE;		// have found the index
			index = i;
		}
		if (!fDone) i++;
	}
	InitVariantFromInt32(index, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetNodeName(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szNodeName, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetNodeName(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szNodeName, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetGrating(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_grating, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetGrating(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_grating = varg.lVal;
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetFilter(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szFilter, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetFilter(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szFilter, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetDispersion(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_dispersion, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetDispersion(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_dispersion = varg.dblVal;
	return S_OK;
}
HRESULT CImpGratingScanInfo::GetWavelengths(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pWavelengths, this->m_nWaves, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetWavelengths(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long		ub;
	long		i;
	double	*	arr;
	SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
	SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
	if (NULL != this->m_pWavelengths)
	{
		delete[] this->m_pWavelengths;
	}
	this->m_nWaves = ub + 1;
	this->m_pWavelengths = new double[this->m_nWaves];
	for (i = 0; i < this->m_nWaves; i++)
		this->m_pWavelengths[i] = arr[i];
	SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetSignals(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pSignals, this->m_nSignals, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::SetSignals(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long		ub;
	long		i;
	double	*	arr;
	SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
	SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
	if (NULL != this->m_pSignals)
	{
		delete[] this->m_pSignals;
	}
	this->m_nSignals = ub + 1;
	this->m_pSignals = new double[this->m_nSignals];
	for (i = 0; i < this->m_nSignals; i++)
		this->m_pSignals[i] = arr[i];
	SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetAmDirty(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_amDirty, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::GetarraySize(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_nSignals, pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::getValue(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (varg.lVal >= 0 && varg.lVal < this->m_nSignals)
		InitVariantFromDouble(this->m_pSignals[varg.lVal], pVarResult);
	return S_OK;
}

HRESULT CImpGratingScanInfo::setValue(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	long			index;
	double			value;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	index = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	value = varg.dblVal;
	if (index >= 0 && index < this->m_nSignals)
		this->m_pSignals[index] = value;
	return S_OK;
}

HRESULT CImpGratingScanInfo::setExtraValue(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	TCHAR			szTitle[MAX_PATH];
	double			wavelength;
	double			extraValue;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szTitle, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	wavelength = varg.dblVal;
	hr = DispGetParam(pDispParams, 2, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	extraValue = varg.dblVal;
	// add the extra value
	this->AddExtraValue(szTitle, wavelength, extraValue);
	return S_OK;
}

void CImpGratingScanInfo::AddExtraValue(
	LPCTSTR			szTitle,
	double			wavelength,
	double			extraValue)
{
	CExtraValues * pExtraValues = NULL;
	if (!this->FindExtraValueComponent(szTitle, &pExtraValues))
	{
		pExtraValues = new CExtraValues;
		pExtraValues->SetTitle(szTitle);
		pExtraValues->AddValue(wavelength, extraValue);
		// add to the end array
		if (NULL == this->m_ppExtraValues)
		{
			this->m_ppExtraValues = new CExtraValues *[1];
			this->m_ppExtraValues[0] = pExtraValues;
			this->m_nExtraValues = 1;
		}
		else
		{
			long			nAlloc = this->m_nExtraValues + 1;
			CExtraValues**	ppExtraValues = new CExtraValues * [nAlloc];
			long			i;
			for (i = 0; i < this->m_nExtraValues; i++)
				ppExtraValues[i] = this->m_ppExtraValues[i];
			ppExtraValues[nAlloc - 1] = pExtraValues;
			delete[] this->m_ppExtraValues;
			this->m_ppExtraValues = ppExtraValues;
			this->m_nExtraValues = nAlloc;
		}
	}
	else
	{
		// have found the desired component
		pExtraValues->AddValue(wavelength, extraValue);
	}
}

/*
void CImpGratingScanInfo::AddExtraValue(
	double			extraValue)
{
	if (NULL == this->m_pExtraValues)
	{
		this->m_pExtraValues = new double[1];
		this->m_pExtraValues[0] = extraValue;
		this->m_nExtraValues = 0;
	}
	else
	{
		long		nValues = this->m_nExtraValues + 1;
		long		i;
		double	*	tValues = new double[nValues];
		for (i = 0; i < this->m_nExtraValues; i++)
			tValues[i] = this->m_pExtraValues[i];
		tValues[nValues - 1] = extraValue;
		delete[] this->m_pExtraValues;
		this->m_pExtraValues = tValues;
		this->m_nExtraValues = nValues;
	}
}

*/
/*
HRESULT CImpGratingScanInfo::GetextraValues(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (NULL != this->m_pExtraValues)
	{
		InitVariantFromDoubleArray(this->m_pExtraValues, this->m_nExtraValues,
			pVarResult);
	}
	return S_OK;
}
*/

HRESULT CImpGratingScanInfo::GethaveExtraValue(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_nExtraValues > 0, pVarResult);
	return S_OK;
}

BOOL CImpGratingScanInfo::FindExtraValueComponent(
	LPCTSTR				szTitle,
	CExtraValues	**	ppExtraValues)
{
	long			i = 0;
	TCHAR			szExtraValue[MAX_PATH];
	BOOL			fDone = FALSE;

	*ppExtraValues = NULL;
	while (i < this->m_nExtraValues && !fDone)
	{
		if (this->m_ppExtraValues[i]->GetTitle(szExtraValue, MAX_PATH) &&
			0 == lstrcmpi(szExtraValue, szTitle))
		{
			fDone = TRUE;
			*ppExtraValues = this->m_ppExtraValues[i];
		}
		if (!fDone) i++;
	}
	return fDone;
}

void CImpGratingScanInfo::ClearExtraValues()
{
	if (NULL != this->m_ppExtraValues)
	{
		long		i;
		for (i = 0; i < this->m_nExtraValues; i++)
		{
			delete this->m_ppExtraValues[i];
			this->m_ppExtraValues[i] = NULL;
		}
		delete[] this->m_ppExtraValues;
		this->m_ppExtraValues = NULL;
	}
	this->m_nExtraValues = 0;
}

HRESULT CImpGratingScanInfo::GetExtraValuesTitles(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_nExtraValues > 0 && NULL != this->m_ppExtraValues)
	{
		SAFEARRAYBOUND		sab;
		BSTR			*	arr;
		long				i;
		TCHAR				szTitle[MAX_PATH];
		sab.lLbound = 0;
		sab.cElements = this->m_nExtraValues;
		VariantInit(pVarResult);
		pVarResult->vt = VT_ARRAY | VT_BSTR;
		pVarResult->parray = SafeArrayCreate(VT_BSTR, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**)&arr);
		for (i = 0; i < this->m_nExtraValues; i++)
		{
			this->m_ppExtraValues[i]->GetTitle(szTitle, MAX_PATH);
			arr[i] = SysAllocString(szTitle);
		}
		SafeArrayUnaccessData(pVarResult->parray);
	}
	return S_OK;
}
HRESULT CImpGratingScanInfo::GetSignalUnits(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szSignal, pVarResult);
	return S_OK;
}
HRESULT CImpGratingScanInfo::SetSignalUnits(
	DISPPARAMS		*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szSignal, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}
HRESULT CImpGratingScanInfo::GetExtraValueComponent(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	CExtraValues*	pExtraValues;
	long			nValues;
	double		*	pValues = NULL;

	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (this->FindExtraValueComponent(varg.bstrVal, &pExtraValues))
	{
		if (pExtraValues->GetExtraValues(&nValues, &pValues))
		{
			InitVariantFromDoubleArray(pValues, nValues, pVarResult);
			delete[] pValues;
		}
	}
	VariantClear(&varg);
	return S_OK;
}
HRESULT CImpGratingScanInfo::GetExtraValueWavelengths(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	CExtraValues*	pExtraValues;
	long			nValues;
	double		*	pWaves;

	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (this->FindExtraValueComponent(varg.bstrVal, &pExtraValues))
	{
		if (pExtraValues->GetWavelengths(&nValues, &pWaves))
		{
			InitVariantFromDoubleArray(pWaves, nValues, pVarResult);
			delete[] pWaves;
		}
	}
	VariantClear(&varg);
	return S_OK;
}

CImpGratingScanInfo::CExtraValues::CExtraValues()
{
	this->m_szTitle[0] = '\0';
	this->m_ExtraValues.clear();
}

CImpGratingScanInfo::CExtraValues::~CExtraValues()
{
	this->m_ExtraValues.clear();
}
void CImpGratingScanInfo::CExtraValues::SetTitle(
	LPCTSTR				szTitle)
{
	StringCchCopy(this->m_szTitle, MAX_PATH, szTitle);
}
BOOL CImpGratingScanInfo::CExtraValues::GetTitle(
	LPTSTR				szTitle,
	UINT				nBufferSize)
{
	StringCchCopy(szTitle, nBufferSize, this->m_szTitle);
	return TRUE;
}
void CImpGratingScanInfo::CExtraValues::AddValue(
	double				wavelength,
	double				newValue)
{
	long			index = this->FindIndex(wavelength, 0.001);
	if (index >= 0)
	{
		this->m_ExtraValues.at(index).AddSignal(newValue);
	}
	else
	{
		CSpectralData		spectralData(wavelength, newValue);
		this->m_ExtraValues.push_back(spectralData);
	}
}
BOOL CImpGratingScanInfo::CExtraValues::doFindValue(
	double				wavelength,
	double				tolerance,
	double			*	value)
{
	long			index = this->FindIndex(wavelength, tolerance);
	if (index >= 0)
	{
		*value = this->m_ExtraValues.at(index).GetSignal();
		return TRUE;
	}
	else
	{
		*value = 0.0;
		return FALSE;
	}
}

long CImpGratingScanInfo::CExtraValues::FindIndex(
	double				wavelength,
	double				tolerance)
{
	long			i = 0;
	BOOL			fDone = FALSE;
	double			tMin = wavelength - tolerance;
	double			tMax = wavelength + tolerance;
	long			retVal = -1;
	long			nValues = this->m_ExtraValues.size();
	double			wave;
	while (i < nValues && !fDone)
	{
		wave = this->m_ExtraValues.at(i).GetWavelength();
		if (wave >= tMin && wave <= tMax)
		{
			fDone = TRUE;
			retVal = i;
		}
		if (!fDone) i++;
	}
	return retVal;
}

long CImpGratingScanInfo::CExtraValues::GetNValues()
{
	return this->m_ExtraValues.size();
}
BOOL CImpGratingScanInfo::CExtraValues::GetWavelengths(
	long			*	nValues,
	double			**	ppWavelengths)
{
	long		nSize = this->m_ExtraValues.size();
	long		i;
	if (nSize > 0)
	{
		*nValues = nSize;
		*ppWavelengths = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
		{
			(*ppWavelengths)[i] = this->m_ExtraValues.at(i).GetWavelength();
		}
		return TRUE;
	}
	else
	{
		*nValues = 0;
		*ppWavelengths = NULL;
		return FALSE;
	}
}
BOOL CImpGratingScanInfo::CExtraValues::GetExtraValues(
	long			*	nValues,
	double			**	ppExtraValues)
{
	long			i;
	*nValues = this->m_ExtraValues.size();
	if ((*nValues) > 0)
	{
		*ppExtraValues = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
			(*ppExtraValues)[i] = this->m_ExtraValues.at(i).GetSignal();
		return TRUE;
	}
	else
	{
		*ppExtraValues = NULL;
		return FALSE;
	}
}

HRESULT CImpGratingScanInfo::GetDetector(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_detector, pVarResult);
	return S_OK;
}
HRESULT CImpGratingScanInfo::SetDetector(
	DISPPARAMS		*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_detector = varg.lVal;
	return S_OK;
}

CImpGratingScanInfo::CSpectralData::CSpectralData() :
	m_Wavelength(0.0)
{
	this->m_signals.clear();
}
CImpGratingScanInfo::CSpectralData::CSpectralData(const CSpectralData& src) :
	m_Wavelength(src.m_Wavelength)
{
	m_signals.clear();
	m_signals = src.m_signals;
}
CImpGratingScanInfo::CSpectralData::CSpectralData(double wavelength, double signal) :
	m_Wavelength(wavelength)
{
	m_signals.clear();
	m_signals.push_back(signal);
}
CImpGratingScanInfo::CSpectralData::~CSpectralData()
{
	m_signals.clear();
}
const CImpGratingScanInfo::CSpectralData& CImpGratingScanInfo::CSpectralData::operator=(const CSpectralData& src)
{
	if (this != &src)
	{
		this->m_Wavelength = src.m_Wavelength;
		this->m_signals.clear();
		this->m_signals = src.m_signals;
	}
	return *this;
}
double CImpGratingScanInfo::CSpectralData::GetWavelength()
{
	return this->m_Wavelength;
}
void CImpGratingScanInfo::CSpectralData::SetWavelength(
	double			wavelength)
{
	this->m_Wavelength = wavelength;
}
double CImpGratingScanInfo::CSpectralData::GetSignal()
{
	long		nSignals = this->m_signals.size();
	double		total = 0.0;
	long		i;

	if (nSignals > 0)
	{
		for (i = 0; i < nSignals; i++)
			total += (this->m_signals.at(i));
		return total / nSignals;
	}
	else
	{
		return 0.0;
	}
}
void CImpGratingScanInfo::CSpectralData::AddSignal(
	double			signal)
{
	this->m_signals.push_back(signal);
}