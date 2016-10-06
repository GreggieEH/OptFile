#include "stdafx.h"
#include "ImpICalibrationStandard.h"
#include "MyGuids.h"
#include "Server.h"
//#include "DynamicArray.h"
#include "SciOutFile.h"
#include "SciXML.h"
#include "dispids.h"

CImpICalibrationStandard::CImpICalibrationStandard() :
	m_cRefs(0),
//	m_nArray(0),
//	m_pWaves(NULL),
//	m_pCalibration(NULL),
	m_expectedDistance(20.0),			// expected distance in cm
	m_sourceDistance(20.0)				// source distance in cm
{
	this->m_szFileName[0] = '\0';
	m_Data.clear();
}

CImpICalibrationStandard::~CImpICalibrationStandard()
{
	this->ClearData();
}

// IUnknown methods
STDMETHODIMP CImpICalibrationStandard::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || IID_ICalibrationStandard == riid)
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
STDMETHODIMP_(ULONG) CImpICalibrationStandard::AddRef()
{
	return ++m_cRefs;
}
STDMETHODIMP_(ULONG) CImpICalibrationStandard::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}
// IDispatch methods
STDMETHODIMP CImpICalibrationStandard::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}
STDMETHODIMP CImpICalibrationStandard::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	HRESULT			hr;
	ITypeLib	*	pTypeLib = NULL;
	*ppTInfo = NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(IID_ICalibrationStandard, ppTInfo);
		pTypeLib->Release();
	}
	return hr;

}
STDMETHODIMP CImpICalibrationStandard::GetIDsOfNames(
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
STDMETHODIMP CImpICalibrationStandard::Invoke(
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
	case DISPID_CalibrationStandard_amLoaded:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetamLoaded(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetamLoaded(pDispParams);
		break;
	case DISPID_CalibrationStandard_fileName:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetfileName(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetfileName(pDispParams);
		break;
	case DISPID_CalibrationStandard_wavelengths:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->Getwavelengths(pVarResult);
		break;
	case DISPID_CalibrationStandard_calibration:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->Getcalibration(pVarResult);
		break;
	case DISPID_CalibrationStandard_loadFromXML:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->loadFromXML(pDispParams, pVarResult);
		break;
	case DISPID_CalibrationStandard_saveToXML:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->saveToXML(pDispParams, pVarResult);
		break;
	case DISPID_CalibrationStandard_loadExcelFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->loadExcelFile(pDispParams, pVarResult);
		break;
	case DISPID_CalibrationStandard_checkExcelFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->checkExcelFile(pDispParams, pVarResult);
		break;
	case DISPID_CalibrationStandard_expectedDistance:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetexpectedDistance(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetexpectedDistance(pDispParams);
		break;
	case DISPID_CalibrationStandard_sourceDistance:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetsourceDistance(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetsourceDistance(pDispParams);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CImpICalibrationStandard::GetamLoaded(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->GetamLoaded(), pVarResult);
	return S_OK;
}
HRESULT CImpICalibrationStandard::SetamLoaded(
	DISPPARAMS		*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->SetamLoaded(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}
HRESULT CImpICalibrationStandard::GetfileName(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szFileName, pVarResult);
	return S_OK;
}
HRESULT CImpICalibrationStandard::SetfileName(
	DISPPARAMS		*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->SetfileName(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}
HRESULT CImpICalibrationStandard::Getwavelengths(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long		nX;
	double	*	pX;
	if (this->Getwavelengths(&nX, &pX))
	{
		InitVariantFromDoubleArray(pX, nX, pVarResult);
		delete[] pX;
	}
	return S_OK;
}
HRESULT CImpICalibrationStandard::Getcalibration(
	VARIANT			*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long		nY;
	double	*	pY;
	if (this->Getcalibration(&nY, &pY))
	{
		InitVariantFromDoubleArray(pY, nY, pVarResult);
		delete[] pY;
	}
	return S_OK;
}
HRESULT CImpICalibrationStandard::loadFromXML(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	BOOL		fSuccess;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_DISPATCH != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	fSuccess = this->loadFromXML(pDispParams->rgvarg[0].pdispVal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}
HRESULT CImpICalibrationStandard::saveToXML(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	BOOL		fSuccess;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_DISPATCH != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	fSuccess = this->saveToXML(pDispParams->rgvarg[0].pdispVal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}
HRESULT CImpICalibrationStandard::loadExcelFile(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->loadExcelFile(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}
HRESULT	CImpICalibrationStandard::checkExcelFile(
	DISPPARAMS		*	pDispParams,
	VARIANT			*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	if (NULL == pVarResult) return E_INVALIDARG;
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
InitVariantFromBoolean(this->checkExcelFile(varg.bstrVal), pVarResult);
VariantClear(&varg);
return S_OK;
}

void CImpICalibrationStandard::ClearData()
{
	this->m_Data.clear();
/*
	if (NULL != this->m_pWaves)
	{
		delete[] this->m_pWaves;
		this->m_pWaves = NULL;
	}
	if (NULL != this->m_pCalibration)
	{
		delete[] this->m_pCalibration;
		this->m_pCalibration = NULL;
	}
	this->m_nArray = 0;
*/
}

BOOL CImpICalibrationStandard::loadFromXML(
	IDispatch	*	pdispxml)
{
	/*
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;

	TCHAR			szFileName[MAX_PATH];

	VariantInit(&varg);
	varg.vt		= VT_DISPATCH;
	varg.pdispVal	= xml;
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
	*/
	/*
	var root = xml.rootNode;
	var fSuccess = false;
	var node;
	var child;
	if (null != root) {
	node = xml.findNamedChild(root, "CalibrationStandard");
	if (null != node) {
	fSuccess = true;
	child = xml.findNamedChild(node, "fileName");
	if (null != child) {
	put_fileName(xml.getNodeValue(child));
	}
	}
	}
	return fSuccess;

	*/
	//	HRESULT			hr;
	CSciXML			xml;
	IDispatch	*	pdispRoot = NULL;		// root node
	BOOL			fSuccess = FALSE;
	IDispatch	*	pdispNode;
	IDispatch	*	pdispChild;
	TCHAR			szValue[MAX_PATH];
	float			f1;

	if (xml.doInit(pdispxml))
	{
		if (xml.GetrootNode(&pdispRoot))
		{
			if (xml.findNamedChild(pdispRoot, L"CalibrationStandard", &pdispNode))
			{
				if (xml.findNamedChild(pdispNode, L"fileName", &pdispChild))
				{
					fSuccess = xml.getNodeValue(pdispChild, this->m_szFileName, MAX_PATH);
					pdispChild->Release();
				}
				if (xml.findNamedChild(pdispNode, L"expectedDistance", &pdispChild))
				{
					if (xml.getNodeValue(pdispChild, szValue, MAX_PATH))
					{
						if (1 == _stscanf_s(szValue, L"%f", &f1))
						{
							this->m_expectedDistance = f1;
						}
					}
					pdispChild->Release();
				}
				if (xml.findNamedChild(pdispNode, L"sourceDistance", &pdispChild))
				{
					if (xml.getNodeValue(pdispChild, szValue, MAX_PATH))
					{
						if (1 == _stscanf_s(szValue, L"%f", &f1))
						{
							this->m_sourceDistance = f1;
						}
					}
					pdispChild->Release();
				}
				pdispNode->Release();
			}
			pdispRoot->Release();
		}
	}
	return fSuccess;
}

BOOL CImpICalibrationStandard::saveToXML(
	IDispatch	*	pdispxml)
{
	/*
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
	*/
	/*		var root = xml.rootNode;
	var node;
	var child;
	node = xml.createNode("CalibrationStandard");
	child = xml.createNode("fileName");
	xml.setNodeValue(child, get_fileName());
	xml.appendChildNode(node, child);
	xml.appendChildNode(root, node);
	return true;
	*/
	CSciXML			xml;
	BOOL			fSuccess = FALSE;
	IDispatch	*	pdispRoot;
	IDispatch	*	pdispNode;
	IDispatch	*	pdispChild;
	TCHAR			szValue[MAX_PATH];

	if (xml.doInit(pdispxml))
	{
		if (xml.GetrootNode(&pdispRoot))
		{
			if (xml.createNode(L"CalibrationStandard", &pdispNode))
			{
				if (xml.createNode(L"fileName", &pdispChild))
				{
					xml.setNodeValue(pdispChild, this->m_szFileName);
					xml.appendChildNode(pdispNode, pdispChild);
					xml.appendChildNode(pdispRoot, pdispNode);
					fSuccess = TRUE;
					pdispChild->Release();
				}
				if (xml.createNode(L"expectedDistance", &pdispChild))
				{
					_stprintf_s(szValue, L"%8.4f cm", this->m_expectedDistance);
					xml.setNodeValue(pdispChild, szValue);
					xml.appendChildNode(pdispNode, pdispChild);
					xml.appendChildNode(pdispRoot, pdispNode);
					pdispChild->Release();
				}
				if (xml.createNode(L"sourceDistance", &pdispChild))
				{
					_stprintf_s(szValue, L"%8.4f cm", this->m_sourceDistance);
					xml.setNodeValue(pdispChild, szValue);
					xml.appendChildNode(pdispNode, pdispChild);
					xml.appendChildNode(pdispRoot, pdispNode);
					pdispChild->Release();
				}
				pdispNode->Release();
			}
			pdispRoot->Release();
		}
	}
	return fSuccess;
}

BOOL CImpICalibrationStandard::GetfileName(
	LPTSTR			szFileName,
	UINT			nBufferSize)
{
	if ('\0' != this->m_szFileName[0])
	{
		StringCchCopy(szFileName, nBufferSize, this->m_szFileName);
		return TRUE;
	}
	else
	{
		szFileName[0] = '\0';
		return FALSE;
	}
}

void CImpICalibrationStandard::SetfileName(
	LPCTSTR			szfileName)
{
	StringCchCopy(this->m_szFileName, MAX_PATH, szfileName);
}

BOOL CImpICalibrationStandard::GetamLoaded()
{
	return this->m_Data.size() > 0;
//	return NULL != this->m_pWaves && NULL != this->m_pCalibration && this->m_nArray > 0;
}

void CImpICalibrationStandard::SetamLoaded(
	BOOL			amLoaded)
{
	this->ClearData();		// clear current data
	if (!amLoaded)
	{
		return;
	}
	HRESULT				hr;
	CLSID				clsid;
	IDispatch		*	pdisp		= NULL;
	DISPID				dispid;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	double			*	pX		= NULL;
	ULONG				nX		= 0;
	double			*	pY		= NULL;
	ULONG				nY		= 0;
	ULONG				i;

	// use SciDataSet to load the calibration standard
	hr = CLSIDFromProgID(L"Sciencetech.SciDataSet.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, L"FileName", &dispid);
		Utils_SetStringProperty(pdisp, dispid, this->m_szFileName);
		// read the file
		Utils_GetMemid(pdisp, L"ReadDataFile", &dispid);
		hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		// get the data arrays
		if (fSuccess)
		{
			// x values
			Utils_GetMemid(pdisp, L"GetWavelengths", &dispid);
			hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
			if (SUCCEEDED(hr))
			{
				hr = VariantToDoubleArrayAlloc(varResult, &pX, &nX);
				VariantClear(&varResult);
			}
			if (NULL != pX)
			{
				Utils_GetMemid(pdisp, L"GetSpectra", &dispid);
				hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
				if (SUCCEEDED(hr))
				{
					hr = VariantToDoubleArrayAlloc(varResult, &pY, &nY);
					VariantClear(&varResult);
				}
				if (NULL != pY)
				{
					if (nY == nX)
					{
						for (i = 0; i < nX; i++)
							this->AddValue(pX[i], pY[i]);
					}
					CoTaskMemFree((LPVOID)pY);
				}
				CoTaskMemFree((LPVOID)pX);
			}
		}
		pdisp->Release();
	}
	

	/*
	if (this->checkExcelFile(this->m_szFileName))
	{
		this->loadExcelFile(this->m_szFileName);
	}
	else
	{
		// use SciOutFile to read in
		CSciOutFile			outFile;
		long				nX;
		double			*	pX;
		long				nY;
		double			*	pY;
		long				i;
		outFile.SetfileName(this->m_szFileName);
		outFile.loadFile();
		if (outFile.GetxValues(&nX, &pX))
		{
			if (outFile.GetyValues(&nY, &pY))
			{
				if (nX == nY)
				{
					for (i = 0; i < nX; i++)
					{
						this->AddValue(pX[i], pY[i]);
					}
//					this->m_nArray = nX;
//					this->m_pWaves = new double[this->m_nArray];
//					this->m_pCalibration = new double[this->m_nArray];
//					for (i = 0; i < this->m_nArray; i++)
//					{
//						this->m_pWaves[i] = pX[i];
//						this->m_pCalibration[i] = pY[i];
//					}
				}
				delete[] pY;
			}
			delete[] pX;
		}
	}
	*/
}

void CImpICalibrationStandard::AddValue(
	double			x,
	double			y)
{

	CDataPoint		dataPoint(x, y);
	this->m_Data.push_back(dataPoint);
}


BOOL CImpICalibrationStandard::Getwavelengths(
	long		*	nValues,
	double		**	ppWaves)
{
	/*
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*nValues		= 0;
	*ppWaves		= NULL;
	hr = this->getProperty(TEXT("wavelengths"), &varResult);
	if (SUCCEEDED(hr))
	{
	fSuccess = TranslateJScriptArray(&varResult, nValues, ppWaves);
	VariantClear(&varResult);
	}
	return fSuccess;
	*/
	if (this->GetamLoaded())
	{
		long		i;
		*nValues = this->m_Data.size();
		*ppWaves = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
			(*ppWaves)[i] = this->m_Data.at(i).GetX();
		return TRUE;
	}
	else
	{
		*nValues = 0;
		*ppWaves = NULL;
		return FALSE;
	}
}

BOOL CImpICalibrationStandard::Getcalibration(
	long		*	nValues,
	double		**	ppcalibration)
{
	/*
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*nValues		= 0;
	*ppcalibration	= NULL;
	hr = this->getProperty(TEXT("calibration"), &varResult);
	if (SUCCEEDED(hr))
	{
	fSuccess = TranslateJScriptArray(&varResult, nValues, ppcalibration);
	VariantClear(&varResult);
	}
	return fSuccess;
	*/
	if (this->GetamLoaded())
	{
		long		i;
		*nValues = this->m_Data.size();
		*ppcalibration = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
			(*ppcalibration)[i] = this->m_Data.at(i).GetY();				// this->m_pCalibration[i];
		return TRUE;
	}
	else
	{
		*nValues = 0;
		*ppcalibration = NULL;
		return FALSE;
	}
}

BOOL CImpICalibrationStandard::loadExcelFile(
	LPCTSTR			szFileName)
{
	HRESULT				hr;
	CLSID				clsid;
	IDispatch		*	pdispExcel = NULL;
	IDispatch		*	pdispWorkbooks = NULL;
//	IDispatch		*	pdispWorkbook = NULL;
	VARIANT				Worksheets;
	VARIANT				Worksheet;
//	IDispatch		*	pdispWorkSheet = NULL;
	DISPID				dispid;
	VARIANT				Cells;
	IDispatch		*	pdispCells = NULL;
	VARIANTARG			varg;
	VARIANT				varResult;
	VARIANT				Value;
	VARIANTARG			avarg[2];
	long				i;		//, j;
	BOOL				fDone;
	double				xvalue, yvalue;
//	VARIANT				Value;
//	CDynamicArray		dynArray;			// dynamic array
//	BOOL				fSuccess = FALSE;
//	long				nX;
//	long				nY;
//	double			*	pX;
//	double			*	pY;
	VARIANT				Range;
//	IDispatch		*	pdispRange;

	this->ClearData();			// clear current data
	return FALSE;
	hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pdispExcel);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdispExcel, L"Workbooks", &dispid);
		hr = Utils_InvokePropertyGet(pdispExcel, dispid, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
			{
				pdispWorkbooks = varResult.pdispVal;
				varResult.pdispVal->AddRef();
				Utils_GetMemid(varResult.pdispVal, L"Open", &dispid);
				InitVariantFromString(szFileName, &varg);
				hr = Utils_InvokeMethod(varResult.pdispVal, dispid, &varg, 1, NULL);
				VariantClear(&varg);
			}
			VariantClear(&varResult);
		}
		Utils_GetMemid(pdispWorkbooks, L"Item", &dispid);
		InitVariantFromInt32(1, &varg);
		hr = Utils_InvokePropertyGet(pdispWorkbooks, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
			{
				// varResult.pdispVal is first workbook
				Utils_GetMemid(varResult.pdispVal, L"Worksheets", &dispid);
				hr = Utils_InvokePropertyGet(varResult.pdispVal, dispid, NULL, 0, &Worksheets);
				if (SUCCEEDED(hr))
				{
					if (VT_DISPATCH == Worksheets.vt && NULL != Worksheets.pdispVal)
					{
						Utils_GetMemid(Worksheets.pdispVal, L"Item", &dispid);
						InitVariantFromInt32(1, &varg);
						hr = Utils_InvokePropertyGet(Worksheets.pdispVal, dispid, &varg, 1, &Worksheet);
						if (SUCCEEDED(hr))
						{
							if (VT_DISPATCH == Worksheet.vt && NULL != Worksheet.pdispVal)
							{
								Utils_GetMemid(Worksheet.pdispVal, L"Cells", &dispid);
								hr = Utils_InvokePropertyGet(Worksheet.pdispVal, dispid, NULL, 0, &Cells);
								if (SUCCEEDED(hr))
								{
									if (VT_DISPATCH == Cells.vt && NULL != Cells.ppdispVal)
									{
										pdispCells = Cells.pdispVal;
										pdispCells->AddRef();
									}
									VariantClear(&Cells);
								}
							}
							VariantClear(&Worksheet);
						}
					}
					VariantClear(&Worksheets);
				}
			}
			VariantClear(&varResult);
		}
		if (NULL != pdispCells)
		{
			Utils_GetMemid(pdispCells, L"Range", &dispid);
			InitVariantFromString(L"A3", &avarg[1]);
			InitVariantFromString(L"B65535", &avarg[0]);
			hr = Utils_InvokePropertyGet(pdispCells, dispid, avarg, 2, &Range);
			VariantClear(&avarg[1]);
			VariantClear(&avarg[0]);
			if (SUCCEEDED(hr))
			{
				if (VT_DISPATCH == Range.vt && NULL != Range.pdispVal)
				{
					Utils_GetMemid(Range.pdispVal, L"Value", &dispid);
					Utils_PutOptional(&varg);
					hr = Utils_DoInvoke(Range.pdispVal, dispid, DISPATCH_PROPERTYGET, &varg, 1, &varResult);
					VariantClear(&varg);
					if (SUCCEEDED(hr))
					{
						i = 1;
						long		arr[2];
						fDone = FALSE;
						while (!fDone)
						{
							arr[0] = i;
							arr[1] = 1;
							SafeArrayGetElement(varResult.parray, arr, &Value);
							if (VT_EMPTY != Value.vt)
							{
								xvalue = Value.dblVal;
								arr[1] = 2;
								SafeArrayGetElement(varResult.parray, arr, &Value);
								if (VT_EMPTY != Value.vt)
								{
									yvalue = Value.dblVal;
									this->AddValue(xvalue, yvalue);
								}
								else fDone = TRUE;
							}
							else fDone = TRUE;
							if (!fDone)i++;
						}
						VariantClear(&varResult);
					}
				}
				VariantClear(&Range);
			}
			pdispCells->Release();
		}
		if (NULL != pdispWorkbooks)
		{
			Utils_GetMemid(pdispWorkbooks, L"Close", &dispid);
			Utils_InvokeMethod(pdispWorkbooks, dispid, NULL, 0, NULL);
			pdispWorkbooks->Release();
		}
		pdispExcel->Release();
	}
	return this->m_Data.size() > 0;
}

BOOL CImpICalibrationStandard::checkExcelFile(
	LPCTSTR			szFileName)
{
	return 0 == lstrcmpi(PathFindExtension(szFileName), L".xls");
}

HRESULT CImpICalibrationStandard::GetexpectedDistance(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_expectedDistance, pVarResult);
	return S_OK;
}

HRESULT CImpICalibrationStandard::SetexpectedDistance(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_expectedDistance = varg.dblVal;
	return S_OK;
}

HRESULT CImpICalibrationStandard::GetsourceDistance(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_sourceDistance, pVarResult);
	return S_OK;
}

HRESULT CImpICalibrationStandard::SetsourceDistance(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_sourceDistance = varg.dblVal;
	return S_OK;
}


CImpICalibrationStandard::CDataPoint::CDataPoint() :
	m_X(0.0),
	m_Y(0.0)
{

}
CImpICalibrationStandard::CDataPoint::CDataPoint(const CDataPoint& src) :
	m_X(src.m_X),
	m_Y(src.m_Y)
{

}
CImpICalibrationStandard::CDataPoint::CDataPoint(double x, double y) :
	m_X(x),
	m_Y(y)
{

}
CImpICalibrationStandard::CDataPoint::~CDataPoint()
{

}
const CImpICalibrationStandard::CDataPoint& CImpICalibrationStandard::CDataPoint::operator=(const CDataPoint& src)
{
	this->m_X = src.m_X;
	this->m_Y = src.m_Y;
	return *this;
}
