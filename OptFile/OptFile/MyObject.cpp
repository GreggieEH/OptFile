#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "MyOPTFile.h"
#include "DlgDisplayRadiance.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	m_pImp_clsIDataSet(NULL),
	m_pImpIPersistStreamInit(NULL),
	m_pImpIPersistFile(NULL),
	m_pImpISummaryInfo(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
	// _clsIDataSet interface ID
	m_iid_clsIDataSet(IID_NULL),
	// opt file
	m_pMyOPTFile(NULL),
	// __clsIDataSet event interface
	m_iid__clsIDataSet(IID_NULL),
	m_dispidrequestDispersion(DISPID_UNKNOWN),
	m_dispidrequestOpticalTransfer(DISPID_UNKNOWN),
	m_dispidrequestCalibrationScan(DISPID_UNKNOWN),
	m_dispidrequestCalibrationGain(DISPID_UNKNOWN),
	// summary info type info
	m_iidISummayInfo(IID_NULL)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	Utils_DELETE_POINTER(this->m_pImp_clsIDataSet);
	Utils_DELETE_POINTER(this->m_pImpIPersistStreamInit);
	Utils_DELETE_POINTER(this->m_pImpIPersistFile);
	Utils_DELETE_POINTER(this->m_pImpISummaryInfo);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_DELETE_POINTER(this->m_pMyOPTFile);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	else if (riid == this->m_iid_clsIDataSet)
		*ppv = this->m_pImp_clsIDataSet;
	else if (IID_IPersist == riid || IID_IPersistStreamInit == riid)
		*ppv = this->m_pImpIPersistStreamInit;
	else if (IID_IPersistFile == riid)
		*ppv = this->m_pImpIPersistFile;
	else if (riid == this->m_iidISummayInfo)
		*ppv = this->m_pImpISummaryInfo;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pImp_clsIDataSet				= new CImp_clsIDataSet(this, this->m_pUnkOuter);
	this->m_pImpIPersistStreamInit			= new CImpIPersistStreamInit(this, this->m_pUnkOuter);
	this->m_pImpIPersistFile				= new CImpIPersistFile(this, this->m_pUnkOuter);
	this->m_pImpISummaryInfo				= new CImpISummaryInfo(this, this->m_pUnkOuter);
	this->m_pMyOPTFile						= new CMyOPTFile(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pImp_clsIDataSet				&&
		NULL != this->m_pImpIPersistStreamInit			&&
		NULL != this->m_pImpIPersistFile				&&
		NULL != this->m_pImpISummaryInfo				&&
		NULL != this->m_pMyOPTFile)
	{
		hr = this->Init__clsIDataSet();
		if (SUCCEEDED(hr))
		{
			hr = Utils_CreateConnectionPoint(this, MY_IIDSINK,
				 &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
		}
		if (SUCCEEDED(hr))
		{
			hr = Utils_CreateConnectionPoint(this, this->m_iid__clsIDataSet,
				&(this->m_paConnPts[CONN_PT__clsIDataSet]));
		}
		if (SUCCEEDED(hr))
		{
			hr = this->m_pImp_clsIDataSet->MyInit();
		}
		if (SUCCEEDED(hr))
		{
			hr = this->m_pImpISummaryInfo->MyInit();
		}
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

HRESULT CMyObject::Init__clsIDataSet()
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->GetRefTypeInfo(TEXT("__clsIDataSet"), &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_iid__clsIDataSet	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, TEXT("requestDispersion"), &(this->m_dispidrequestDispersion));
		Utils_GetMemid(pTypeInfo, TEXT("requestOpticalTransfer"), &(this->m_dispidrequestOpticalTransfer));
		Utils_GetMemid(pTypeInfo, L"requestCalibrationScan", &this->m_dispidrequestCalibrationScan);
		Utils_GetMemid(pTypeInfo, L"requestCalibrationGain", &m_dispidrequestCalibrationGain);
	}
	return hr;
}

// request slit width for named slit
BOOL CMyObject::FireRequestSlitWidth(
									LPCTSTR			szSlitName,
									double		*	slitWidth)
{
	VARIANTARG			avarg[3];
	VARIANT_BOOL		fHandled		= VARIANT_FALSE;
	*slitWidth		= 1.0;
	InitVariantFromString(szSlitName, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_R8;
	avarg[1].pdblVal	= slitWidth;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSlitWidth, avarg, 3);
	VariantClear(&avarg[2]);
	return VARIANT_TRUE == fHandled;
}

double CMyObject::FirerequestDispersion(
									long			grating)
{
	VARIANTARG			avarg[2];
	double				dispersion		= 2.0;		// default value
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I4;
	avarg[1].plVal		= &grating;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= &dispersion;
	Utils_NotifySinks(this, this->m_iid__clsIDataSet, this->m_dispidrequestDispersion, avarg, 2);
	return dispersion;
}

BOOL CMyObject::FirerequestOpticalTransfer(
									long		*	nValues,
									double		**	ppWaves,
									double		**	ppSignals)
{
	VARIANTARG			avarg[2];
	SAFEARRAY		*	pSAX		= NULL;
	SAFEARRAY		*	pSAY		= NULL;
	long				ubx;
	long				uby;
	double			*	arrX;
	double			*	arrY;
	long				i;
	BOOL				fSuccess	= FALSE;
	*nValues			= 0;
	*ppWaves			= NULL;
	*ppSignals			= NULL;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_ARRAY | VT_BYREF | VT_R8;
	avarg[1].pparray	= &pSAX;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_ARRAY | VT_BYREF | VT_R8;
	avarg[0].pparray	= &pSAY;
	Utils_NotifySinks(this, this->m_iid__clsIDataSet, this->m_dispidrequestOpticalTransfer, avarg, 2);
	// have both arrays
	if (NULL != pSAX && NULL != pSAY)
	{
		// compare array sizes
		SafeArrayGetUBound(pSAX, 1, &ubx);
		SafeArrayGetUBound(pSAY, 1, &uby);
		if (ubx == uby)
		{
			fSuccess = TRUE;
			// allocate space
			*nValues	= ubx+1;
			*ppWaves	= new double [*nValues];
			*ppSignals	= new double [*nValues];
			// copy the data
			SafeArrayAccessData(pSAX, (void**) &arrX);
			SafeArrayAccessData(pSAY, (void**) &arrY);
			for (i=0; i<(*nValues); i++)
			{
				(*ppWaves)[i]	= arrX[i];
				(*ppSignals)[i]	= arrY[i];
			}
			SafeArrayUnaccessData(pSAX);
			SafeArrayUnaccessData(pSAY);
		}
	}
	SafeArrayDestroy(pSAX);
	SafeArrayDestroy(pSAY);
	return fSuccess;
}

BOOL CMyObject::FirerequestCalibrationScan(
	short int		Grating,
	LPCTSTR			filter,
	short int		detector,
	long		*	nValues,
	double		**	ppWaves,
	double		**	ppSignals)
{
	VARIANTARG			avarg[5];
	SAFEARRAY		*	pSAX = NULL;
	SAFEARRAY		*	pSAY = NULL;
	long				ubx;
	long				uby;
	double			*	arrX;
	double			*	arrY;
	long				i;
	BOOL				fSuccess = FALSE;
	*nValues = 0;
	*ppWaves = NULL;
	*ppSignals = NULL;
	InitVariantFromInt16(Grating, &avarg[4]);
	InitVariantFromString(filter, &avarg[3]);
	InitVariantFromInt16(detector, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_ARRAY | VT_BYREF | VT_R8;
	avarg[1].pparray = &pSAX;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_ARRAY | VT_BYREF | VT_R8;
	avarg[0].pparray = &pSAY;
	Utils_NotifySinks(this, this->m_iid__clsIDataSet, 
		this->m_dispidrequestCalibrationScan, avarg, 5);
	VariantClear(&avarg[3]);
	// have both arrays
	if (NULL != pSAX && NULL != pSAY)
	{
		// compare array sizes
		SafeArrayGetUBound(pSAX, 1, &ubx);
		SafeArrayGetUBound(pSAY, 1, &uby);
		if (ubx == uby)
		{
			fSuccess = TRUE;
			// allocate space
			*nValues = ubx + 1;
			*ppWaves = new double[*nValues];
			*ppSignals = new double[*nValues];
			// copy the data
			SafeArrayAccessData(pSAX, (void**)&arrX);
			SafeArrayAccessData(pSAY, (void**)&arrY);
			for (i = 0; i<(*nValues); i++)
			{
				(*ppWaves)[i] = arrX[i];
				(*ppSignals)[i] = arrY[i];
			}
			SafeArrayUnaccessData(pSAX);
			SafeArrayUnaccessData(pSAY);
		}
	}
	SafeArrayDestroy(pSAX);
	SafeArrayDestroy(pSAY);
	return fSuccess;
}

double CMyObject::FirerequestCalibrationGain(
	double			wavelength)
{
	VARIANTARG		avarg[2];
	double			calibrationGainFactor = 1.0;
	InitVariantFromDouble(wavelength, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = &calibrationGainFactor;
	Utils_NotifySinks(this, this->m_iid__clsIDataSet,
		this->m_dispidrequestCalibrationGain, avarg, 2);
	return calibrationGainFactor;
}

BOOL CMyObject::FireRequestGetInputPort(
									LPTSTR			szInputPort,
									UINT			nBufferSize)
{
	VARIANTARG		avarg[2];
	BSTR			bstr		= NULL;
	VARIANT_BOOL	fHandled	= VARIANT_FALSE;
	LPTSTR			szTemp		= NULL;
	BOOL			fSuccess	= FALSE;
	szInputPort[0]		= '\0';
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_BSTR;
	avarg[1].pbstrVal	= &bstr;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestGetInputPort, avarg, 2);
	if (VARIANT_TRUE == fHandled)
	{
		if (NULL != bstr)
		{
	//		SHStrDup(bstr, &szTemp);
			SHStrDup(bstr, &szTemp);
			StringCchCopy(szInputPort, nBufferSize, szTemp);
			fSuccess = TRUE;
			CoTaskMemFree((LPVOID) szTemp);
		}
	}
	SysFreeString(bstr);
	return fSuccess;
}

BOOL CMyObject::FireRequestSetInputPort(
									LPCTSTR			szInputPort)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	fHandled	= VARIANT_FALSE;
	InitVariantFromString(szInputPort, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetInputPort, avarg, 2);
	VariantClear(&avarg[1]);
	return VARIANT_TRUE == fHandled;
}

BOOL CMyObject::FireRequestGetOutputPort(
									LPTSTR			szOutputPort,
									UINT			nBufferSize)
{
	VARIANTARG		avarg[2];
	BSTR			bstr		= NULL;
	VARIANT_BOOL	fHandled	= VARIANT_FALSE;
	LPTSTR			szTemp		= NULL;
	BOOL			fSuccess	= FALSE;
	szOutputPort[0]		= '\0';
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_BSTR;
	avarg[1].pbstrVal	= &bstr;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestGetOutputPort, avarg, 2);
	if (VARIANT_TRUE == fHandled)
	{
		if (NULL != bstr)
		{
//		SHStrDup(bstr, &szTemp);
			StringCchCopy(szOutputPort, nBufferSize, bstr);
			fSuccess = TRUE;
	//		CoTaskMemFree((LPVOID) szTemp);
		}
	}
	SysFreeString(bstr);
	return fSuccess;
}

BOOL CMyObject::FireRequestSetOutputPort(
									LPCTSTR			szOutputPort)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	fHandled	= VARIANT_FALSE;
	InitVariantFromString(szOutputPort, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetOutputPort, avarg, 2);
	VariantClear(&avarg[1]);
	return VARIANT_TRUE == fHandled;
}

HRESULT CMyObject::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

HRESULT CMyObject::GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo;
	BOOL			fSuccess	= FALSE;
	*ppTypeInfo	= NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
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
	case DISPID_DetectorInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDetectorInfo(pVarResult);
		break;
	case DISPID_InputInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetInputInfo(pVarResult);
		break;
	case DISPID_SlitInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSlitInfo(pVarResult);
		break;
	case DISPID_NumGratingScans:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNumGratingScans(pVarResult);
		break;
	case DISPID_RadianceAvailable:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRadianceAvailable(pVarResult);
		break;
	case DISPID_CalibrationStandard:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCalibrationStandard(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCalibrationStandard(pDispParams);
		break;
	case DISPID_CalibrationMeasurement:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCalibrationMeasurement(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCalibrationMeasurement(pDispParams);
		break;
	case DISPID_Clone:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->Clone(pVarResult);
		break;
	case DISPID_Setup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->Setup(pDispParams, pVarResult);
		break;
	case DISPID_GetGratingScan:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetGratingScan(pDispParams, pVarResult);
		break;
	case DISPID_InitNew:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->InitNew(pVarResult);
		break;
	case DISPID_ReadConfig:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadConfig(pDispParams, pVarResult);
		break;
	case DISPID_ReadINI:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadINI(pDispParams);
		break;
	case DISPID_SetExp:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetExp(pDispParams);
		break;
	case DISPID_AddValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->AddValue(pDispParams);
		break;
	case DISPID_SaveToFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return SaveToFile(pDispParams);
		break;
	case DISPID_LoadFromFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->LoadFromFile(pDispParams, pVarResult);
		break;
	case DISPID_ClearScanData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ClearScanData();
		break;
	case DISPID_SetSlitInfo:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetSlitInfo(pDispParams);
		break;
	case DISPID_SelectCalibrationStandard:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SelectCalibrationStandard(pDispParams, pVarResult);
		break;
	case DISPID_ClearCalibrationStandard:
		return this->ClearCalibrationStandard();
	case DISPID_OpticalTransferFunction:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetOpticalTransferFunction(pVarResult);
		break;
	case DISPID_LoadOpticalTransferFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->LoadOpticalTransferFile(pDispParams, pVarResult);
		break;
	case DISPID_DisplayRadiance:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayRadiance(pDispParams, pVarResult);
		break;
	case DISPID_TimeStamp:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetTimeStamp(pVarResult);
		break;
	case DISPID_InitAcquisition:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->InitAcquisition(pDispParams, pVarResult);
		break;
	case DISPID_ConfigFile:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetConfigFile(pVarResult);
		break;
	case DISPID_ClearReadonly:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ClearReadonly();
		break;
	case DISPID_SaveToString:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SaveToString(pVarResult);
		break;
	case DISPID_LoadFromString:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->LoadFromString(pDispParams, pVarResult);
		break;
	case DISPID_FormOpticalTransferFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->FormOpticalTransferFile(pDispParams, pVarResult);
		break;
	case DISPID_ADInfoString:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetADInfoString(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetADInfoString(pDispParams);
		break;
	case DISPID_Comment:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetComment(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetComment(pDispParams);
		break;
	case DISPID_AllowNegativeValues:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAllowNegativeValues(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAllowNegativeValues(pDispParams);
		break;
	case DISPID_GetExtraValues:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetExtraValues(pDispParams, pVarResult);
		break;
	case DISPID_SetTimeStamp:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetTimeStamp();
		break;
	case DISPID_ExtraValueTitles:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetExtraValueTitles(pVarResult);
		break;
	case DISPID_SignalUnits:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSignalUnits(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSignalUnits(pDispParams);
		break;
	case DISPID_GetADGainFactor:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetADGainFactor(pDispParams, pVarResult);
		break;
	case DISPID_RecallSettings:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->RecallSettings(pDispParams, pVarResult);
		break;
	case DISPID_SaveSettings:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SaveSettings(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpIDispatch::GetDetectorInfo(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_DISPATCH;
	this->m_pBackObj->m_pMyOPTFile->GetDetectorInfo(&(pVarResult->pdispVal));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetInputInfo(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_DISPATCH;
	this->m_pBackObj->m_pMyOPTFile->GetInputInfo(&(pVarResult->pdispVal));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSlitInfo(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_DISPATCH;
	this->m_pBackObj->m_pMyOPTFile->GetSlitInfo(&(pVarResult->pdispVal));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetNumGratingScans(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyOPTFile->GetNumGratingScans(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetRadianceAvailable(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetRadianceAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetCalibrationStandard(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (VT_DISPATCH == pVarResult->vt)
	{
		this->m_pBackObj->m_pMyOPTFile->GetCalibrationStandard(&(pVarResult->pdispVal));
	}
	else
	{
		// get the calibration standard file name
		TCHAR		szCalibrationStandard[MAX_PATH];
		if (this->m_pBackObj->m_pMyOPTFile->GetCalibrationStandard(szCalibrationStandard, MAX_PATH))
		{
			InitVariantFromString(szCalibrationStandard, pVarResult);
		}
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCalibrationStandard(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szString);
	VariantClear(&varg);
	this->m_pBackObj->m_pMyOPTFile->SetCalibrationStandard(szString);
	CoTaskMemFree((LPVOID) szString);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetCalibrationMeasurement(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetCalibrationMeasurement(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCalibrationMeasurement(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetCalibrationMeasurement(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::Clone(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_DISPATCH;
	this->m_pBackObj->m_pMyOPTFile->Clone(&(pVarResult->pdispVal));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::Setup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	HWND				hwnd;
	LPTSTR				szString		= NULL;
	TCHAR				szPart[MAX_PATH];
	BOOL				fSuccess		= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hwnd = (HWND) varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szString);
	StringCchCopy(szPart, MAX_PATH, szString);
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varg);
	fSuccess = this->m_pBackObj->m_pMyOPTFile->Setup(hwnd, szPart);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetGratingScan(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	VariantInit(pVarResult);
	pVarResult->vt		= VT_DISPATCH;
	this->m_pBackObj->m_pMyOPTFile->GetGratingScan(varg.lVal, &(pVarResult->pdispVal));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::InitNew(
									VARIANT		*	pVarResult)
{
	BOOL			fSuccess	= this->m_pBackObj->m_pMyOPTFile->InitNew();
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szConfigFile	= NULL;
	BOOL				fSuccess		= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szConfigFile);
	VariantClear(&varg);
	fSuccess	= this->m_pBackObj->m_pMyOPTFile->ReadConfig(szConfigFile);
	CoTaskMemFree((LPVOID) szConfigFile);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadINI(
	DISPPARAMS	*	 pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szFolder[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szFolder, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	PathRemoveFileSpec(szFolder);
	InitVariantFromString(szFolder, &varg);
	Utils_InvokeMethod(this, DISPID_RecallSettings, &varg, 1, NULL);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::WriteINI(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szFolder[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szFolder, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	PathRemoveFileSpec(szFolder);
	InitVariantFromString(szFolder, &varg);
	Utils_InvokeMethod(this, DISPID_SaveSettings, &varg, 1, NULL);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetExp(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;
	TCHAR				szFilter[MAX_PATH];
	long				grating;
	long				detector = 1;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szString);
	StringCchCopy(szFilter, MAX_PATH, szString);
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.lVal;
	hr = DispGetParam(pDispParams, 2, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	detector = varg.lVal;
	this->m_pBackObj->m_pMyOPTFile->SetCurrentExp(szFilter, grating, detector);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::AddValue(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	double				NM;
	double				Signal;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	NM = varg.dblVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Signal = varg.dblVal;
	this->m_pBackObj->m_pMyOPTFile->AddValue(NM, Signal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SaveToFile(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szFileName		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szFileName);
	VariantClear(&varg);
	this->m_pBackObj->m_pMyOPTFile->SaveToFile(szFileName, TRUE);
	CoTaskMemFree((LPVOID) szFileName);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::LoadFromFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	LPTSTR			szFileName		= NULL;
	BOOL			fSuccess		= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szFileName);
	VariantClear(&varg);
	fSuccess	= this->m_pBackObj->m_pMyOPTFile->LoadFromFile(szFileName);
	CoTaskMemFree((LPVOID) szFileName);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ClearScanData()
{
	this->m_pBackObj->m_pMyOPTFile->ClearScanData();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSlitInfo(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;
	TCHAR				szSlitTitle[MAX_PATH];
	TCHAR				szSlitName[MAX_PATH];
	double				slitWidth;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szString);
	StringCchCopy(szSlitTitle, MAX_PATH, szString);
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szString);
	StringCchCopy(szSlitName, MAX_PATH, szString);
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 2, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	slitWidth = varg.dblVal;
	this->m_pBackObj->m_pMyOPTFile->SetSlitInfo(szSlitTitle, szSlitName, slitWidth);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::SelectCalibrationStandard(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess		= FALSE;
	IDispatch		*	pdisp;
	DISPID				dispid;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->Setup((HWND) varg.lVal, TEXT("Calibration Standard"));
	if (this->m_pBackObj->m_pMyOPTFile->GetCalibrationStandard(&pdisp))
	{
		Utils_GetMemid(pdisp, TEXT("amLoaded"), &dispid);
		fSuccess = Utils_GetBoolProperty(pdisp, dispid);
		pdisp->Release();
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::ClearCalibrationStandard()
{
	this->m_pBackObj->m_pMyOPTFile->ClearCalibrationStandard();
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::GetOpticalTransferFunction(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetOpticalTransferFunction(), pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::LoadOpticalTransferFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szFileName	= NULL;
	BOOL				fSuccess	= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szFileName);
	VariantClear(&varg);
	fSuccess = this->m_pBackObj->m_pMyOPTFile->LoadOpticalTransferFile(szFileName);
	CoTaskMemFree((LPVOID) szFileName);
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::DisplayRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	HWND				hwndParent;
	BOOL				fRadiance;
	BOOL				fSuccess		= FALSE;
	long				nValues;
	double			*	pX;
	double			*	pY;
	CDlgDisplayRadiance	dlg;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hwndParent	= (HWND) varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fRadiance	= VARIANT_TRUE == varg.boolVal;
	if (fRadiance) //this is the radiance calculation
		fSuccess = this->m_pBackObj->m_pMyOPTFile->CalculateRadiance(&nValues, &pX, &pY);
	else // this is irradiance calc
		fSuccess = this->m_pBackObj->m_pMyOPTFile->CalculateIrradiance(&nValues, &pX, &pY);
	if (fSuccess)
	{
		dlg.SetArrayData(nValues, pX, pY);
		dlg.DoOpenDialog(hwndParent, GetServer()->GetInstance());
		delete [] pX;
		delete [] pY;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::GetTimeStamp(
									VARIANT		*	pVarResult)
{
	TCHAR				szTimeStamp[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pMyOPTFile->GetTimeStamp(szTimeStamp, MAX_PATH))
	{
		InitVariantFromString(szTimeStamp, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::InitAcquisition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	HWND			hwndParent;
	BOOL			fCalibration;
	BOOL			fSuccess	= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hwndParent	= (HWND) varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fCalibration = VARIANT_TRUE == varg.boolVal;
	fSuccess = this->m_pBackObj->m_pMyOPTFile->InitAcquisition(hwndParent, fCalibration);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetConfigFile(
									VARIANT		*	pVarResult)
{
	TCHAR			szConfig[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pMyOPTFile->GetConfigFile(szConfig, MAX_PATH))
		InitVariantFromString(szConfig, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ClearReadonly()
{
	this->m_pBackObj->m_pMyOPTFile->ClearReadonly();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetADChannel(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromInt32(this->m_pBackObj->m_pMyOPTFile->GetADChannel(varg.dblVal), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SaveToString(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	LPTSTR			szString = NULL;
	if (this->m_pBackObj->m_pMyOPTFile->SaveToString(&szString, TRUE))
	{
		InitVariantFromString(szString, pVarResult);
		CoTaskMemFree((LPVOID)szString);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::LoadFromString(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pMyOPTFile->LoadFromString(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::FormOpticalTransferFile(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pMyOPTFile->FormOpticalTransfer(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::GetADInfoString(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	LPTSTR			infoString = NULL;
	if (this->m_pBackObj->m_pMyOPTFile->GetADInfoString(&infoString))
	{
		InitVariantFromString(infoString, pVarResult);
		CoTaskMemFree((LPVOID)infoString);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetADInfoString(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetADInfoString(varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_ADInfoString);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetComment(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szComment[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetComment(szComment, MAX_PATH))
	{
		InitVariantFromString(szComment, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetComment(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetComment(varg.bstrVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAllowNegativeValues(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetAllowNegativeValues(),
		pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAllowNegativeValues(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetAllowNegativeValues(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetExtraValues(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	TCHAR			szTitle[MAX_PATH];
	long			nValues;
	double		*	pExtraValues = NULL;
	long			i;
	SAFEARRAYBOUND	sab;
	double		*	arr;
	BOOL			fSuccess = FALSE;
	VARIANTARG		varg;
	UINT			uArgErr;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt)
	{
		return DISP_E_TYPEMISMATCH;
	}
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szTitle, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	if (this->m_pBackObj->m_pMyOPTFile->GetExtraValues(szTitle, &nValues, &pExtraValues))
	{
		fSuccess = TRUE;
//		*(pDispParams->rgvarg[1].pbstrVal) = SysAllocString(szTitle);
		sab.lLbound = 0;
		sab.cElements = nValues;
		*(pDispParams->rgvarg[0].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arr);
		for (i = 0; i < nValues; i++)
			arr[i] = pExtraValues[i];
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		delete[] pExtraValues;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetTimeStamp()
{
	this->m_pBackObj->m_pMyOPTFile->SetTimeStamp();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetExtraValueTitles(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	ULONG			nTitles = 0;
	LPTSTR		*	pszTitles = NULL;
	ULONG			i;
	if (this->m_pBackObj->m_pMyOPTFile->GetExtraValueTitles(&pszTitles, &nTitles))
	{
		InitVariantFromStringArray((PCWSTR*) pszTitles, nTitles, pVarResult);
		// cleanup
		for (i = 0; i < nTitles; i++)
		{
			CoTaskMemFree((LPVOID)pszTitles[i]);
			pszTitles[i] = NULL;
		}
		CoTaskMemFree((LPVOID)pszTitles);
	}
	return S_OK;
}
HRESULT CMyObject::CImpIDispatch::GetSignalUnits(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szSignalUnits[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetSignalUnits(szSignalUnits, MAX_PATH))
	{
		InitVariantFromString(szSignalUnits, pVarResult);
	}
	return S_OK;
}
HRESULT CMyObject::CImpIDispatch::SetSignalUnits(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetSignalUnits(varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_SignalUnits);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetADGainFactor(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	// input wavelength
	VARIANTARG		varg;
	UINT			uArgErr;
	double			wavelength;
	HRESULT			hr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	wavelength = varg.dblVal;

	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(
		this->m_pBackObj->m_pMyOPTFile->DetermineDataGainFactor(wavelength),
		pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::RecallSettings(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess = FALSE;
	TCHAR				szFilePath[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// form recall file name
	StringCchCopy(szFilePath, MAX_PATH, varg.bstrVal);
	PathAppend(szFilePath, L"OptFile.xml");
	if (PathFileExists(szFilePath))
	{
		fSuccess = this->m_pBackObj->m_pMyOPTFile->LoadFromFile(szFilePath);
	}
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SaveSettings(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess = FALSE;
	TCHAR				szFilePath[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// form save file name
	StringCchCopy(szFilePath, MAX_PATH, varg.bstrVal);
	PathAppend(szFilePath, L"OptFile.xml");
	fSuccess = this->m_pBackObj->m_pMyOPTFile->SaveToFile(szFilePath, TRUE);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt	= this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (riid == this->m_pBackObj->m_iid__clsIDataSet)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT__clsIDataSet];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}

CMyObject::CImp_clsIDataSet::CImp_clsIDataSet(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_AnalysisMode(0),
	m_dispidReadDataFile(DISPID_UNKNOWN),
	m_dispidWriteDataFile(DISPID_UNKNOWN),
	m_dispidGetWavelengths(DISPID_UNKNOWN),
	m_dispidGetSpectra(DISPID_UNKNOWN),
	m_dispidAddValue(DISPID_UNKNOWN),
	m_dispidViewSetup(DISPID_UNKNOWN),
	m_dispidGetArraySize(DISPID_UNKNOWN),
	m_dispidSetCurrentExp(DISPID_UNKNOWN),
	m_dispidverifyCalibration(DISPID_UNKNOWN),
	m_dispidRadianceAvailable(DISPID_UNKNOWN),
	m_dispidIrradianceAvailable(DISPID_UNKNOWN),
	m_dispidcalcRadiance(DISPID_UNKNOWN),
	m_dispidAnalysisMode(DISPID_UNKNOWN),
	m_dispidGetCreateTime(DISPID_UNKNOWN),
	m_dispidfileName(DISPID_UNKNOWN),
	m_dispidamCalibration(DISPID_UNKNOWN),
	m_dispidGetDataValue(DISPID_UNKNOWN),
	m_dispidSetDataValue(DISPID_UNKNOWN),
	m_dispidComment(DISPID_UNKNOWN),
	m_dispidExtraScanData(DISPID_UNKNOWN),
	m_dispidDefaultUnits(DISPID_UNKNOWN),
	m_dispidDefaultTitle(DISPID_UNKNOWN),
	m_dispidCountsAvailable(DISPID_UNKNOWN),
	m_dispidNumberExtraScanArrays(DISPID_UNKNOWN),
	m_dispidGetExtraScanArray(DISPID_UNKNOWN),
	m_dispidGetGratingScan(DISPID_UNKNOWN),
	m_dispidExportToCSV(DISPID_UNKNOWN),
	m_dispidGetADGainFactor(DISPID_UNKNOWN)
{
}

CMyObject::CImp_clsIDataSet::~CImp_clsIDataSet()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImp_clsIDataSet::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIDataSet::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIDataSet::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImp_clsIDataSet::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImp_clsIDataSet::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	if (NULL != this->m_pTypeInfo)
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
		return S_OK;
	}
	else
	{
		*ppTInfo	= NULL;
		return E_FAIL;
	}
}

STDMETHODIMP CMyObject::CImp_clsIDataSet::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImp_clsIDataSet::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	if (dispIdMember == m_dispidReadDataFile)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadDataFile(pVarResult);
	}
	else if (dispIdMember == m_dispidWriteDataFile)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteDataFile(pVarResult);
	}
	else if (dispIdMember == m_dispidGetWavelengths)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetWavelengths(pVarResult);
	}
	else if (dispIdMember == m_dispidGetSpectra)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetSpectra(pVarResult);
	}
	else if (dispIdMember == m_dispidAddValue)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->AddValue(pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidViewSetup)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ViewSetup(pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidGetArraySize)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetArraySize(pVarResult);
	}
	else if (dispIdMember == m_dispidSetCurrentExp)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetCurrentExp(pDispParams);
	}
	else if (dispIdMember == m_dispidverifyCalibration)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->verifyCalibration(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidRadianceAvailable)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRadianceAvailable(pVarResult);
		else
			return S_OK;
	}
	else if (dispIdMember == this->m_dispidIrradianceAvailable)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetIrradianceAvailable(pVarResult);
		else
			return S_OK;
	}
	else if (dispIdMember == this->m_dispidcalcRadiance)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->calcRadiance(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidAnalysisMode)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAnalysisMode(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAnalysisMode(pDispParams);
	}
	else if (dispIdMember == this->m_dispidGetCreateTime)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetCreateTime(pVarResult);
	}
	else if (dispIdMember == this->m_dispidfileName)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetfileName(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetfileName(pDispParams);
	}
	else if (dispIdMember == this->m_dispidamCalibration)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetamCalibration(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetamCalibration(pDispParams);
	}
	else if (dispIdMember == this->m_dispidGetDataValue)
	{
		return this->GetDataValue(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidSetDataValue)
	{
		return this->SetDataValue(pDispParams);
	}
	else if (dispIdMember == this->m_dispidComment)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetComment(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetComment(pDispParams);
	}
	else if (dispIdMember == this->m_dispidExtraScanData)
	{
		return this->ExtraScanData(pDispParams);
	}
	else if (dispIdMember == this->m_dispidDefaultUnits)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDefaultUnits(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDefaultUnits(pDispParams);
	}
	else if (dispIdMember == this->m_dispidDefaultTitle)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetDefaultTitle(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetDefaultTitle(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidCountsAvailable)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCountsAvailable(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCountsAvailable(pDispParams);
	}
	else if (dispIdMember == this->m_dispidNumberExtraScanArrays)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNumberExtraScanArrays(pVarResult);
	}
	else if (dispIdMember == this->m_dispidGetExtraScanArray)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetExtraScanArray(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidGetGratingScan)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetGratingScan(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidExportToCSV)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ExportToCSV(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidGetADGainFactor)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetADGainFactor(pDispParams, pVarResult);
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImp_clsIDataSet::MyInit()
{
	HRESULT				hr;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pBackObj->GetRefTypeInfo(TEXT("_clsIDataSet"), &(this->m_pTypeInfo));
	if (SUCCEEDED(hr))
	{
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pBackObj->m_iid_clsIDataSet	= pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadDataFile"), &m_dispidReadDataFile);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("WriteDataFile"), &m_dispidWriteDataFile);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetWavelengths"), &m_dispidGetWavelengths);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetSpectra"), &m_dispidGetSpectra);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AddValue"), &m_dispidAddValue);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ViewSetup"), &m_dispidViewSetup);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetArraySize"), &m_dispidGetArraySize);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SetCurrentExp"), &m_dispidSetCurrentExp);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("RadianceAvailable"), &m_dispidRadianceAvailable);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("IrradianceAvailable"), &m_dispidIrradianceAvailable);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("calcRadiance"), &m_dispidcalcRadiance);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AnalysisMode"), &m_dispidAnalysisMode);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetCreateTime"), &m_dispidGetCreateTime);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("fileName"), &m_dispidfileName);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("amCalibration"), &m_dispidamCalibration);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetDataValue"), &m_dispidGetDataValue);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SetDataValue"), &m_dispidSetDataValue);
		Utils_GetMemid(this->m_pTypeInfo, L"Comment", &m_dispidComment);
		Utils_GetMemid(this->m_pTypeInfo, L"ExtraScanData", &m_dispidExtraScanData);
		Utils_GetMemid(this->m_pTypeInfo, L"DefaultUnits", &m_dispidDefaultUnits);
		Utils_GetMemid(this->m_pTypeInfo, L"DefaultTitle", &m_dispidDefaultTitle);
		Utils_GetMemid(this->m_pTypeInfo, L"CountsAvailable", &m_dispidCountsAvailable);
		Utils_GetMemid(this->m_pTypeInfo, L"NumberExtraScanArrays", &m_dispidNumberExtraScanArrays);
		Utils_GetMemid(this->m_pTypeInfo, L"GetExtraScanArray", &m_dispidGetExtraScanArray);
		Utils_GetMemid(this->m_pTypeInfo, L"GetGratingScan", &this->m_dispidGetGratingScan);
		Utils_GetMemid(this->m_pTypeInfo, L"ExportToCSV", &m_dispidExportToCSV);
		Utils_GetMemid(this->m_pTypeInfo, L"GetADGainFactor", &m_dispidGetADGainFactor);
	}
	return hr;
}

HRESULT CMyObject::CImp_clsIDataSet::ReadDataFile(
									VARIANT		*	pVarResult)
{
	BOOL fSuccess = this->m_pBackObj->m_pMyOPTFile->LoadFromFile();
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::WriteDataFile(
									VARIANT		*	pVarResult)
{
	BOOL		fSuccess = this->m_pBackObj->m_pMyOPTFile->SaveToFile();
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetWavelengths(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long			nArray;
	double		*	pWaves;
	SAFEARRAYBOUND	sab;
	long			i;
	double		*	arr;

	if (this->m_pBackObj->m_pMyOPTFile->GetWavelengths(&nArray, &pWaves))
	{
		sab.lLbound			= 0;
		sab.cElements		= nArray;
		VariantInit(pVarResult);
		pVarResult->vt		= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**) &arr);
		for (i=0; i<nArray; i++) arr[i]	= pWaves[i];
		SafeArrayUnaccessData(pVarResult->parray);
		delete [] pWaves;
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetSpectra(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long			nArray;
	double		*	pSignals;
	SAFEARRAYBOUND	sab;
	long			i;
	double		*	arr;
	if (this->m_pBackObj->m_pMyOPTFile->GetSpectra(&nArray, &pSignals))
	{
		sab.lLbound			= 0;
		sab.cElements		= nArray;
		VariantInit(pVarResult);
		pVarResult->vt		= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**) &arr);
		for (i=0; i<nArray; i++) arr[i]	= pSignals[i];
		SafeArrayUnaccessData(pVarResult->parray);
		delete [] pSignals;
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::AddValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANT			varCopy;
//	UINT			uArgErr;
	double			NM;
	double			Signal;
	BOOL			fSuccess	= FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	NM = varCopy.dblVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	Signal	= varCopy.dblVal;
	fSuccess	= this->m_pBackObj->m_pMyOPTFile->AddValue(NM, Signal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::ViewSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess	= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess	= this->m_pBackObj->m_pMyOPTFile->ViewSetup((HWND) varg.lVal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetArraySize(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyOPTFile->GetArraySize(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetCurrentExp(
									DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANT			varCopy;
//	UINT			uArgErr;
	TCHAR			szFilter[MAX_PATH];
	LPTSTR			szString		= NULL;
	long			grating;
	long			detector;
	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[2]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_BSTR);
	if (FAILED(hr)) return hr;
	SHStrDup(varCopy.bstrVal, &szString);
	StringCchCopy(szFilter, MAX_PATH, szString);
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_I4);
	if (FAILED(hr)) return hr;
	grating = varCopy.lVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_I4);
	if (FAILED(hr)) return hr;
	detector = varCopy.lVal;
	this->m_pBackObj->m_pMyOPTFile->SetCurrentExp(szFilter, grating, detector);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::verifyCalibration(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetRadianceAvailable(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetRadianceAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetIrradianceAvailable(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetIrradianceAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::calcRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varCopy;
	BOOL				fRadianceCalc		= TRUE;
	long				nValues				= 0;
	double			*	pWaves				= NULL;
	double			*	pCalc				= NULL;
	SAFEARRAYBOUND		sab;
	double			*	arrX;
	double			*	arrY;
	long				i;
	BOOL				fSuccess			= FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_ARRAY | VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_BOOL);
	if (FAILED(hr)) return hr;
	if (VARIANT_TRUE == varCopy.boolVal)
	{
		fSuccess = this->m_pBackObj->m_pMyOPTFile->CalculateRadiance(&nValues, &pWaves, &pCalc);
	}
	else
	{
		fSuccess = this->m_pBackObj->m_pMyOPTFile->CalculateIrradiance(&nValues, &pWaves, &pCalc);
	}
	if (fSuccess)
	{
		sab.lLbound		= 0;
		sab.cElements	= nValues;
		*(pDispParams->rgvarg[0].pparray)	= SafeArrayCreate(VT_R8, 1, &sab);
		VariantInit(pVarResult);
		pVarResult->vt	= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**) &arrX);
		SafeArrayAccessData(pVarResult->parray, (void**) &arrY);
		for (i=0; i<nValues; i++)
		{
			arrX[i]	= pWaves[i];
			arrY[i] = pCalc[i];
		}
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		SafeArrayUnaccessData(pVarResult->parray);
		// cleanup
		delete [] pWaves;
		delete [] pCalc;
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetAnalysisMode(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt16(this->m_AnalysisMode, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetAnalysisMode(
									DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_AnalysisMode = varg.iVal;
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetCreateTime(
									VARIANT		*	pVarResult)
{
	TCHAR			szTimeStamp[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pMyOPTFile->GetTimeStamp(szTimeStamp, MAX_PATH))
	{
		InitVariantFromString(szTimeStamp, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetfileName(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szFilePath[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetFileName(szFilePath, MAX_PATH))
	{
		InitVariantFromString(szFilePath, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetfileName(
									DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	UINT			uArgErr;
	VARIANTARG		varg;
	LPTSTR			szFileName	= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szFileName);
	VariantClear(&varg);
	this->m_pBackObj->m_pMyOPTFile->SetFileName(szFileName);
	CoTaskMemFree((LPVOID) szFileName);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetamCalibration(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyOPTFile->GetCalibrationMeasurement(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetamCalibration(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetCalibrationMeasurement(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetDataValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANT			varCopy;

	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	InitVariantFromDouble(this->m_pBackObj->m_pMyOPTFile->GetDataValue(varCopy.dblVal), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetDataValue(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANT				varCopy;
	double				wavelength;
	double				signal;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	wavelength = varCopy.dblVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	signal = varCopy.dblVal;
	this->m_pBackObj->m_pMyOPTFile->SetDataValue(wavelength, signal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetComment(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szComment[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetComment(szComment, MAX_PATH))
	{
		InitVariantFromString(szComment, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetComment(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetComment(varg.bstrVal);
	return S_OK;

}

HRESULT CMyObject::CImp_clsIDataSet::ExtraScanData(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szTitle[MAX_PATH];
	double				wavelength;
	double				extraValue;
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
	this->m_pBackObj->m_pMyOPTFile->ExtraScanValue(szTitle, wavelength, extraValue);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetDefaultUnits(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
//	InitVariantFromString(L"V", pVarResult);
	TCHAR			szUnits[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetSignalUnits(szUnits, MAX_PATH))
	{
		InitVariantFromString(szUnits, pVarResult);
	}
	return S_OK;
}
HRESULT CMyObject::CImp_clsIDataSet::SetDefaultUnits(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetSignalUnits(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetDefaultTitle(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szTitle[MAX_PATH];
	this->m_pBackObj->m_pMyOPTFile->GetDefaultTitle(szTitle, MAX_PATH);
	InitVariantFromString(szTitle, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::SetDefaultTitle(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyOPTFile->SetDefaultTitle(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetCountsAvailable(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(FALSE, pVarResult);
	return S_OK;
}
HRESULT CMyObject::CImp_clsIDataSet::SetCountsAvailable(
	DISPPARAMS	*	pDispParams)
{
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetNumberExtraScanArrays(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyOPTFile->GetNumberExtraScanArrays(), pVarResult);
	return S_OK;
}
HRESULT CMyObject::CImp_clsIDataSet::GetExtraScanArray(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				nArray;
	double			*	pWaves;
	double			*	pValues;
	TCHAR				szTitle[MAX_PATH];
	SAFEARRAYBOUND		sab;
	long				i;
	double			*	arrWaves;
	double			*	arrValues;
	szTitle[0] = '\0';
	if (NULL == pVarResult) return E_INVALIDARG;
	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[1].vt ||
		(VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt)
	{
		return DISP_E_TYPEMISMATCH;
	}
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (this->m_pBackObj->m_pMyOPTFile->GetExtraScanArray(
		varg.lVal, &nArray, &pWaves, &pValues, szTitle, MAX_PATH))
	{
		sab.lLbound = 0;
		sab.cElements = nArray;
		*(pDispParams->rgvarg[1].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
		*(pDispParams->rgvarg[0].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**)&arrWaves);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arrValues);
		for (i = 0; i < nArray; i++)
		{
			arrWaves[i] = pWaves[i];
			arrValues[i] = pValues[i];
		}
		SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		// title string
		InitVariantFromString(szTitle, pVarResult);
		// cleanup
		delete[] pWaves;
		delete[] pValues;
	}
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::GetGratingScan(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				nArray;
	double			*	pWaves;
	double			*	pValues;
	short int			grating;
	TCHAR				szFilter[MAX_PATH];
	short int			detector;
	SAFEARRAYBOUND		sab;
	long				i;
	double			*	arrWaves;
	double			*	arrValues;
	BOOL				fSuccess = FALSE;


	if (5 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
//	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[1].vt ||
//		(VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt)
	if (0 == (pDispParams->rgvarg[1].vt & VT_BYREF) ||
		0 == (pDispParams->rgvarg[0].vt & VT_BYREF))
	{
		return DISP_E_TYPEMISMATCH;
	}
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.iVal;
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szFilter, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 2, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	detector = varg.iVal;
	fSuccess = this->m_pBackObj->m_pMyOPTFile->GetSingleGratingScan(
		grating, szFilter, detector, &nArray, &pWaves, &pValues);
	if (fSuccess)
	{
		sab.lLbound = 0;
		sab.cElements = nArray;
		VariantInit(pDispParams->rgvarg[1].pvarVal);
		pDispParams->rgvarg[1].pvarVal->vt = VT_ARRAY | VT_R8;
		pDispParams->rgvarg[1].pvarVal->parray = SafeArrayCreate(VT_R8, 1, &sab);
		VariantInit(pDispParams->rgvarg[0].pvarVal);
		pDispParams->rgvarg[0].pvarVal->vt = VT_ARRAY | VT_R8;
		pDispParams->rgvarg[0].pvarVal->parray = SafeArrayCreate(VT_R8, 1, &sab);
//		*(pDispParams->rgvarg[1].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
//		*(pDispParams->rgvarg[0].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pDispParams->rgvarg[1].pvarVal->parray, (void**)&arrWaves);
		SafeArrayAccessData(pDispParams->rgvarg[0].pvarVal->parray, (void**)&arrValues);
		for (i = 0; i < nArray; i++)
		{
			arrWaves[i] = pWaves[i];
			arrValues[i] = pValues[i];
		}
		SafeArrayUnaccessData(pDispParams->rgvarg[1].pvarVal->parray);
		SafeArrayUnaccessData(pDispParams->rgvarg[0].pvarVal->parray);
		// cleanup
		delete[] pWaves;
		delete[] pValues;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIDataSet::ExportToCSV(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pMyOPTFile->ExportToCSV(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

/*
' GetADGainFactor returns true if ADGainFactor available
' wavelength in nm
' AcquireMode = 1 for MultiChannel, 2 for pulsed mode
' LockinGain factor
' detector = 1 for PMT, 2 for UVS , 3 for PbS, 4 for PbSe
' detector gain = "Hi" or "Lo"
*/
HRESULT CMyObject::CImp_clsIDataSet::GetADGainFactor(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	// input wavelength
	VARIANTARG		varg;
	UINT			uArgErr;
	double			wavelength;
	HRESULT			hr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	wavelength = varg.dblVal;

	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(
		this->m_pBackObj->m_pMyOPTFile->DetermineDataGainFactor(wavelength),
		pVarResult);
	return S_OK;
}

CMyObject::CImpIPersistStreamInit::CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIPersistStreamInit::~CImpIPersistStreamInit()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetClassID(
									CLSID		*	pClassID)
{
	*pClassID	= MY_CLSID;
	return S_OK;
}

// IPersistStreamInit methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetSizeMax(
									ULARGE_INTEGER *pCbSize)
{
	pCbSize->QuadPart	= this->m_pBackObj->m_pMyOPTFile->GetSaveSize();
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::InitNew()
{
	return this->m_pBackObj->m_pMyOPTFile->InitNew();
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::IsDirty()
{
	return this->m_pBackObj->m_pMyOPTFile->GetAmDirty() ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Load(
									LPSTREAM		pStm)
{
	return this->m_pBackObj->m_pMyOPTFile->LoadFromStream(pStm);
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Save(
									LPSTREAM		pStm,
									BOOL			fClearDirty)
{
	return this->m_pBackObj->m_pMyOPTFile->SaveToStream(pStm, fClearDirty);
}

CMyObject::CImpIPersistFile::CImpIPersistFile(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_szFileName(NULL),
	m_fNoScribble(FALSE)
{

}

CMyObject::CImpIPersistFile::~CImpIPersistFile()
{
	if (NULL != this->m_szFileName)
	{
		CoTaskMemFree((LPVOID)this->m_szFileName);
		this->m_szFileName = NULL;
	}
}
// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistFile::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistFile::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistFile::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistFile::GetClassID(
	CLSID		*	pClassID)
{
	*pClassID = MY_CLSID;
	return S_OK;
}

// IPersistFile methods
STDMETHODIMP CMyObject::CImpIPersistFile::GetCurFile(
	LPOLESTR	*	ppszFileName)
{
	*ppszFileName = NULL;
	if (NULL != this->m_szFileName)
	{
		SHStrDup(this->m_szFileName, ppszFileName);
		return S_OK;
	}
	else
	{
		SHStrDup(L".opt", ppszFileName);
		return S_FALSE;
	}
}

STDMETHODIMP CMyObject::CImpIPersistFile::IsDirty()
{
	return this->m_pBackObj->m_pMyOPTFile->GetAmDirty() ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIPersistFile::Load(
	LPCOLESTR		pszFileName,
	DWORD			dwMode)
{
	if (NULL != this->m_szFileName)
	{
		CoTaskMemFree((LPVOID)this->m_szFileName);
		this->m_szFileName = NULL;
	}
	if (this->m_pBackObj->m_pMyOPTFile->LoadFromFile(pszFileName))
	{
		SHStrDup(pszFileName, &this->m_szFileName);
		return S_OK;
	}
	else
		return E_FAIL;
}

STDMETHODIMP CMyObject::CImpIPersistFile::Save(
	LPCOLESTR		pszFileName,
	BOOL			fRemember)
{
	if (NULL == pszFileName)
	{
		// save case
		if (!this->m_fNoScribble && NULL != this->m_szFileName)
		{
			if (this->m_pBackObj->m_pMyOPTFile->SaveToFile(this->m_szFileName, TRUE))
			{
				this->m_fNoScribble = TRUE;
				return S_OK;
			}
			else
				return E_FAIL;
		}
		else return E_UNEXPECTED;
	}
	else
	{
		if (fRemember)
		{
			// save as case
			if (NULL != this->m_szFileName)
			{
				CoTaskMemFree((LPVOID)this->m_szFileName);
				this->m_szFileName = NULL;
			}
			if (this->m_pBackObj->m_pMyOPTFile->SaveToFile(pszFileName, TRUE))
			{
				SHStrDup(pszFileName, &this->m_szFileName);
				this->m_fNoScribble = TRUE;
				return S_OK;
			}
			else
			{
				return E_FAIL;
			}
		}
		else
		{
			// save a copy as
			return this->m_pBackObj->m_pMyOPTFile->SaveToFile(pszFileName, FALSE) ? S_OK : E_FAIL;
		}
	}
}
STDMETHODIMP CMyObject::CImpIPersistFile::SaveCompleted(
	LPCOLESTR		pszFileName)
{
	this->m_fNoScribble = FALSE;
	return S_OK;
}

CMyObject::CImpISummaryInfo::CImpISummaryInfo(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_dispidComment(DISPID_UNKNOWN),
	m_dispidXData(DISPID_UNKNOWN),
	m_dispidYData(DISPID_UNKNOWN),
	m_dispidAcquisitionDate(DISPID_UNKNOWN),
	m_dispidXRange(DISPID_UNKNOWN),
	m_dispidYRange(DISPID_UNKNOWN)
{
}

CMyObject::CImpISummaryInfo::~CImpISummaryInfo()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpISummaryInfo::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpISummaryInfo::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpISummaryInfo::Release()
{
	return this->m_punkOuter->Release();
}
// IDispatch methods
STDMETHODIMP CMyObject::CImpISummaryInfo::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpISummaryInfo::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	if (NULL != this->m_pTypeInfo)
	{
		*ppTInfo = this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
		return S_OK;
	}
	else
	{
		*ppTInfo = NULL;
		return E_FAIL;
	}
}

STDMETHODIMP CMyObject::CImpISummaryInfo::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo;
	hr = this->GetTypeInfo(0, lcid, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpISummaryInfo::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	if (dispIdMember == m_dispidComment)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetComment(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidXData)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetXData(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidYData)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetYData(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidAcquisitionDate)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAcquisitionDate(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidXRange)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetXRange(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == m_dispidYRange)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetYRange(pDispParams, pVarResult);
		}
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpISummaryInfo::MyInit()
{
	HRESULT				hr;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pBackObj->GetRefTypeInfo(TEXT("ISummaryInfo"), &(this->m_pTypeInfo));
	if (SUCCEEDED(hr))
	{
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pBackObj->m_iidISummayInfo = pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(this->m_pTypeInfo, L"Comment", &m_dispidComment);
		Utils_GetMemid(this->m_pTypeInfo, L"XData", &m_dispidXData);
		Utils_GetMemid(this->m_pTypeInfo, L"YData", &m_dispidYData);
		Utils_GetMemid(this->m_pTypeInfo, L"AcquisitionDate", &m_dispidAcquisitionDate);
		Utils_GetMemid(this->m_pTypeInfo, L"XRange", &m_dispidXRange);
		Utils_GetMemid(this->m_pTypeInfo, L"YRange", &m_dispidYRange);
	}
	return hr;
}

HRESULT CMyObject::CImpISummaryInfo::GetComment(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szComment[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetComment(szComment, MAX_PATH))
	{
		InitVariantFromString(szComment, pVarResult);
	}
	return S_OK;
}
HRESULT CMyObject::CImpISummaryInfo::GetXData(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long			nData = 0;
	double		*	pData = NULL;
	if (this->m_pBackObj->m_pMyOPTFile->GetWavelengths(&nData, &pData))
	{
		InitVariantFromDoubleArray(pData, (ULONG)nData, pVarResult);
		delete[] pData;
	}
	return S_OK;
}
HRESULT CMyObject::CImpISummaryInfo::GetYData(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long			nData = 0;
	double		*	pData = NULL;
	if (this->m_pBackObj->m_pMyOPTFile->GetSpectra(&nData, &pData))
	{
		InitVariantFromDoubleArray(pData, (ULONG)nData, pVarResult);
		delete[] pData;
	}
	return S_OK;
}
HRESULT	CMyObject::CImpISummaryInfo::GetAcquisitionDate(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szTimeStamp[MAX_PATH];
	if (this->m_pBackObj->m_pMyOPTFile->GetTimeStamp(szTimeStamp, MAX_PATH))
	{
		VariantInit(pVarResult);
		pVarResult->vt = VT_DATE;
		VarDateFromStr(szTimeStamp, LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE, &pVarResult->date);
	}
	return S_OK;
}
HRESULT CMyObject::CImpISummaryInfo::GetXRange(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long			nData = 0;
	double		*	pData = NULL;
	double			xRange[2];
	long			i;
	if (this->m_pBackObj->m_pMyOPTFile->GetWavelengths(&nData, &pData))
	{
		xRange[0] = pData[0];
		xRange[1] = pData[1];
		for (i = 1; i < nData; i++)
		{
			xRange[0] = xRange[0] < pData[i] ? xRange[0] : pData[i];
			xRange[1] = xRange[1] >= pData[i] ? xRange[1] : pData[i];
		}
		delete[] pData;
		*(pDispParams->rgvarg[0].pdblVal) = xRange[0];
		InitVariantFromDouble(xRange[1], pVarResult);
	}
	return S_OK;
}
HRESULT CMyObject::CImpISummaryInfo::GetYRange(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long			nData = 0;
	double		*	pData = NULL;
	double			yRange[2];
	long			i;
	if (this->m_pBackObj->m_pMyOPTFile->GetSpectra(&nData, &pData))
	{
		yRange[0] = pData[0];
		yRange[1] = pData[1];
		for (i = 1; i < nData; i++)
		{
			yRange[0] = yRange[0] < pData[i] ? yRange[0] : pData[i];
			yRange[1] = yRange[1] >= pData[i] ? yRange[1] : pData[i];
		}
		delete[] pData;
		*(pDispParams->rgvarg[0].pdblVal) = yRange[0];
		InitVariantFromDouble(yRange[1], pVarResult);
	}

	return S_OK;
}
