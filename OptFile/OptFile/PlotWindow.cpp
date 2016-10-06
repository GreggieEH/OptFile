#include "stdafx.h"
#include "PlotWindow.h"
#include "dispids.h"
#include "PlotWindowEventHandler.h"

CPlotWindow::CPlotWindow(CPlotWindowEventHandler * pPlotWindowEventHandler) :
	CMyContainer(),
	m_pPlotWindowEventHandler(pPlotWindowEventHandler),
	m_hwndParent(NULL),
	m_nextCurveIndex(200)
{
}

CPlotWindow::~CPlotWindow()
{
}

BOOL CPlotWindow::CreatePlotWindow(
	HWND			hwnd)
{
	HRESULT			hr;
	RECT			rc;
	IDispatch	*	pdispMyContainer;
	VARIANTARG		avarg[7];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	LPTSTR			szStorageFile = NULL;
//	// store the main window
//	this->m_hwndParent = hwndParent;
//	GetClientRect(this->m_hwndParent, &rc);
//	this->Init();
	HWND			hwndParent = GetParent(hwnd);

	ShowWindow(hwnd, SW_HIDE);
	// store the main window
	this->m_hwndParent = hwndParent;
	GetWindowRect(hwnd, &rc);
	MapWindowPoints(NULL, hwndParent, (LPPOINT)&rc, 2);

	this->Init();

	hr = this->QueryInterface(IID_IDispatch, (LPVOID*)&pdispMyContainer);
	if (SUCCEEDED(hr))
	{
		// form the parameters
		InitVariantFromString(L"Sciencetech.SciPlotWindow.1", &avarg[6]);
		InitVariantFromString(PLOT_WINDOW, &avarg[5]);
		InitVariantFromString(szStorageFile, &avarg[4]);
		InitVariantFromInt32(rc.left, &avarg[3]);
		InitVariantFromInt32(rc.top, &avarg[2]);
		InitVariantFromInt32(rc.right, &avarg[1]);
		InitVariantFromInt32(rc.bottom, &avarg[0]);
		hr = Utils_InvokeMethod(pdispMyContainer, DISPID_Container_CreateObject, avarg, 7, &varResult);
		VariantClear(&avarg[6]);
		VariantClear(&avarg[5]);
		VariantClear(&avarg[4]);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdispMyContainer->Release();
	}
	return fSuccess;
}

void CPlotWindow::ClosePlotWindow()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	hr = this->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_InvokeMethod(pdisp, DISPID_Container_CloseAllObjects, NULL, 0, NULL);
		pdisp->Release();
	}
}


HWND CPlotWindow::GetMainWindow()				// pure virtual
{
	return this->m_hwndParent;
}
void CPlotWindow::OnObjectCreated(
	LPCTSTR			szObject,
	IUnknown	*	pUnkSite)
{
	CMyContainer::OnObjectCreated(szObject, pUnkSite);
}

BOOL CPlotWindow::OnPropRequestEdit(
	LPCTSTR			szObject,
	DISPID			dispid)
{
	return TRUE;
}

void CPlotWindow::OnPropChanged(
	LPCTSTR			szObject,
	DISPID			dispid)
{
}

void CPlotWindow::OnSinkEvent(
	LPCTSTR			ObjectName,
	long			Dispid,
	long			nParams,
	VARIANT		*	Parameters)
{
	if (0 == lstrcmpi(ObjectName, PLOT_WINDOW))
	{
		if (DISPID_MOUSEMOVE == Dispid && 4 == nParams)
		{
			double			x;
			double			y;
			VariantToDouble(Parameters[1], &x);
			VariantToDouble(Parameters[0], &y);
			this->m_pPlotWindowEventHandler->OnPlotWindowMouseMove(x, y);
		}
	}
}

HRESULT CPlotWindow::MySetActiveObject(
	IOleInPlaceActiveObject *pActiveObject)
{
	return S_OK;
}

BOOL CPlotWindow::SetPlotData(
	long			nValues,
	double		*	pX,
	double		*	pY)
{
	if (nValues > 0 && NULL != pX && NULL != pY)
	{
		// check if the curve exists
		long		index = this->FindCurve(OUR_CURVE);
		if (index < 0)
		{
			this->AddDataSet(OUR_CURVE, RGB(0, 0, 255));
		}
		this->SetCurveData(OUR_CURVE, nValues, pX, pY);
		this->ZoomOut();
	}
	else
	{
		this->RemoveSpectra(OUR_CURVE);		// remove the curve
	}
	return TRUE;
}

BOOL CPlotWindow::GetOurObject(
	IDispatch	**	ppdisp)
{
	HRESULT				hr;
	hr = this->GetActiveX(PLOT_WINDOW, ppdisp);
	return SUCCEEDED(hr);
}

BOOL CPlotWindow::AddDataSet(
	LPCTSTR			szName,
	unsigned long	curveColor)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"AddDataSet", &dispid);
		InitVariantFromString(szName, &avarg[1]);
		InitVariantFromUInt32(curveColor, &avarg[0]);
		hr = Utils_InvokeMethod(pdisp, dispid, avarg, 2, &varResult);
		VariantClear(&avarg[1]);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdisp->Release();
	}
	return fSuccess;
}
void CPlotWindow::SetCurveData(
	LPCTSTR			szName,
	long			nValues,
	double		*	pX,
	double		*	pY)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	if (this->GetOurObject(&pdisp))
	{
		InitVariantFromString(szName, &avarg[1]);
		Utils_GetMemid(pdisp, L"CurveXData", &dispid);
		InitVariantFromDoubleArray(pX, nValues, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, dispid, avarg, 2);
		VariantClear(&avarg[0]);
		Utils_GetMemid(pdisp, L"CurveYData", &dispid);
		InitVariantFromDoubleArray(pY, nValues, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, dispid, avarg, 2);
		VariantClear(&avarg[0]);
		VariantClear(&avarg[1]);
		pdisp->Release();
	}
}

BOOL CPlotWindow::ZoomOut()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"ZoomOut", &dispid);
		hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdisp->Release();
	}
	return fSuccess;
}

long CPlotWindow::GetNumCurves()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANT			varResult;
	long			nCurves = 0;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"NumberOfCurves", &dispid);
		hr = Utils_InvokePropertyGet(pdisp, dispid, NULL, 0, &varResult);
		if (SUCCEEDED(hr)) VariantToInt32(varResult, &nCurves);
		pdisp->Release();
	}
	return nCurves;
}

void CPlotWindow::Refresh()
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_InvokeMethod(pdisp, DISPID_REFRESH, NULL, 0, NULL);
		pdisp->Release();
	}
}

void CPlotWindow::SetXRange(
	double			xMin,
	double			xMax)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"XRange", &dispid);
		VariantInit(&avarg[1]);
		avarg[1].vt = VT_BYREF | VT_R8;
		avarg[1].pdblVal = &xMin;
		InitVariantFromDouble(xMax, &avarg[0]);
		Utils_DoInvoke(pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 2, NULL);
		pdisp->Release();
	}
}

void CPlotWindow::GetYRange(
	double		*	yMin,
	double		*	yMax)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	*yMin = 0.0;
	*yMax = 0.0;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"YRange", &dispid);
		VariantInit(&varg);
		varg.vt = VT_BYREF | VT_R8;
		varg.pdblVal = yMin;
		hr = Utils_InvokePropertyGet(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, yMax);
		pdisp->Release();
	}
}

void CPlotWindow::SetYRange(
	double			yMin,
	double			yMax)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"YRange", &dispid);
		VariantInit(&avarg[1]);
		avarg[1].vt = VT_BYREF | VT_R8;
		avarg[1].pdblVal = &yMin;
		InitVariantFromDouble(yMax, &avarg[0]);
		Utils_DoInvoke(pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 2, NULL);
		pdisp->Release();
	}
}

void CPlotWindow::GetDataXRange(
	double		*	xMin,
	double		*	xMax)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	*xMin = 0.0;
	*xMax = 0.0;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"DataXRange", &dispid);
		VariantInit(&varg);
		varg.vt = VT_BYREF | VT_R8;
		varg.pdblVal = xMin;
		hr = Utils_DoInvoke(pdisp, dispid, DISPATCH_PROPERTYGET, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, xMax);
		pdisp->Release();
	}
}

void CPlotWindow::GetDataYRange(
	double			xMin,
	double			xMax,
	double		*	yMin,
	double		*	yMax)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[3];
	VARIANT			varResult;
	*yMin = 0.0;
	*yMax = 0.0;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"DataYRange", &dispid);
		InitVariantFromDouble(xMin, &avarg[2]);
		InitVariantFromDouble(xMax, &avarg[1]);
		VariantInit(&avarg[0]);
		avarg[0].vt = VT_BYREF | VT_R8;
		avarg[0].pdblVal = yMin;
		hr = Utils_DoInvoke(pdisp, dispid, DISPATCH_PROPERTYGET, avarg, 3, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, yMax);
		pdisp->Release();
	}
}

BOOL CPlotWindow::TryMNemonic(
	LPMSG			pmsg)
{
	HRESULT			hr;
	IDispatch	*	pdispMyContainer;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;

	hr = this->QueryInterface(IID_IDispatch, (LPVOID*)&pdispMyContainer);
	if (SUCCEEDED(hr))
	{
		InitVariantFromBuffer((const void*)pmsg, sizeof(MSG), &varg);
		hr = Utils_InvokeMethod(pdispMyContainer, DISPID_Container_TryMNemonic, &varg, 1, &varResult);
		VariantClear(&varg);
		if (SUCCEEDED(hr))VariantToBoolean(varResult, &fSuccess);
		pdispMyContainer->Release();
	}
	return fSuccess;
}

void CPlotWindow::SetTitles(
	LPCTSTR			szTitle,
	LPCTSTR			szXTitle,
	LPCTSTR			szYTitle)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"Title", &dispid);
		Utils_SetStringProperty(pdisp, dispid, szTitle);
		Utils_GetMemid(pdisp, L"XTitle", &dispid);
		Utils_SetStringProperty(pdisp, dispid, szXTitle);
		Utils_GetMemid(pdisp, L"YTitle", &dispid);
		Utils_SetStringProperty(pdisp, dispid, szYTitle);
		pdisp->Release();
	}
}

void CPlotWindow::SetCurveVisible(
	LPCTSTR			szName,
	BOOL			fVisible)
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	VARIANTARG			avarg[2];
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveVisible", &dispid);
		InitVariantFromString(szName, &avarg[1]);
		InitVariantFromBoolean(fVisible, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, dispid, avarg, 2);
		VariantClear(&avarg[1]);
		pdisp->Release();
	}
}

HWND CPlotWindow::GetControlWindow()
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	IOleWindow	*	pOleWindow = NULL;
	HWND			hwndPlot = NULL;
	if (this->GetOurObject(&pdisp))
	{
		hr = pdisp->QueryInterface(IID_IOleWindow, (LPVOID*)&pOleWindow);
		pdisp->Release();
	}
	else
		hr = E_FAIL;
	if (SUCCEEDED(hr))
	{
		hr = pOleWindow->GetWindow(&hwndPlot);
		pOleWindow->Release();
	}
	return hwndPlot;
}

BOOL CPlotWindow::GetZoomedIn()
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	BOOL			zoomedIn = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"ZoomedIn", &dispid);
		zoomedIn = Utils_GetBoolProperty(pdisp, dispid);
		pdisp->Release();
	}
	return zoomedIn;
}

BOOL CPlotWindow::GetCurveXData(
	long			curveIndex,
	long		*	nArray,
	double		**	ppX)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			ub;
	double		*	arr;
	BOOL			fSuccess = FALSE;
	long			i;
	*nArray = 0;
	*ppX = NULL;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveXData", &dispid);
		InitVariantFromInt32(curveIndex, &varg);
		hr = Utils_InvokePropertyGet(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
			{
				fSuccess = TRUE;
				SafeArrayGetUBound(varResult.parray, 1, &ub);
				*nArray = ub + 1;
				*ppX = new double[*nArray];
				SafeArrayAccessData(varResult.parray, (void**)&arr);
				for (i = 0; i < (*nArray); i++)
				{
					(*ppX)[i] = arr[i];
				}
				SafeArrayUnaccessData(varResult.parray);
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

void CPlotWindow::SetCurveXData(
	long			curveIndex,
	long			nArray,
	double		*	pX)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];

	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveXData", &dispid);
		InitVariantFromInt32(curveIndex, &avarg[1]);
		InitVariantFromDoubleArray(pX, (ULONG)nArray, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, dispid, avarg, 2);
		VariantClear(&avarg[0]);
		pdisp->Release();
	}
}

BOOL CPlotWindow::GetCurveYData(
	long			curveIndex,
	long		*	nArray,
	double		**	ppY)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			ub;
	double		*	arr;
	BOOL			fSuccess = FALSE;
	long			i;
	*nArray = 0;
	*ppY = NULL;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveYData", &dispid);
		InitVariantFromInt32(curveIndex, &varg);
		hr = Utils_InvokePropertyGet(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
			{
				fSuccess = TRUE;
				SafeArrayGetUBound(varResult.parray, 1, &ub);
				*nArray = ub + 1;
				*ppY = new double[*nArray];
				SafeArrayAccessData(varResult.parray, (void**)&arr);
				for (i = 0; i < (*nArray); i++)
				{
					(*ppY)[i] = arr[i];
				}
				SafeArrayUnaccessData(varResult.parray);
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

void CPlotWindow::SetCurveYData(
	long			curveIndex,
	long			nArray,
	double		*	pY)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];

	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveYData", &dispid);
		InitVariantFromInt32(curveIndex, &avarg[1]);
		InitVariantFromDoubleArray(pY, (ULONG)nArray, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, dispid, avarg, 2);
		VariantClear(&avarg[0]);
		pdisp->Release();
	}
}

BOOL CPlotWindow::AddDataValue(
	long			Curve,
	double			xValue,
	double			yValue)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[3];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"AddDataValue", &dispid);
		InitVariantFromInt32(Curve, &avarg[2]);
		InitVariantFromDouble(xValue, &avarg[1]);
		InitVariantFromDouble(yValue, &avarg[0]);
		hr = Utils_InvokeMethod(pdisp, dispid, avarg, 3, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdisp->Release();
	}
	return fSuccess;
}

void CPlotWindow::GetCurveName(
	long			index,
	LPTSTR			szCurve,
	UINT			nBufferSize)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	szCurve[0] = '\0';
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurveName", &dispid);
		InitVariantFromInt32(index, &varg);
		hr = Utils_InvokePropertyGet(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, szCurve, nBufferSize);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}

BOOL CPlotWindow::RemoveSpectra(
	LPCTSTR			szCurve)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetMemid(pdisp, L"RemoveSpectra", &dispid);
		InitVariantFromString(szCurve, &varg);
		hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
		VariantClear(&varg);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdisp->Release();
	}
	return fSuccess;
}

long CPlotWindow::FindCurve(
	LPCTSTR			szName)
{
	long			numCurves = this->GetNumCurves();
	TCHAR			szCurve[MAX_PATH];
	long			i = 0;
	BOOL			fDone = FALSE;
	long			retVal = -1;

	while (i < numCurves && !fDone)
	{
		this->GetCurveName(i, szCurve, MAX_PATH);
		fDone = 0 == lstrcmpi(szCurve, szName);
		if (fDone)
			retVal = i;
		else
			i++;
	}
	return retVal;
}

void CPlotWindow::ClearPlotData()
{
	TCHAR			szCurve[MAX_PATH];
	long			numCurves = this->GetNumCurves();
	long			i;
	for (i = numCurves = 1; i >= 0; i--)
	{
		this->GetCurveName(i, szCurve, MAX_PATH);
		this->RemoveSpectra(szCurve);
	}
}
