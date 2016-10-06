#include "stdafx.h"
#include "PropPageGratingScanInfo.h"
#include "GratingScanInfo.h"
#include "dispids.h"
//#include "MyPlotWindow.h"
#include "SlitInfo.h"
#include "PlotWindow.h"

CPropPageGratingScanInfo::CPropPageGratingScanInfo(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_GratingScanInfo, IDS_GratingScanInfo),
	m_pMyPlotWindow(NULL)
{
}

CPropPageGratingScanInfo::~CPropPageGratingScanInfo(void)
{
//	Utils_DELETE_POINTER(this->m_pMyPlotWindow);
	if (NULL != this->m_pMyPlotWindow)
		this->m_pMyPlotWindow->ClosePlotWindow();
	Utils_RELEASE_INTERFACE(this->m_pMyPlotWindow);
}

BOOL CPropPageGratingScanInfo::OnInitDialog()
{
	// subclass the edit box
	this->SubclassEditBox(IDC_EDITSCANSEGMENT);
//	// subclass the plot window
//	this->SubclassControl(IDC_PLOT);
//	// create the plot window object
//	this->m_pMyPlotWindow		= new CMyPlotWindow(GetDlgItem(this->GetMyPage(), IDC_PLOT));
	this->m_pMyPlotWindow = new CPlotWindow(this);
	this->m_pMyPlotWindow->AddRef();
	this->m_pMyPlotWindow->CreatePlotWindow(GetDlgItem(this->GetMyPage(), IDC_PLOT));

	// number of scan segments
	SetDlgItemInt(this->GetMyPage(), IDC_NUMBEROFSCANSEGMENTS, this->GetNumGratingScans(), FALSE);
	// display the first scan segment
	SetDlgItemInt(this->GetMyPage(), IDC_EDITSCANSEGMENT, 0, FALSE);
	this->DisplayGratingScan(0);
	// display the time stamp
	this->DisplayTimeStamp();
	return CMyPropPage::OnInitDialog();
}

BOOL CPropPageGratingScanInfo::OnNotify(
						LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code && IDC_UPDSCANSEGMENT == pnmh->idFrom)
	{
		long		nSegment	= GetDlgItemInt(this->GetMyPage(), IDC_EDITSCANSEGMENT, NULL, FALSE);
		LPNMUPDOWN	pnmud		= (LPNMUPDOWN) pnmh;
		nSegment	-= pnmud->iDelta;
		long		nSegments	= this->GetNumGratingScans();
		if (nSegment < 0) nSegment = nSegments-1;
		if (nSegment >= nSegments) nSegment = 0;
		SetDlgItemInt(this->GetMyPage(), IDC_EDITSCANSEGMENT, nSegment, FALSE);
		this->DisplayGratingScan(nSegment);
		return TRUE;
	}
	return FALSE;
}

void CPropPageGratingScanInfo::OnEditReturnClicked(
						UINT			nID)
{
	long		nSegment	= GetDlgItemInt(this->GetMyPage(), IDC_EDITSCANSEGMENT, NULL, FALSE);
	this->DisplayGratingScan(nSegment);
}

/*
LRESULT CPropPageGratingScanInfo::SubclassProc(
						HWND			hwnd,
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam)
{
	UINT			nID		= GetDlgCtrlID(hwnd);
	if (IDC_PLOT == nID)
	{
		switch (uMsg)
		{
		case WM_PAINT:
			if (NULL != this->m_pMyPlotWindow)
			{
				this->m_pMyPlotWindow->OnPaint();
				return 0;
			}
			break;
		case WM_MOUSEMOVE:
			if (NULL != this->m_pMyPlotWindow)
			{
				TCHAR			szString[MAX_PATH];
				szString[0]	= '\0';
				this->m_pMyPlotWindow->OnMouseMove(
					GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), szString, MAX_PATH);
				SetDlgItemText(this->GetMyPage(), IDC_STATUS, szString);
			}
		default:
			break;
		}
	}
	return CMyPropPage::SubclassProc(hwnd, uMsg, wParam, lParam);
}
*/

void CPropPageGratingScanInfo::OnPlotWindowMouseMove(
	double			x,
	double			y)
{
	TCHAR			szString[MAX_PATH];
	_stprintf_s(szString, MAX_PATH, TEXT("X = %4.1f  Y = %8.4g"), x, y);
	SetDlgItemText(this->GetMyPage(), IDC_STATUS, szString);
}


void CPropPageGratingScanInfo::DisplayGratingScan(
						long			scanIndex)
{
	IDispatch		*	pdisp;
	CGratingScanInfo	gratingScan;
	TCHAR				szFilter[MAX_PATH];
	long				nValues;
	double			*	pX;
	double			*	pY;
	TCHAR				szDispersion[MAX_PATH];
	double				dispersion;
	double				outputSlitWidth;
	TCHAR				szMessage[MAX_PATH];

	if (NULL == this->m_pMyPlotWindow) return;
	this->m_pMyPlotWindow->ClearPlotData();
	SetDlgItemText(this->GetMyPage(), IDC_GRATING, TEXT(""));
	SetDlgItemText(this->GetMyPage(), IDC_FILTER, TEXT(""));
	if (this->GetGratingScan(scanIndex, &pdisp))
	{
		gratingScan.doInit(pdisp);
		SetDlgItemInt(this->GetMyPage(), IDC_GRATING, gratingScan.Getgrating(), FALSE);
		if (gratingScan.Getfilter(szFilter, MAX_PATH))
			SetDlgItemText(this->GetMyPage(), IDC_FILTER, szFilter);
		dispersion = gratingScan.Getdispersion();
		outputSlitWidth = this->GetOutputSlitWidth();
		_stprintf_s(szMessage, L"Dispersion = %5.2f outputslit width = %5.2f", dispersion, outputSlitWidth);
//		MessageBox(this->GetMyPage(), szMessage, L"OPTFile", MB_OK);

		_stprintf_s(szDispersion, MAX_PATH, TEXT("%5.2f"), (float) gratingScan.Getdispersion() * this->GetOutputSlitWidth());
		SetDlgItemText(this->GetMyPage(), IDC_EDITDISPERSION, szDispersion);
		if (gratingScan.GetGratingScan(&nValues, &pX, &pY))
		{
			this->m_pMyPlotWindow->SetPlotData(nValues, pX, pY);
			delete [] pX;
			delete [] pY;
		}
		pdisp->Release();
	}
//	// replot
//	InvalidateRect(GetDlgItem(this->GetMyPage(), IDC_PLOT), NULL, TRUE); 
	this->m_pMyPlotWindow->Refresh();
}

long CPropPageGratingScanInfo::GetNumGratingScans()
{
	IDispatch	*	pdisp;
	long			nScans		= 0;
	if (this->GetOurObject(&pdisp))
	{
		nScans	= Utils_GetLongProperty(pdisp, DISPID_NumGratingScans);
		pdisp->Release();
	}
	return nScans;
}

BOOL CPropPageGratingScanInfo::GetGratingScan(
						long			index,
						IDispatch	**	ppdispGratingScan)
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*ppdispGratingScan			= NULL;
	if (this->GetOurObject(&pdisp))
	{
		InitVariantFromInt32(index, &varg);
		hr = Utils_InvokeMethod(pdisp, DISPID_GetGratingScan, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
			{
				*ppdispGratingScan	= varResult.pdispVal;
				varResult.pdispVal->AddRef();
				fSuccess	= TRUE;
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

void CPropPageGratingScanInfo::DisplayTimeStamp()
{
	LPTSTR			szTimeStamp	= NULL;
	IDispatch	*	pdisp;

	SetDlgItemText(this->GetMyPage(), IDC_ACQUIRETIME, TEXT(""));
	if (this->GetOurObject(&pdisp))
	{
		Utils_GetStringProperty(pdisp, DISPID_TimeStamp, &szTimeStamp);
		SetDlgItemText(this->GetMyPage(), IDC_ACQUIRETIME, szTimeStamp);
		CoTaskMemFree((LPVOID) szTimeStamp);
		pdisp->Release();
	}
}

double CPropPageGratingScanInfo::GetOutputSlitWidth()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varSlitInfo;
	double			slitWidth = 1.0;
	CSlitInfo		slitInfo;

	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_SlitInfo, NULL, 0, &varSlitInfo);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varSlitInfo.vt && NULL != varSlitInfo.pdispVal)
			{
				if (slitInfo.doInit(varSlitInfo.pdispVal))
				{
					slitInfo.GetOutputSlitWidth(&slitWidth);
				}
			}
			VariantClear(&varSlitInfo);
		}
		pdisp->Release();
	}
	return slitWidth;
}


