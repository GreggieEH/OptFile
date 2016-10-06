#include "stdafx.h"
#include "PropPageCalibrationStandard.h"
#include "dispids.h"
#include "CalibrationStandard.h"
//#include "MyPlotWindow.h"
#include "SciFileOpen.h"
#include "PlotWindow.h"

CPropPageCalibrationStandard::CPropPageCalibrationStandard(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_CalibrationStandard, IDS_CalibrationStandard),
	m_pMyPlotWindow(NULL),
	m_fAllowClose(TRUE)
{
}

CPropPageCalibrationStandard::~CPropPageCalibrationStandard(void)
{
//	Utils_DELETE_POINTER(this->m_pMyPlotWindow);
	if (NULL != this->m_pMyPlotWindow)
		this->m_pMyPlotWindow->ClosePlotWindow();
	Utils_RELEASE_INTERFACE(this->m_pMyPlotWindow);
}

BOOL CPropPageCalibrationStandard::OnInitDialog()
{
	this->SubclassEditBox(IDC_EDITFILENAME);
	this->SubclassEditBox(IDC_EDITEXPECTEDDISTANCE);
	this->SubclassEditBox(IDC_EDITSOURCEDISTANCE);
//	this->SubclassControl(IDC_PLOT);
//	// create the plot window object
//	this->m_pMyPlotWindow		= new CMyPlotWindow(GetDlgItem(this->GetMyPage(), IDC_PLOT));
	this->m_pMyPlotWindow = new CPlotWindow(this);
	this->m_pMyPlotWindow->AddRef();
	this->m_pMyPlotWindow->CreatePlotWindow(GetDlgItem(this->GetMyPage(), IDC_PLOT));

	// display the current data
	this->DisplayCurrentData();
	this->DisplayexpectedDistance();
	this->DisplaysourceDistance();
	return CMyPropPage::OnInitDialog();
}

BOOL CPropPageCalibrationStandard::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_BROWSE:
		this->m_fAllowClose = FALSE;
		this->SelectCalibrationFile();
		this->DisplayCurrentData();
		this->m_fAllowClose = TRUE;
		return TRUE;
	case IDC_LOADREFCALIBRATION:
		this->m_fAllowClose = FALSE;
		this->DisplayCurrentData();
		this->m_fAllowClose = TRUE;
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

/*
LRESULT CPropPageCalibrationStandard::SubclassProc(
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

BOOL CPropPageCalibrationStandard::OnNotify(
	LPNMHDR			pnmh)
{
	if (PSN_KILLACTIVE == pnmh->code)
	{
		if (!this->m_fAllowClose)
		{
			SetWindowLongPtr(this->GetMyPage(), DWLP_MSGRESULT, (LONG)TRUE);
			return TRUE;
		}
		if (!this->GetAmLoaded())
		{
			MessageBox(this->GetMyPage(), L"Reference Calibration Not Loaded",
				L"OPTFile", MB_OK);
			SetWindowLongPtr(this->GetMyPage(), DWLP_MSGRESULT, (LONG)TRUE);
		}
		else
		{
			// update source distance and expected distance
			this->UpdateexpectedDistance();
			this->UpdatesourceDistance();
			// allow closing
			SetWindowLongPtr(this->GetMyPage(), DWLP_MSGRESULT, (LONG)FALSE);
		}

		return TRUE;
	}
	return FALSE;
}


BOOL CPropPageCalibrationStandard::GetCalibrationStandard(
						IDispatch	**	ppdisp)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	DISPPARAMS		dispParams;
	*ppdisp		= NULL;
	if (this->GetOurObject(&pdisp))
	{
		VariantInit(&varResult);
		varResult.vt		= VT_DISPATCH;
		varResult.pdispVal	= NULL;
//		hr = Utils_InvokePropertyGet(pdisp, DISPID_CalibrationStandard, NULL, 0, &varResult);
		dispParams.cArgs			= 0;
		dispParams.cNamedArgs		= 0;
		dispParams.rgdispidNamedArgs	= NULL;
		dispParams.rgvarg			= NULL;
		hr = pdisp->Invoke(DISPID_CalibrationStandard, IID_NULL, 0x0409, DISPATCH_PROPERTYGET,
			&dispParams, &varResult, NULL, NULL);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
			{
				*ppdisp	= varResult.pdispVal;
				varResult.pdispVal->AddRef();
				fSuccess = TRUE;
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CPropPageCalibrationStandard::GetFileName(
						LPTSTR			szFileName,
						UINT			nBufferSize)
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	BOOL					fSuccess		= FALSE;
	szFileName[0]	= '\0';
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		fSuccess = calStandard.GetfileName(szFileName, nBufferSize);
		pdisp->Release();
	}
	return fSuccess;
}

void CPropPageCalibrationStandard::SetFileName(
						LPCTSTR			szFileName)
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		BOOL		fExcel	= calStandard.checkExcelFile(szFileName);

		calStandard.SetfileName(szFileName);
//		calStandard.loadExcelFile(szFileName);
		pdisp->Release();
	}
}

BOOL CPropPageCalibrationStandard::GetAmLoaded()
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	BOOL					amLoaded		= FALSE;
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		amLoaded	= calStandard.GetamLoaded();
		pdisp->Release();
	}
	return amLoaded;
}

void CPropPageCalibrationStandard::SetAmLoaded(
						BOOL			fAmLoaded)
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		calStandard.SetamLoaded(fAmLoaded);
		pdisp->Release();
	}
}

BOOL CPropPageCalibrationStandard::Getwavelengths(
						long		*	nValues,
						double		**	ppWaves)
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	BOOL					fSuccess		= FALSE;
	*nValues	= 0;
	*ppWaves	= NULL;
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		fSuccess = calStandard.Getwavelengths(nValues, ppWaves);
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CPropPageCalibrationStandard::Getcalibration(
						long		*	nValues,
						double		**	ppcalibration)
{
	IDispatch			*	pdisp;
	CCalibrationStandard	calStandard;
	BOOL					fSuccess	= FALSE;
	*nValues		= 0;
	*ppcalibration	= NULL;
	if (this->GetCalibrationStandard(&pdisp))
	{
		calStandard.doInit(pdisp);
		fSuccess = calStandard.Getcalibration(nValues, ppcalibration);
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CPropPageCalibrationStandard::SelectCalibrationFile()
{
	CSciFileOpen	fileOpen;
	TCHAR			szFileName[MAX_PATH];
	BOOL			fSuccess		= FALSE;
	LPTSTR			szConfigFile = NULL;
	IDispatch	*	pdisp;
	HRESULT			hr;
	TCHAR			szBaseFolder[MAX_PATH];
	LPTSTR			szFolder = NULL;

	if (this->GetOurObject(&pdisp))
	{
		Utils_GetStringProperty(pdisp, DISPID_ConfigFile, &szConfigFile);
		if (NULL != szConfigFile)
		{
			PathRemoveFileSpec(szConfigFile);
			// check if this base working directory exists
			if (PathFileExists(szConfigFile))
			{
				fileOpen.SetBaseWorkingDirectory(szConfigFile);
			}
			else
			{
				// folder doesn't exist (probably file comes from other computer)
				// use My Documents // SciSpec if it exists
				hr = SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &szFolder);
				if (SUCCEEDED(hr))
				{
					StringCchCopy(szBaseFolder, MAX_PATH, szFolder);
					PathAppend(szBaseFolder, L"SciSpec");
					if (PathFileExists(szBaseFolder))
					{
						fileOpen.SetBaseWorkingDirectory(szBaseFolder);
					}
					else
					{
						fileOpen.SetBaseWorkingDirectory(szFolder);
					}
					CoTaskMemFree((LPVOID)szFolder);
				}
			}
			CoTaskMemFree((LPVOID)szConfigFile);
		}
	}

//	fileOpen.AddFileExtension(TEXT("*.*"), TEXT("All Files"));
//	fileOpen.AddFileExtension(TEXT("*.src"), TEXT("Calibration Standard Files"));
//	if (fileOpen.SelectDataFile(this->GetMyPage(), szFileName, MAX_PATH))
	if (fileOpen.SelectCalibrationFile(this->GetMyPage(), szFileName, MAX_PATH))
	{
		fSuccess = TRUE;
		this->SetFileName(szFileName);
	}
	return fSuccess;
/*

	HRESULT			hr;
	LPOLESTR		ProgID		= NULL;
	CLSID			clsid;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	LPTSTR			szFileName	= NULL;
	BOOL			fSuccess	= FALSE;

	this->SetAmLoaded(FALSE);
	Utils_AnsiToUnicode(TEXT("Sciencetech.SciFileOpen.1"), &ProgID);
	hr = CLSIDFromProgID(ProgID, &clsid);
	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*) &pdisp);
	}
	if (SUCCEEDED(hr))
	{
		


		Utils_GetMemid(pdisp, TEXT("SelectDataFile"), &dispid);
		InitVariantFromInt32((long) this->GetMyPage(), &varg);
		hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			Utils_VariantToString(&varResult, &szFileName);
			this->SetFileName(szFileName);
			CoTaskMemFree((LPVOID) szFileName);
			fSuccess	= TRUE;
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
*/
}

void CPropPageCalibrationStandard::DisplayCurrentData()
{
	TCHAR			szFileName[MAX_PATH];

	if (NULL == this->m_pMyPlotWindow) return;
	this->m_pMyPlotWindow->ClearPlotData();
	if (this->GetFileName(szFileName, MAX_PATH))
	{
		SetDlgItemText(this->GetMyPage(), IDC_EDITFILENAME, szFileName);
//		this->SetAmLoaded(FALSE);
		this->SetAmLoaded(TRUE);
		if (this->GetAmLoaded())
		{
			long		nX;
			double	*	pX	= NULL;
			long		nY;
			double	*	pY	= NULL;
			if (this->Getwavelengths(&nX, &pX))
			{
				if (this->Getcalibration(&nY, &pY))
				{
					if (nX == nY)
					{
						this->m_pMyPlotWindow->SetPlotData(nX, pX, pY);
					}
					delete [] pY;
				}
				delete [] pX;
			}
		}
	}
//	InvalidateRect(GetDlgItem(this->GetMyPage(), IDC_PLOT), NULL, TRUE);
	this->m_pMyPlotWindow->Refresh();
}

void CPropPageCalibrationStandard::DisplayexpectedDistance()
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	TCHAR				szString[MAX_PATH];
	if (this->GetCalibrationStandard(&pdisp))
	{
		Utils_GetMemid(pdisp, L"expectedDistance", &dispid);
		_stprintf_s(szString, L"%8.4f cm", Utils_GetDoubleProperty(pdisp, dispid));
		SetDlgItemText(this->GetMyPage(), IDC_EDITEXPECTEDDISTANCE, szString);
		pdisp->Release();
	}
}

void CPropPageCalibrationStandard::UpdateexpectedDistance()
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	TCHAR				szString[MAX_PATH];
	float				f1;
	if (this->GetCalibrationStandard(&pdisp))
	{
		Utils_GetMemid(pdisp, L"expectedDistance", &dispid);
		if (GetDlgItemText(this->GetMyPage(), IDC_EDITEXPECTEDDISTANCE, szString, MAX_PATH) > 0)
		{
			if (1 == _stscanf_s(szString, L"%f", &f1))
			{
				Utils_SetDoubleProperty(pdisp, dispid, f1);
			}
		}
		pdisp->Release();
	}
}

void CPropPageCalibrationStandard::DisplaysourceDistance()
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	TCHAR				szString[MAX_PATH];
	if (this->GetCalibrationStandard(&pdisp))
	{
		Utils_GetMemid(pdisp, L"sourceDistance", &dispid);
		_stprintf_s(szString, L"%8.4f cm", Utils_GetDoubleProperty(pdisp, dispid));
		SetDlgItemText(this->GetMyPage(), IDC_EDITSOURCEDISTANCE, szString);
		pdisp->Release();
	}
}

void CPropPageCalibrationStandard::UpdatesourceDistance()
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	TCHAR				szString[MAX_PATH];
	float				f1;
	if (GetDlgItemText(this->GetMyPage(), IDC_EDITSOURCEDISTANCE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &f1))
		{
			if (this->GetCalibrationStandard(&pdisp))
			{
				Utils_GetMemid(pdisp, L"sourceDistance", &dispid);
				Utils_SetDoubleProperty(pdisp, dispid, f1);
				pdisp->Release();
			}
		}
	}
}

// plot window events
void CPropPageCalibrationStandard::OnPlotWindowMouseMove(
	double			x,
	double			y)
{
	TCHAR			szString[MAX_PATH];
	_stprintf_s(szString, MAX_PATH, TEXT("X = %4.1f  Y = %8.4g"), x, y);
	SetDlgItemText(this->GetMyPage(), IDC_STATUS, szString);
}
