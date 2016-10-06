#include "stdafx.h"
#include "PropPageExtraData.h"
#include "MyPlotWindow.h"
#include "dispids.h"
#include "Server.h"

CPropPageExtraData::CPropPageExtraData(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_ExtraData, IDS_ExtraData),
	m_pMyPlotWindow(NULL),
	m_nExtraData(0),
	m_pszExtraData(NULL),
	m_nCurrentData(0)
{
}

CPropPageExtraData::~CPropPageExtraData()
{
	Utils_DELETE_POINTER(this->m_pMyPlotWindow);
	if (NULL != this->m_pszExtraData)
	{
		ULONG			i;
		for (i = 0; i < this->m_nExtraData; i++)
		{
			CoTaskMemFree((LPVOID) this->m_pszExtraData[i]);
			this->m_pszExtraData[i] = NULL;
		}
		CoTaskMemFree((LPVOID) this->m_pszExtraData);
	}
	this->m_pszExtraData = NULL;
	this->m_nExtraData = 0;
}

/*
HPROPSHEETPAGE CPropPageExtraData::doCreatePropPage()
{
	PROPSHEETPAGE		psp;
	TCHAR				szTitle[MAX_PATH];

	this->GetTitleString(szTitle, MAX_PATH);

	ZeroMemory((PVOID)&psp, sizeof(PROPSHEETPAGE));
	psp.dwSize = sizeof(PROPSHEETPAGE);
	psp.dwFlags = PSP_USETITLE;
	psp.hInstance = GetServer()->GetInstance();
	psp.pszTemplate = MAKEINTRESOURCE(IDD_PROPPAGE_ExtraData);
	psp.pfnDlgProc = (DLGPROC)MyPropPage;
	psp.pszTitle = szTitle;
	psp.lParam = (LPARAM) this;
	psp.pfnCallback = NULL;
	return CreatePropertySheetPage(&psp);
}
*/

BOOL CPropPageExtraData::OnInitDialog()
{
	this->SubclassControl(IDC_PLOT);
	// create the plot window object
	this->m_pMyPlotWindow = new CMyPlotWindow(GetDlgItem(this->GetMyPage(), IDC_PLOT));
	// get the extra data titles
	this->GetExtraDataTitles();
	EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDEXTRADATA), this->m_nExtraData > 0);		// enable the up down button if there are more then one set
	// display extra data
	this->DisplayExtraData();
	return CMyPropPage::OnInitDialog();
}

LRESULT CPropPageExtraData::SubclassProc(
	HWND			hwnd,
	UINT			uMsg,
	WPARAM			wParam,
	LPARAM			lParam)
{
	UINT			nID = GetDlgCtrlID(hwnd);
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
				szString[0] = '\0';
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

void CPropPageExtraData::DisplayExtraData()
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	VARIANTARG			avarg[2];
//	BSTR				bstr = NULL;
	SAFEARRAY		*	pSA = NULL;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	long				ub;
	double			*	arr;
	double			*	pX;
	long				i;
	TCHAR				szTitle[MAX_PATH];

	if (!this->GetTitleString(szTitle, MAX_PATH))
		return;

	if (this->GetOurObject(&pdisp))
	{
//		VariantInit(&avarg[1]);
//		avarg[1].vt = VT_BYREF | VT_BSTR;
//		avarg[1].pbstrVal = &bstr;
		InitVariantFromString(szTitle, &avarg[1]);
		VariantInit(&avarg[0]);
		avarg[0].vt = VT_BYREF | VT_ARRAY | VT_R8;
		avarg[0].pparray = &pSA;
		hr = Utils_InvokeMethod(pdisp, DISPID_GetExtraValues, avarg, 2, &varResult);
		VariantClear(&avarg[1]);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		if (fSuccess)
		{
			SetDlgItemText(this->GetMyPage(), IDC_EXTRADATATITLE, szTitle);
//			SendMessage(this->GetMyPage(), WM_SETTEXT, )
			SafeArrayGetUBound(pSA, 1, &ub);
			pX = new double[ub + 1];
			for (i = 0; i <= ub; i++)
				pX[i] = i * 1.0;
			SafeArrayAccessData(pSA, (void**)&arr);
			this->m_pMyPlotWindow->SetPlotData(ub + 1, pX, arr);
			InvalidateRect(GetDlgItem(this->GetMyPage(), IDC_PLOT), NULL, TRUE);
			SafeArrayUnaccessData(pSA);
			delete[] pX;
		}
//		SysFreeString(bstr);
		SafeArrayDestroy(pSA);
		pdisp->Release();
	}
}

BOOL CPropPageExtraData::GetTitleString(
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
/*
	HRESULT				hr;
	IDispatch		*	pdisp;
	VARIANTARG			avarg[2];
	BSTR				bstr = NULL;
	SAFEARRAY		*	pSA = NULL;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;

	if (this->GetOurObject(&pdisp))
	{
		VariantInit(&avarg[1]);
		avarg[1].vt = VT_BYREF | VT_BSTR;
		avarg[1].pbstrVal = &bstr;
		VariantInit(&avarg[0]);
		avarg[0].vt = VT_BYREF | VT_ARRAY | VT_R8;
		avarg[0].pparray = &pSA;
		hr = Utils_InvokeMethod(pdisp, DISPID_GetExtraValues, avarg, 2, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		if (fSuccess)
		{
			StringCchCopy(szTitle, nBufferSize, bstr);
		}
		SysFreeString(bstr);
		SafeArrayDestroy(pSA);
		pdisp->Release();
	}
	return fSuccess;
*/
	if (this->m_nCurrentData >= 0 && this->m_nCurrentData < this->m_nExtraData)
	{
		StringCchCopy(szTitle, nBufferSize, this->m_pszExtraData[this->m_nCurrentData]);
		return TRUE;
	}
	else
	{
		szTitle[0] = '\0';
		return FALSE;
	}
}

void CPropPageExtraData::GetExtraDataTitles()
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	VARIANT				varResult;
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_ExtraValueTitles, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToStringArrayAlloc(varResult, (PWSTR**)&this->m_pszExtraData, &this->m_nExtraData);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}

BOOL CPropPageExtraData::OnNotify(
	LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code && IDC_UPDEXTRADATA == pnmh->idFrom)
	{
		LPNMUPDOWN		pnmud = (LPNMUPDOWN)pnmh;
		ULONG			currentData = this->m_nCurrentData -= pnmud->iDelta;
		if (currentData >= this->m_nExtraData) currentData = 0;
		this->m_nCurrentData = currentData;
		this->DisplayExtraData();
		return TRUE;
	}
	return FALSE;
}
