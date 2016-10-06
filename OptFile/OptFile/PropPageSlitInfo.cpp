#include "stdafx.h"
#include "PropPageSlitInfo.h"
#include "dispids.h"
#include "slitInfo.h"

CPropPageSlitInfo::CPropPageSlitInfo(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_SlitInfo, IDS_SlitInfo)
{
}

CPropPageSlitInfo::~CPropPageSlitInfo(void)
{
}

BOOL CPropPageSlitInfo::OnInitDialog()
{
	// subclass the edit controls
	this->SubclassEditBox(IDC_EDITINPUTSLITNAME);
	this->SubclassEditBox(IDC_EDITINPUTSLIT);
	this->SubclassEditBox(IDC_EDITINTERMEDIATESLITNAME);
	this->SubclassEditBox(IDC_EDITINTERMEDIATESLIT);
	this->SubclassEditBox(IDC_EDITOUTPUTSLITNAME);
	this->SubclassEditBox(IDC_EDITOUTPUTSLIT);
	// display the slit info
	this->DisplayInputSlit();
	this->DisplayIntermediateSlit();
	this->DisplayOutputSlit();
	return CMyPropPage::OnInitDialog();
}

BOOL CPropPageSlitInfo::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_SETINPUTSLIT:
		this->SetSlitInfo(TEXT("InputSlit"), IDC_EDITINPUTSLITNAME, IDC_EDITINPUTSLIT);
		return TRUE;
	case IDC_EDITINPUTSLIT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetSlitInfo(TEXT("InputSlit"), IDC_EDITINPUTSLITNAME, IDC_EDITINPUTSLIT);
			return TRUE;
		}
		break;
	case IDC_SETINTERMEDIATESLIT:
		this->SetSlitInfo(TEXT("IntermediateSlit"), IDC_EDITINTERMEDIATESLITNAME, IDC_EDITINTERMEDIATESLIT);
		return TRUE;
	case IDC_EDITINTERMEDIATESLIT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetSlitInfo(TEXT("IntermediateSlit"), IDC_EDITINTERMEDIATESLITNAME, IDC_EDITINTERMEDIATESLIT);
			return TRUE;
		}
		break;
	case IDC_SETOUTPUTSLIT:
		this->SetSlitInfo(TEXT("OutputSlit"), IDC_EDITOUTPUTSLITNAME, IDC_EDITOUTPUTSLIT);
		return TRUE;
	case IDC_EDITOUTPUTSLIT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetSlitInfo(TEXT("OutputSlit"), IDC_EDITOUTPUTSLITNAME, IDC_EDITOUTPUTSLIT);
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

void CPropPageSlitInfo::OnEditReturnClicked(
						UINT			nID)
{
	switch (nID)
	{
	case IDC_EDITINPUTSLITNAME:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITINPUTSLIT));
		break;
	case IDC_EDITINPUTSLIT:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITINTERMEDIATESLITNAME));
		break;
	case IDC_EDITINTERMEDIATESLITNAME:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITINTERMEDIATESLIT));
		break;
	case IDC_EDITINTERMEDIATESLIT:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITOUTPUTSLITNAME));
		break;
	case IDC_EDITOUTPUTSLITNAME:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITOUTPUTSLIT));
		break;
	case IDC_EDITOUTPUTSLIT:
		SetFocus(GetDlgItem(this->GetMyPage(), IDC_EDITINPUTSLITNAME));
		break;
	default:
		break;
	}
}


BOOL CPropPageSlitInfo::GetSlitInfo(
						IDispatch	**	ppdisp)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	BOOL			fSuccess	= FALSE;
	VARIANT			varResult;
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_SlitInfo, NULL, 0, &varResult);
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

void CPropPageSlitInfo::DisplayInputSlit()
{
	this->DisplaySlitInfo(TEXT("InputSlit"), IDC_EDITINPUTSLITNAME, IDC_EDITINPUTSLIT);
}

void CPropPageSlitInfo::DisplayIntermediateSlit()
{
	this->DisplaySlitInfo(TEXT("intermediateSlit"), IDC_EDITINTERMEDIATESLITNAME, IDC_EDITINTERMEDIATESLIT);
}

void CPropPageSlitInfo::DisplayOutputSlit()
{
	this->DisplaySlitInfo(TEXT("outputSlit"), IDC_EDITOUTPUTSLITNAME, IDC_EDITOUTPUTSLIT);
}

void CPropPageSlitInfo::DisplaySlitInfo(
						LPCTSTR			szSlitTitle,
						UINT			nSlitName,
						UINT			nSlitWidth)
{
	IDispatch	*	pdisp;
	CSlitInfo		slitInfo;
	double			slitWidth	= 0.0;
	TCHAR			szString[MAX_PATH];
	SetDlgItemText(this->GetMyPage(), nSlitName, TEXT(""));
	SetDlgItemText(this->GetMyPage(), nSlitWidth, TEXT(""));
	if (this->GetSlitInfo(&pdisp))
	{
		slitInfo.doInit(pdisp);
		if (slitInfo.haveSlit(szSlitTitle))
		{
			if (slitInfo.getSlitName(szSlitTitle, szString, MAX_PATH))
			{
				SetDlgItemText(this->GetMyPage(), nSlitName, szString);
			}
			_stprintf_s(szString, MAX_PATH, TEXT("%5.2f"), slitInfo.getSlitWidth(szSlitTitle));
			SetDlgItemText(this->GetMyPage(), nSlitWidth, szString);
		}
		pdisp->Release();
	}
}

void CPropPageSlitInfo::SetSlitInfo(
						LPCTSTR			szSlitTitle,
						UINT			nSlitName,
						UINT			nSlitWidth)
{
	IDispatch	*	pdisp;
	CSlitInfo		slitInfo;
	float			f1;
	double			slitWidth	= 0.0;
	TCHAR			szString[MAX_PATH];
	TCHAR			szSlitName[MAX_PATH];

	if (GetDlgItemText(this->GetMyPage(), nSlitName, szSlitName, MAX_PATH) > 0	&&
		GetDlgItemText(this->GetMyPage(), nSlitWidth, szString, MAX_PATH) > 0	&&
		1 == _stscanf_s(szString, TEXT("%f"), &f1))
	{
		slitWidth = f1;
		if (this->GetSlitInfo(&pdisp))
		{
			slitInfo.doInit(pdisp);
			if (slitInfo.haveSlit(szSlitTitle))
			{
				slitInfo.setSlitName(szSlitTitle, szSlitName);
				slitInfo.setSlitWidth(szSlitTitle, slitWidth);
			}
			else
			{
				slitInfo.addSlit(szSlitTitle, szSlitName, slitWidth);
			}
			pdisp->Release();
		}
	}
	this->DisplaySlitInfo(szSlitTitle, nSlitName, nSlitWidth);
}
