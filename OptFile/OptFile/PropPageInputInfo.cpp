#include "stdafx.h"
#include "PropPageInputInfo.h"
#include "dispids.h"
#include "InputInfo.h"

// number of defined inputs
#define				NUM_INPUTS			7

CPropPageInputInfo::CPropPageInputInfo(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_InputInfo, IDS_InputInfo)
{
}

CPropPageInputInfo::~CPropPageInputInfo(void)
{
}

BOOL CPropPageInputInfo::OnInitDialog()
{
	// display the available input optics
	this->DisplayAvailableConfigs();
	// display the current values
	this->DisplayLampDistance();
	this->DisplayInputConfig();
	this->DisplayComment();
	// subclass the edit boxes
	this->SubclassEditBox(IDC_LAMPDISTANCE);
	this->SubclassEditBox(IDC_EDITUSERINPUT);
	return CMyPropPage::OnInitDialog();
}

BOOL CPropPageInputInfo::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_LAMPDISTANCE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->UpdateLampDistance();
			return TRUE;
		}
		break;
	case IDC_CHKinput1:
	case IDC_CHKinput2:
	case IDC_CHKinput3:
	case IDC_CHKinput4:
	case IDC_CHKinput5:
	case IDC_CHKinput6:
	case IDC_CHKinput7:
		if (BN_CLICKED == wmEvent)
		{
			TCHAR			szItem[MAX_PATH];
			if (GetDlgItemText(this->GetMyPage(), wmID, szItem, MAX_PATH))
			{
				SetInputConfig(szItem);
			}
			return TRUE;
		}
		break;
	case IDC_CHKUSERSET:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedChkUserSet();
			return TRUE;
		}
		break;
	case IDC_EDITUSERINPUT:
		if (EN_CHANGE == wmEvent)
		{
			// text changed need to reclick the checkbox
			Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKUSERSET), BST_UNCHECKED);
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

void CPropPageInputInfo::OnEditReturnClicked(
						UINT			nID)
{
	if (IDC_LAMPDISTANCE == nID)
	{
		this->UpdateLampDistance();
	}
	else if (IDC_EDITUSERINPUT == nID)
	{

	}
}

BOOL CPropPageInputInfo::GetInputInfo(
						IDispatch	**	ppdisp)
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*ppdisp		= NULL;
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_InputInfo, NULL, 0, &varResult);
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

void CPropPageInputInfo::DisplayLampDistance()
{
	IDispatch	*	pdisp;
	double			lampDistance		= 0.0;
	TCHAR			szString[MAX_PATH];
	if (this->GetInputInfo(&pdisp))
	{
		CInputInfo		inputInfo;
		if (inputInfo.doInit(pdisp))
		{
			lampDistance	= inputInfo.GetlampDistance();
		}
		pdisp->Release();
	}
	_stprintf_s(szString, MAX_PATH, TEXT("%5.2f cm"), lampDistance);
	SetDlgItemText(this->GetMyPage(), IDC_LAMPDISTANCE, szString);
}

void CPropPageInputInfo::SetLampDistance(
						double			lampDistance)
{
	IDispatch	*	pdisp;
	if (this->GetInputInfo(&pdisp))
	{
		CInputInfo		inputInfo;
		if (inputInfo.doInit(pdisp))
		{
			inputInfo.SetlampDistance(lampDistance);
		}
		pdisp->Release();
	}
}

void CPropPageInputInfo::UpdateLampDistance()
{
	TCHAR			szString[MAX_PATH];
	float			fval;

	if (GetDlgItemText(this->GetMyPage(), IDC_LAMPDISTANCE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
		{
			this->SetLampDistance(fval);
		}
	}
	this->DisplayLampDistance();
}

void CPropPageInputInfo::DisplayAvailableConfigs()
{
	IDispatch	*	pdisp;
	if (this->GetInputInfo(&pdisp))
	{
		CInputInfo		inputInfo;
		inputInfo.doInit(pdisp);
		UINT			nConfigs	= (UINT) inputInfo.GetnumInputConfigs();
		UINT			nStartID	= IDC_CHKinput1;
		UINT			i;
		TCHAR			szConfig[MAX_PATH];
		UINT			nID;
		for (i=0; i<nConfigs; i++)
		{
			nID		= nStartID + i;
			if (inputInfo.getInputConfig((long) i, szConfig, MAX_PATH))
			{
				EnableWindow(GetDlgItem(this->GetMyPage(), nID), TRUE);
				SetDlgItemText(this->GetMyPage(), nID, szConfig);
			}
		}
		// disable remainder
		while (i <= NUM_INPUTS)
		{
			nID		= nStartID + i;
			EnableWindow(GetDlgItem(this->GetMyPage(), nID), FALSE);
			i++;
		}
		pdisp->Release();
	}
}

void CPropPageInputInfo::DisplayInputConfig()
{
	TCHAR			szInputConfig[MAX_PATH];
	UINT			nStartID	= IDC_CHKinput1;
	IDispatch	*	pdisp;
	CInputInfo		inputInfo;
	UINT			nID;
	UINT			i;
	BOOL			fDone		= FALSE;
	TCHAR			szItem[MAX_PATH];
	BOOL			fUserInput = FALSE;
	double			userInput;

	for (i=0; i<NUM_INPUTS; i++)
	{
		nID = nStartID + i;
		SendMessage(GetDlgItem(this->GetMyPage(), nID), BM_SETCHECK, BST_UNCHECKED, 0);
	}
	if (this->GetInputInfo(&pdisp))
	{
		inputInfo.doInit(pdisp);
		fUserInput = inputInfo.GetuserSetinput(&userInput);
		if (inputInfo.GetinputOpticConfig(szInputConfig, MAX_PATH))
		{
			fDone = FALSE;
			if (!fUserInput)
			{
				i = 0;
				while (i < NUM_INPUTS && !fDone)
				{
					nID = nStartID + i;
					if (GetDlgItemText(this->GetMyPage(), nID, szItem, MAX_PATH) > 0)
					{
						if (0 == lstrcmpi(szItem, szInputConfig))
						{
							fDone = TRUE;
							SendMessage(GetDlgItem(this->GetMyPage(), nID), BM_SETCHECK,
								BST_CHECKED, 0);
							Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKUSERSET), BST_UNCHECKED);
						}
					}
					if (!fDone) i++;
				}
			}
			if (!fDone)
			{
				float		fval;
				size_t		slen;
				BOOL		fDone;
				size_t		i;
				size_t		c;
				TCHAR		szString[MAX_PATH];
				if (1 == _stscanf_s(szInputConfig, L"FOV %f", &fval))
				{
					StringCchLength(szInputConfig, MAX_PATH, &slen);
					fDone = FALSE;
					i = 4;
					c = 0;
					while (i < slen && !fDone)
					{
						if (' ' != szInputConfig[i])
						{
							szString[c] = szInputConfig[i];
							c++;
						}
						else
						{
							fDone = c > 0;
						}
						if (!fDone) i++;
					}
					szString[c] = '\0';
					SetDlgItemText(this->GetMyPage(), IDC_EDITUSERINPUT, szString);
					Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKUSERSET), BST_CHECKED);
				}
			}
		}
		pdisp->Release();
	}
}

void CPropPageInputInfo::SetInputConfig(
						LPCTSTR			szInputConfig)
{
	IDispatch	*	pdisp;
	if (this->GetInputInfo(&pdisp))
	{
		CInputInfo		inputInfo;
		inputInfo.doInit(pdisp);
		inputInfo.SetinputOpticConfig(szInputConfig);
		inputInfo.clearUserSetInput();
		pdisp->Release();
	}
	this->DisplayInputConfig();
}

void CPropPageInputInfo::DisplayComment()
{
	IDispatch		*	pdisp;
	HRESULT				hr;
	VARIANT				varResult;
	TCHAR				szComment[MAX_PATH];
	BOOL				fSuccess = FALSE;

	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_Comment, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, szComment, MAX_PATH);
			fSuccess = SUCCEEDED(hr);
			if (fSuccess)
				SetDlgItemText(this->GetMyPage(), IDC_EDITCOMMENT, szComment);
			VariantClear(&varResult);

		}
		pdisp->Release();
	}
	if (!fSuccess)
	{
		SetDlgItemText(this->GetMyPage(), IDC_EDITCOMMENT, L"");
	}
}

void CPropPageInputInfo::OnClickedChkUserSet()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	IDispatch	*	pdisp;

	if (GetDlgItemText(this->GetMyPage(), IDC_EDITUSERINPUT, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			if (this->GetInputInfo(&pdisp))
			{
				CInputInfo		inputInfo;
				inputInfo.doInit(pdisp);
				inputInfo.SetuserSetinput(fval);
		//		double		dval = inputInfo.GetuserSetinput();
				pdisp->Release();
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKUSERSET), BST_CHECKED);
				// un check the others
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput1), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput2), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput3), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput4), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput5), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput6), BST_UNCHECKED);
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKinput7), BST_UNCHECKED);
			}
		}
	}
}
