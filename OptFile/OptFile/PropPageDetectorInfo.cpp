#include "stdafx.h"
#include "PropPageDetectorInfo.h"
#include "dispids.h"
#include "DetectorInfo.h"
#include "IniFile.h"

CPropPageDetectorInfo::CPropPageDetectorInfo(CMyObject * pMyObject) :
	CMyPropPage(pMyObject, IDD_PROPPAGE_DetectorInfo, IDS_DetectorInfo),
	m_pDetectors(NULL),
	m_nDetectors(0),
	m_currentDetector(0)
{
}

CPropPageDetectorInfo::~CPropPageDetectorInfo(void)
{
//	long			iDet;
//	long			nGains;
//	long			iGain;
	if (NULL != this->m_pDetectors)
	{
		delete [] this->m_pDetectors;
		this->m_pDetectors	= NULL;
		this->m_nDetectors	= 0;
	}
}

BOOL CPropPageDetectorInfo::OnInitDialog()
{
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
//	TCHAR			szDetector[MAX_PATH];
	TCHAR			szgain[MAX_PATH];
//	double			temperature;
	TCHAR			szString[MAX_PATH];
	long			nChannels;

	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		// check for read only
		if (detInfo.GetamReadOnly())
		{
			SendMessage(GetDlgItem(this->GetMyPage(), IDC_TEMPERATURE), EM_SETREADONLY, TRUE, 0);
			EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDDETECTORNAME), FALSE);
			EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDGAINSETTING), FALSE);
		}
		else
		{
			// subclass the edit controls
			this->SubclassEditBox(IDC_TEMPERATURE);
			// initialize the detector info
			this->InitDetectorInfo();
			EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDDETECTORNAME), this->m_nDetectors > 0);
			EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDGAINSETTING), this->m_nDetectors > 0);
		}
		// display the current values
//		if (detInfo.GetdetectorModel(szDetector, MAX_PATH))
//			SetDlgItemText(this->GetMyPage(), IDC_DETECTORNAME, szDetector);
		this->DisplayDetectorName();
		if (detInfo.GetgainSetting(szgain, MAX_PATH))
			SetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, szgain);
		if (detInfo.GetportName(szString, MAX_PATH))
			SetDlgItemText(this->GetMyPage(), IDC_EDITPORTNAME, szString);
		_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), detInfo.Gettemperature());
		SetDlgItemText(this->GetMyPage(), IDC_TEMPERATURE, szString);
		nChannels = detInfo.GetnumADChannels();
		SetDlgItemInt(this->GetMyPage(), IDC_NUMBERCHANNELSETTINGS, nChannels, FALSE);
		EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDCHANNELDISPLAY), nChannels > 1);
		if (nChannels > 0)
		{
			SetDlgItemInt(this->GetMyPage(), IDC_EDITDISPLAYCHANNEL, 0, FALSE);
			SetDlgItemInt(this->GetMyPage(), IDC_ADCHANNEL, detInfo.getADChannel(0), FALSE);
			_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), detInfo.getLowLimit(0));
			SetDlgItemText(this->GetMyPage(), IDC_MINWAVE, szString);
			_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), detInfo.getHiLimit(0));
			SetDlgItemText(this->GetMyPage(), IDC_MAXWAVE, szString);
		}
		else
		{
			SetDlgItemText(this->GetMyPage(), IDC_ADCHANNEL, TEXT(""));
			SetDlgItemText(this->GetMyPage(), IDC_MINWAVE, TEXT(""));
			SetDlgItemText(this->GetMyPage(), IDC_MAXWAVE, TEXT(""));
		}
		this->DisplayADInfoString();
		// gain setting
		this->InitGainSetting();
		pdisp->Release();
	}
	return CMyPropPage::OnInitDialog();
}

BOOL CPropPageDetectorInfo::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_TEMPERATURE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusTemperature();
			return TRUE;
		}
		else if (EN_CHANGE == wmEvent)
		{
			this->SetPageChanged();
			return TRUE;
		}
		break;
	case IDC_CHKPBS:
		if (BN_CLICKED == wmEvent)
		{
			this->SetDetectorName(L"UVSi/PbS");
			return TRUE;
		}
		break;
	case IDC_CHKPBSE:
		if (BN_CLICKED == wmEvent)
		{
			this->SetDetectorName(L"UVSi/PbSe");
			return TRUE;
		}
		break;
	case IDC_CHKPMT:
		if (BN_CLICKED == wmEvent)
		{
			this->SetDetectorName(L"PMT");
			return TRUE;
		}
		break;
	case IDC_COMBOGAINSETTING:
		if (CBN_SELCHANGE == wmEvent)
		{
			this->SetGainSetting();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

void CPropPageDetectorInfo::OnEditReturnClicked(
						UINT			nID)
{
	switch (nID)
	{
	case IDC_TEMPERATURE:
		this->OnKillFocusTemperature();
		break;
	default:
		break;
	}
}

BOOL CPropPageDetectorInfo::GetDetectorInfo(
						IDispatch	**	ppdisp)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	BOOL			fSuccess		= FALSE;
	VARIANT			varResult;

	*ppdisp		= NULL;
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_DetectorInfo, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
			{
				*ppdisp	= varResult.pdispVal;
				varResult.pdispVal->AddRef();
				fSuccess	= TRUE;
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CPropPageDetectorInfo::GetDetectorInfo(
						CDetectorInfo*	pDetInfo)
{
	IDispatch	*	pdisp;
	BOOL			fSuccess	= FALSE;
	if (NULL == pDetInfo) return FALSE;
	if (this->GetDetectorInfo(&pdisp))
	{
		fSuccess = pDetInfo->doInit(pdisp);
		pdisp->Release();
	}
	return fSuccess;
}


void CPropPageDetectorInfo::OnKillFocusTemperature()
{
	TCHAR			szString[MAX_PATH];
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	float			fval;
	BOOL			fSuccess		= FALSE;

	if (this->GetDetectorInfo(&pdisp))
	{	
		detInfo.doInit(pdisp);
		if (GetDlgItemText(this->GetMyPage(), IDC_TEMPERATURE, szString, MAX_PATH) > 0)
		{
			if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
			{
				detInfo.Settemperature(fval);
			}
		}
		_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), (float) detInfo.Gettemperature());
		SetDlgItemText(this->GetMyPage(), IDC_TEMPERATURE, szString);
		pdisp->Release();
	}
}

BOOL CPropPageDetectorInfo::GetConfigFile(
						LPTSTR			szConfig,
						UINT			nBufferSize)
{
	IDispatch	*	pdisp;
	HRESULT			hr;
//	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	LPTSTR			szTemp		= NULL;
	szConfig[0]	= '\0';

	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_GetStringProperty(pdisp, DISPID_ConfigFile, &szTemp);
		if (SUCCEEDED(hr))
		{
			fSuccess = TRUE;
			StringCchCopy(szConfig, nBufferSize, szTemp);
			CoTaskMemFree((LPVOID) szTemp);
		}
		pdisp->Release();
	}
	return fSuccess;
}

long CPropPageDetectorInfo::GetNumberDetectors(
						CIniFile	*	pConfig)
{
	TCHAR			szAppName[MAX_PATH];
	TCHAR			szValue[MAX_PATH];
	TCHAR			szResult[16];
	long			i		= 0;
	BOOL			fDone	= FALSE;

	StringCchCopy(szValue, MAX_PATH, TEXT("Port"));
	while (!fDone)
	{
		GetAppName(i, szAppName, MAX_PATH);
		if (!pConfig->ReadStringValue(szAppName, szValue, NULL, szResult, 16))
		{
			fDone = TRUE;
		}
		else 
		{
			i++;
		}
	}
	return i;
}

BOOL CPropPageDetectorInfo::GetDetectorName(
						CIniFile	*	pConfig,
						long			iDetector,
						LPTSTR			szName,
						UINT			nBufferSize)
{
	TCHAR			szAppName[MAX_PATH];

	szName[0]	= '\0';
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	return pConfig->ReadStringValue(szAppName, TEXT("Name"), NULL, szName, nBufferSize);
}

BOOL CPropPageDetectorInfo::GetDetectorPort(
						CIniFile	*	pConfig,
						long			iDetector,
						LPTSTR			szPort,
						UINT			nBufferSize)
{
	TCHAR			szAppName[MAX_PATH];
	long			lval;
	BOOL			fSuccess	= FALSE;

	szPort[0]	= '\0';
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	if (pConfig->ReadIntValue(szAppName, TEXT("Port"), 0, &lval))
	{
		if (3 == lval)
		{
			StringCchPrintf(szPort, nBufferSize, TEXT("front Output"));
			fSuccess = TRUE;
		}
		else if (4 == lval)
		{
			StringCchPrintf(szPort, nBufferSize, TEXT("side output"));
			fSuccess = TRUE;
		}
		else
		{
			StringCchPrintf(szPort, nBufferSize, TEXT("Unknown"));
		}
	}
	return fSuccess;
}

long CPropPageDetectorInfo::GetNumGains(
						CIniFile	*	pConfig,
						long			iDetector)
{
	TCHAR			szAppName[MAX_PATH];
	TCHAR			szValue[MAX_PATH];
	long			i		= 0;
	BOOL			fDone	= FALSE;
	TCHAR			szDummy[16];
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	while (!fDone)
	{
		StringCchPrintf(szValue, MAX_PATH, TEXT("Gain%1d"), i+1);
		if (!pConfig->ReadStringValue(szAppName, szValue, NULL, szDummy, 16))
			fDone = TRUE;
		if (!fDone) i++;
	}
	return i;
}

BOOL CPropPageDetectorInfo::GetGainSetting(
						CIniFile	*	pConfig,
						long			iDetector,
						long			iGain,
						LPTSTR			szGain,
						UINT			nBufferSize)
{
	TCHAR			szAppName[MAX_PATH];
	TCHAR			szValue[MAX_PATH];
	szGain[0]	= '\0';
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	StringCchPrintf(szValue, MAX_PATH, TEXT("Gain%1d"), iGain+1);
	return pConfig->ReadStringValue(szAppName, szValue, NULL, szGain, nBufferSize);
}

BOOL CPropPageDetectorInfo::GetHaveTempControl(
						CIniFile	*	pConfig,
						long			iDetector)
{
	TCHAR		szAppName[MAX_PATH];
	long		lval;
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	pConfig->ReadIntValue(szAppName, TEXT("HaveTempControl"), 0, &lval);
	return 0 != lval;
}

BOOL CPropPageDetectorInfo::GetAppName(
						long			iDetector,
						LPTSTR			szApp,
						UINT			nBufferSize)
{
	StringCchPrintf(szApp, nBufferSize, TEXT("Detector%1d"), iDetector+1);
	return TRUE;
}

void CPropPageDetectorInfo::InitDetectorInfo()
{
	TCHAR			szConfig[MAX_PATH];
	CIniFile		iniFile;
	long			nDetectors;
	long			iDet;
//	long			iGain;
//	TCHAR			szGain[16];
//	long			iChannel;

	if (this->GetConfigFile(szConfig, MAX_PATH))
	{
		if (iniFile.Init(szConfig))
		{
			nDetectors = this->GetNumberDetectors(&iniFile);
			if (nDetectors > 0)
			{
				this->m_pDetectors	= new CSingleDetector [nDetectors];
				this->m_nDetectors	= nDetectors;
				for (iDet=0; iDet < this->m_nDetectors; iDet++)
				{
					this->m_pDetectors[iDet].SetDetector(iDet);
					this->m_pDetectors[iDet].ReadConfig(this, &iniFile);
				}
			}
/*
					this->GetDetectorName(&iniFile, iDet+1, this->m_pDetectors[iDet].szName, 16);
					this->GetDetectorPort(&iniFile, iDet+1, this->m_pDetectors[iDet].szPort, 16);
					this->m_pDetectors[iDet].haveTempControl	= this->GetHaveTempControl(&iniFile, iDet+1);
					this->m_pDetectors[iDet].nGains		= this->GetNumGains(&iniFile, iDet+1);
					if (this->m_pDetectors[iDet].nGains > 0)
					{
						this->m_pDetectors[iDet].arrGains	= new LPTSTR [this->m_pDetectors[iDet].nGains];
						for (iGain=0; iGain < this->m_pDetectors[iDet].nGains; iGain++)
						{
							this->m_pDetectors[iDet].arrGains[iGain]	= NULL;
							if (this->GetGainSetting(&iniFile, iDet+1, iGain+1, szGain, 16))
							{
								Utils_DoCopyString(
									&(this->m_pDetectors[iDet].arrGains[iGain]), szGain);
							}
						}
					}
					else
					{
						this->m_pDetectors[iDet].arrGains	= NULL;
					}
					this->m_pDetectors[iDet].nChannels		= this->GetNumChannels(&iniFile, iDet+1);
					if (this->m_pDetectors[iDet].nChannels > 0)
					{
						this->m_pDetectors[iDet].arrChannels	= new long [this->m_pDetectors[iDet].nChannels];
						for (iChannel = 0; iChannel < this->m_pDetectors[iDet].nChannels; iChannel++)
						{
							this->m_pDetectors[iDet].arrChannels[iChannel] =
								this->GetChannelSetting(&iniFile, iDet+1, iChannel+1);
						}
					}
					else
					{
						this->m_pDetectors[iDet].arrChannels	= NULL;
					}
				}
			}
*/
		}
	}
}

BOOL CPropPageDetectorInfo::OnNotify(
						LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code)
	{
		LPNMUPDOWN		pnmud	= (LPNMUPDOWN) pnmh;
		long			iDelta	= pnmud->iDelta;
		switch (pnmh->idFrom)
		{
		case IDC_UPDDETECTORNAME:
			this->OnDeltaPosDetectorName(iDelta);
			return TRUE;
		case IDC_UPDGAINSETTING:
			this->OnDeltaPosGainSetting(iDelta);
			return TRUE;
		case IDC_UPDCHANNELDISPLAY:
			this->OnDeltaPosChannelDisplay(iDelta);
			return TRUE;
		default:
			break;
		}
	}
	else if (PSN_APPLY == pnmh->code)
	{
		this->ApplyDetector();
		return TRUE;
	}
	return FALSE;
}

void CPropPageDetectorInfo::OnDeltaPosDetectorName(
						long			iDelta)
{
	// find the current displayed detector
	TCHAR			szDetector[MAX_PATH];
	long			iDet;
	BOOL			fDone;
	TCHAR			szString[MAX_PATH];
	long			nGains;

	if (NULL == this->m_pDetectors) return;

	if (GetDlgItemText(this->GetMyPage(), IDC_DETECTORNAME, szDetector, MAX_PATH) > 0)
	{
		iDet = 0;
		fDone = FALSE;
		while (iDet < this->m_nDetectors && !fDone)
		{
			if (this->m_pDetectors[iDet].CheckName(szDetector))
			{
				fDone = TRUE;
				this->m_currentDetector = iDet;
			}
			if (!fDone) iDet++;
		}
		if (!fDone) this->m_currentDetector = 0;
	}
	else
	{
		this->m_currentDetector	= 0;
	}
	this->m_currentDetector += iDelta;
	if (this->m_currentDetector >= this->m_nDetectors)
		this->m_currentDetector = 0;
	if (this->m_currentDetector < 0)
		this->m_currentDetector = this->m_nDetectors - 1;
	this->m_pDetectors[this->m_currentDetector].GetName(szDetector, MAX_PATH);
	SetDlgItemText(this->GetMyPage(), IDC_DETECTORNAME, szDetector);
	this->m_pDetectors[this->m_currentDetector].GetPort(szString, MAX_PATH);
	SetDlgItemText(this->GetMyPage(), IDC_EDITPORTNAME, szString);
	nGains = this->m_pDetectors[this->m_currentDetector].GetNumGains();
	if (0 == nGains)
	{
		EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDGAINSETTING), FALSE);
		SetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, TEXT(""));
	}
	else
	{
		this->m_pDetectors[this->m_currentDetector].GetGainString(0, szString, MAX_PATH);
		SetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, szString);
		EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDGAINSETTING), nGains > 1);
	}
	EnableWindow(GetDlgItem(this->GetMyPage(), IDC_UPDCHANNELDISPLAY),
		this->m_pDetectors[this->m_currentDetector].GetNumChannels() > 1);
	this->DisplayConfigChannelInfo(0);
	this->SetPageChanged();
}

void CPropPageDetectorInfo::OnDeltaPosGainSetting(
						long			iDelta)
{
	if (NULL == this->m_pDetectors) return;
	if (this->m_currentDetector >= 0 && this->m_currentDetector < this->m_nDetectors)
	{
		long		nGains	= this->m_pDetectors[this->m_currentDetector].GetNumGains();
		if (nGains > 1)
		{
			TCHAR			szGain[MAX_PATH];
			long			nGain		= 0;
			BOOL			fDone		= FALSE;
//			long			i;

			// find current gain
			if (GetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, szGain, MAX_PATH) > 0)
			{
				nGain = this->m_pDetectors[this->m_currentDetector].FindGainIndex(szGain);
				/*
			i = 0;
			fDone = FALSE;
			while (i < this->m_pDetectors[this->m_currentDetector].nGains && !fDone)
			{
				if (0 == lstrcmpi(szGain, this->m_pDetectors[this->m_currentDetector].arrGains[i]))
				{
					fDone = TRUE;
					nGain = i;
				}
				if (!fDone) i++;
				*/
			}
			nGain -= iDelta;
			if (nGain < 0)
				nGain = nGains - 1;
			else if (nGain >= nGains)
				nGain = 0;
			this->m_pDetectors[this->m_currentDetector].GetGainString(
				nGain, szGain, MAX_PATH);
			SetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, szGain);
		}
		this->SetPageChanged();
	}
}

void CPropPageDetectorInfo::OnDeltaPosChannelDisplay(
						long			iDelta)
{
	if (this->CheckReadOnly())
	{
		// channel display from current detector info
		long			channel		= GetDlgItemInt(this->GetMyPage(), IDC_ADCHANNEL, NULL, FALSE);
		long			index;
		CDetectorInfo	detInfo;
		long			nChannels;
		TCHAR			szString[MAX_PATH];
		if (this->GetDetectorInfo(&detInfo))
		{
			nChannels	= detInfo.GetnumADChannels();
			if (nChannels > 1)
			{
				index = detInfo.FindIndexFromChannel(channel);
				index -= iDelta;
				if (index < 0) index = nChannels-1;
				if (index >= nChannels) index = 0;
				SetDlgItemInt(this->GetMyPage(), IDC_ADCHANNEL, 
					detInfo.getADChannel(index), FALSE);
				_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), detInfo.getLowLimit(index));
				SetDlgItemText(this->GetMyPage(), IDC_MINWAVE, szString);
				_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), detInfo.getHiLimit(index));
				SetDlgItemText(this->GetMyPage(), IDC_MAXWAVE, szString);
			}
		}
	}
	else
	{
		// channel display from config file
		if (this->m_currentDetector >= 0 && this->m_currentDetector < this->m_nDetectors &&
			this->m_pDetectors[this->m_currentDetector].GetNumChannels() > 1)
		{
			long			channel = GetDlgItemInt(this->GetMyPage(), IDC_ADCHANNEL, NULL, FALSE);
			long			index	= this->m_pDetectors[this->m_currentDetector].FindChannelIndex(channel);
			long			numChannels	= this->m_pDetectors[this->m_currentDetector].GetNumChannels();
			index -= iDelta;
			if (index < 0) index = numChannels-1;
			if (index >= numChannels) index = 0;
			this->DisplayConfigChannelInfo(index);
		}
	}
}

void CPropPageDetectorInfo::DisplayConfigChannelInfo(
						long			index)
{
	TCHAR			szString[MAX_PATH];
	long			nChannels		= this->m_pDetectors[this->m_currentDetector].GetNumChannels();
	SetDlgItemInt(this->GetMyPage(), IDC_NUMBERCHANNELSETTINGS, nChannels, FALSE);
	if (0 == nChannels)
	{
		SetDlgItemText(this->GetMyPage(), IDC_ADCHANNEL, TEXT(""));
		SetDlgItemText(this->GetMyPage(), IDC_MINWAVE, TEXT(""));
		SetDlgItemText(this->GetMyPage(), IDC_MAXWAVE, TEXT(""));
	}
	else
	{
		SetDlgItemInt(this->GetMyPage(), IDC_EDITDISPLAYCHANNEL, index, FALSE);
		SetDlgItemInt(this->GetMyPage(), IDC_ADCHANNEL, 
			this->m_pDetectors[this->m_currentDetector].GetADChannel(index), FALSE); 
		_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"),
			this->m_pDetectors[this->m_currentDetector].GetLowLimit(index));
		SetDlgItemText(this->GetMyPage(), IDC_MINWAVE, szString);
		_stprintf_s(szString, MAX_PATH, TEXT("%5.1f"), 
			this->m_pDetectors[this->m_currentDetector].GetHiLimit(index));
		SetDlgItemText(this->GetMyPage(), IDC_MAXWAVE, szString);
	}
}

void CPropPageDetectorInfo::ApplyDetector()
{
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	TCHAR			szDetector[MAX_PATH];
//	TCHAR			szgain[MAX_PATH];
	TCHAR			szString[MAX_PATH];
	float			fval;
	long			nChannels;
	long			i;

	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		// check for read only
		if (!detInfo.GetamReadOnly())
		{
			if (GetDlgItemText(this->GetMyPage(), IDC_DETECTORNAME, szDetector, MAX_PATH) > 0)
			{
				detInfo.SetdetectorModel(szDetector);
			}
			if (GetDlgItemText(this->GetMyPage(), IDC_EDITPORTNAME, szString, MAX_PATH) > 0)
			{
				detInfo.SetportName(szString);
			}
			if (GetDlgItemText(this->GetMyPage(), IDC_GAINSETTING, szString, MAX_PATH) > 0)
			{
				detInfo.SetgainSetting(szString);
			}
			// check if the temperature window is read only
			if (0 == (
				GetWindowLongPtr(
				GetDlgItem(this->GetMyPage(), IDC_TEMPERATURE), GWL_STYLE) & ES_READONLY))
			{
				if (GetDlgItemText(this->GetMyPage(), IDC_TEMPERATURE, szString, MAX_PATH) > 0)
				{
					if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
					{
						detInfo.Settemperature(fval);
					}
				}
			}
			detInfo.clearADChannels();
			if (this->m_currentDetector >= 0 && this->m_currentDetector < this->m_nDetectors)
			{
				nChannels = this->m_pDetectors[this->m_currentDetector].GetNumChannels();
				for (i=0; i<nChannels; i++)
				{
					detInfo.addADChannel(
						this->m_pDetectors[this->m_currentDetector].GetADChannel(i),
						this->m_pDetectors[this->m_currentDetector].GetLowLimit(i),
						this->m_pDetectors[this->m_currentDetector].GetHiLimit(i));
				}
			}
		}
		pdisp->Release();
	}
}

long CPropPageDetectorInfo::GetNumChannels(
						CIniFile	*	pConfig,
						long			iDetector)
{
	TCHAR		szAppName[MAX_PATH];
	TCHAR		szValue[MAX_PATH];
	long		i;
	BOOL		fDone;
//	long		lval;
	TCHAR		szTest[16];
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	i = 0;
	fDone = FALSE;
	while (!fDone)
	{
		StringCchPrintf(szValue, MAX_PATH, TEXT("ADChannel%1d"), i+1);
		if (!pConfig->ReadStringValue(szAppName, szValue, NULL, szTest, 16))
		{
			fDone = TRUE;
		}
		if (!fDone) i++;
	}
	return i;
}

long CPropPageDetectorInfo::GetChannelSetting(
						CIniFile	*	pConfig,
						long			iDetector,
						long			iChannel,
						double		*	lowLimit,
						double		*	hiLimit)
{
	TCHAR		szAppName[MAX_PATH];
	TCHAR		szValue[MAX_PATH];
	long		retVal;
	this->GetAppName(iDetector, szAppName, MAX_PATH);
	StringCchPrintf(szValue, MAX_PATH, TEXT("ADChannel%1d"), iChannel+1);
	pConfig->ReadIntValue(szAppName, szValue, 0, &retVal);
	StringCchPrintf(szValue, MAX_PATH, TEXT("LowLimit%1d"), iChannel+1);
	pConfig->ReadDoubleValue(szAppName, szValue, 0.0, lowLimit);
	StringCchPrintf(szValue, MAX_PATH, TEXT("HighLimit%1d"), iChannel+1);
	pConfig->ReadDoubleValue(szAppName, szValue, 0.0, hiLimit);
	return retVal;
}

void CPropPageDetectorInfo::DisplayADInfoString()
{
	IDispatch		*	pdisp;
	VARIANT				varResult;
	HRESULT				hr;
	LPTSTR				infoString = NULL;
	LPTSTR				szTempString = NULL;
	size_t				slen, slen1;
	LPTSTR				szRem;
	LPTSTR				szTemp;
	TCHAR				szOneLine[MAX_PATH];
	size_t				allocSize;

	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_ADInfoString, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToStringAlloc(varResult, &infoString);
			if (SUCCEEDED(hr))
			{
				if (NULL == StrStr(infoString, L"\r\n"))
				{ 
					// need to add carriage returns to string
					StringCchLength(infoString, STRSAFE_MAX_CCH, &slen);
					if (slen > 0)
					{
						allocSize = 2 * slen;
						szTempString = (LPTSTR)CoTaskMemAlloc(allocSize * sizeof(TCHAR));
						szTempString[0] = '\0';
						szTemp = infoString;
						while (NULL != szTemp)
						{
							szRem = StrStr(szTemp, L"\n");
							if (NULL != szRem)
							{
								StringCchLength(szTemp, STRSAFE_MAX_CCH, &slen);
								StringCchLength(szRem, STRSAFE_MAX_CCH, &slen1);
								StringCchCopyN(szOneLine, MAX_PATH,
									szTemp, slen - slen1);
								StringCchLength(szTempString, allocSize, &slen);
								if (slen > 0)
								{
									StringCchCat(szTempString, allocSize, szOneLine);
									StringCchCat(szTempString, allocSize, L"\r\n");
								}
								else
								{
									StringCchCopy(szTempString, allocSize, szOneLine);
									StringCchCat(szTempString, allocSize, L"\r\n");
								}
							}
							else
							{
								StringCchCopy(szOneLine, MAX_PATH, szTemp);
								StringCchLength(szTempString, allocSize, &slen);
								if (slen > 0)
								{
									StringCchCat(szTempString, allocSize, szOneLine);
								}
								else
								{
									StringCchCopy(szTempString, allocSize, szOneLine);
								}
							}
							if (NULL != szRem)
							{
								StringCchLength(szRem, STRSAFE_MAX_CCH, &slen);
								if (slen > 2)
								{
									szTemp = &(szRem[1]);
								}
								else
								{
									szTemp = NULL;
								}
							}
							else
							{
								szTemp = NULL;
							}

						}
						SetDlgItemText(this->GetMyPage(), 
							IDC_EDITADINFOSETTINGS,
							szTempString);
						CoTaskMemFree((LPVOID)szTempString);
					}
					else
					{
						SetDlgItemText(this->GetMyPage(),
							IDC_EDITADINFOSETTINGS, L"");
					}
				}
				else
				{
					SetDlgItemText(this->GetMyPage(), IDC_EDITADINFOSETTINGS, infoString);
				}
				CoTaskMemFree((LPVOID)infoString);
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}


BOOL CPropPageDetectorInfo::CheckReadOnly()			// check if the detector info is read only
{
	BOOL			readOnly	= TRUE;
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		// check for read only
		readOnly = detInfo.GetamReadOnly();
		pdisp->Release();
	}
	return readOnly;
}

CPropPageDetectorInfo::CSingleDetector::CSingleDetector() :
	m_iDetector(-1),
	m_nGains(0),
	m_arrGains(NULL),
	m_haveTempControl(FALSE),
	m_nChannels(0),
	m_arrChannels(NULL),
	m_arrLowLimits(NULL),
	m_arrHiLimits(NULL)
{
	m_szName[0]		= '\0';
	m_szPort[0]		= '\0';
}

CPropPageDetectorInfo::CSingleDetector::~CSingleDetector()
{
	long			i;
	if (NULL != this->m_arrGains)
	{
		for (i=0; i<this->m_nGains; i++)
			CoTaskMemFree((LPVOID) this->m_arrGains[i]);
		delete [] this->m_arrGains;
		this->m_arrGains	= NULL;
	}
	this->m_nGains	= 0;
	if (NULL != this->m_arrChannels)
	{
		delete [] this->m_arrChannels;
		this->m_arrChannels	= NULL;
	}
	if (NULL != this->m_arrLowLimits)
	{
		delete [] this->m_arrLowLimits;
		this->m_arrLowLimits	= NULL;
	}
	if (NULL != this->m_arrHiLimits)
	{
		delete [] this->m_arrHiLimits;
		this->m_arrHiLimits		= NULL;
	}
	this->m_nChannels	= 0;
}

void CPropPageDetectorInfo::CSingleDetector::SetDetector(
							long			detector)
{
	this->m_iDetector	= detector;
}

BOOL CPropPageDetectorInfo::CSingleDetector::ReadConfig(
							CPropPageDetectorInfo * pPropPage,
							CIniFile	*	pConfig)
{
	if (!pPropPage->GetDetectorName(pConfig, this->m_iDetector, this->m_szName, 16))
		return FALSE;
	if (!pPropPage->GetDetectorPort(pConfig, this->m_iDetector, this->m_szPort, 16))
		return FALSE;
	long			nGains		= pPropPage->GetNumGains(pConfig, this->m_iDetector);
	long			i;
	TCHAR			szGain[MAX_PATH];
	if (nGains > 0)
	{
		this->m_nGains		= nGains;
		this->m_arrGains	= new LPTSTR [this->m_nGains];
		for (i=0; i<this->m_nGains; i++)
		{
			this->m_arrGains[i]	= NULL;
			if (pPropPage->GetGainSetting(pConfig, this->m_iDetector, i, szGain, MAX_PATH))
			{
//				Utils_DoCopyString(&(this->m_arrGains[i]), szGain);
				SHStrDup(szGain, &(this->m_arrGains[i]));
			}
		}
	}
	this->m_haveTempControl	= pPropPage->GetHaveTempControl(pConfig, this->m_iDetector);
	long		nChannels	= pPropPage->GetNumChannels(pConfig, this->m_iDetector);
	if (nChannels > 0)
	{
		this->m_arrChannels		= new long [nChannels];
		this->m_arrLowLimits	= new double [nChannels];
		this->m_arrHiLimits		= new double [nChannels];
		this->m_nChannels		= nChannels;
		for (i=0; i<this->m_nChannels; i++)
		{
			this->m_arrChannels[i]	= pPropPage->GetChannelSetting(pConfig, this->m_iDetector,
				i, &(this->m_arrLowLimits[i]), &(this->m_arrHiLimits[i]));
		}
	}
	return TRUE;
}

BOOL CPropPageDetectorInfo::CSingleDetector::CheckName(
							LPCTSTR			szName)
{
	return 0 == lstrcmpi(szName, this->m_szName);
}

void CPropPageDetectorInfo::CSingleDetector::GetName(
							LPTSTR			szName,
							UINT			nBufferSize)
{
	StringCchCopy(szName, nBufferSize, this->m_szName);
}

void CPropPageDetectorInfo::CSingleDetector::GetPort(
							LPTSTR			szPort,
							UINT			nBufferSize)
{
	StringCchCopy(szPort, nBufferSize, this->m_szPort);
}

long CPropPageDetectorInfo::CSingleDetector::GetNumGains()
{
	return this->m_nGains;
}

BOOL CPropPageDetectorInfo::CSingleDetector::GetGainString(
							long			iGain,
							LPTSTR			szGain,
							UINT			nBufferSize)
{
	BOOL			fSuccess	= FALSE;
	szGain[0]	= '\0';
	if (iGain >= 0 && iGain < this->m_nGains)
	{
		if (NULL != this->m_arrGains[iGain])
		{
			StringCchCopy(szGain, nBufferSize, this->m_arrGains[iGain]);
			fSuccess = TRUE;
		}
	}
	return fSuccess;
}

long CPropPageDetectorInfo::CSingleDetector::FindGainIndex(
							LPCTSTR			szGain)
{
	long		i		= 0;
	BOOL		fDone	= FALSE;
	long		retVal	= -1;
	while (i<this->m_nGains && !fDone)
	{
		if (NULL != this->m_arrGains[i])
		{
			if (0 == lstrcmpi(this->m_arrGains[i], szGain))
			{
				fDone = TRUE;
				retVal = i;
			}
		}
		if (!fDone) i++;
	}
	return retVal;
}

BOOL CPropPageDetectorInfo::CSingleDetector::GethaveTempControl()
{
	return this->m_haveTempControl;
}

long CPropPageDetectorInfo::CSingleDetector::GetNumChannels()
{
	return this->m_nChannels;
}

long CPropPageDetectorInfo::CSingleDetector::GetADChannel(
							long			index)
{
	if (index >= 0 && index < this->m_nChannels)
	{
		return this->m_arrChannels[index];
	}
	else
	{
		return -1;
	}
}

double CPropPageDetectorInfo::CSingleDetector::GetLowLimit(
							long			index)
{
	if (index >= 0 && index < this->m_nChannels)
	{
		return this->m_arrLowLimits[index];
	}
	else
	{
		return 0.0;
	}
}

double CPropPageDetectorInfo::CSingleDetector::GetHiLimit(
							long			index)
{
	if (index >= 0 && index < this->m_nChannels)
	{
		return this->m_arrHiLimits[index];
	}
	else
	{
		return 0.0;
	}
}

long CPropPageDetectorInfo::CSingleDetector::FindChannelIndex(
							long			channel)
{
	long			i		= 0;
	BOOL			fDone	= FALSE;
	long			retVal	= -1;
	while (i<this->m_nChannels && !fDone)
	{
		if (channel == this->m_arrChannels[i])
		{
			fDone = TRUE;
			retVal = i;
		}
		if (!fDone) i++;
	}
	return retVal;
}

void CPropPageDetectorInfo::DisplayDetectorName()
{
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	TCHAR			szDetector[MAX_PATH];

	Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPBS), BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPBSE), BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPMT), BST_UNCHECKED);
	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		if (detInfo.GetdetectorModel(szDetector, MAX_PATH))
		{
			if (0 == lstrcmpi(szDetector, L"UVSi/PbS"))
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPBS), BST_CHECKED);
			else if (0 == lstrcmpi(szDetector, L"UVSi/PbSe"))
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPBSE), BST_CHECKED);
			else if (0 == lstrcmpi(szDetector, L"PMT"))
				Button_SetCheck(GetDlgItem(this->GetMyPage(), IDC_CHKPMT), BST_CHECKED);
		}
		pdisp->Release();
	}
}

void CPropPageDetectorInfo::SetDetectorName(
	LPCTSTR			szName)
{
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		detInfo.SetdetectorModel(szName);
		pdisp->Release();
	}
	this->DisplayDetectorName();
}

void CPropPageDetectorInfo::InitGainSetting()
{
	HWND	hwndCombo = GetDlgItem(this->GetMyPage(), IDC_COMBOGAINSETTING);
	ComboBox_AddString(hwndCombo, L"Unknown");
	ComboBox_AddString(hwndCombo, L"Lo");
	ComboBox_AddString(hwndCombo, L"Hi");
	this->DisplayGainSetting();
}

void CPropPageDetectorInfo::DisplayGainSetting()
{
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;
	TCHAR			szGain[MAX_PATH];
	HWND			hwndCombo	= GetDlgItem(this->GetMyPage(), IDC_COMBOGAINSETTING);
	int				curSel = 0;
	if (this->GetDetectorInfo(&pdisp))
	{
		detInfo.doInit(pdisp);
		if (!detInfo.GetgainSetting(szGain, MAX_PATH))
		{
			StringCchCopy(szGain, MAX_PATH, L"Unknown");
		}
		pdisp->Release();
	}
	if (0 == lstrcmpi(szGain, L"Lo"))
		curSel = 1;
	else if (0 == lstrcmpi(szGain, L"Hi"))
		curSel = 2;
	ComboBox_SetCurSel(hwndCombo, curSel);
}

void CPropPageDetectorInfo::SetGainSetting()
{
	TCHAR			szGain[MAX_PATH];
	HWND			hwndCombo	= GetDlgItem(this->GetMyPage(), IDC_COMBOGAINSETTING);
	IDispatch	*	pdisp;
	CDetectorInfo	detInfo;

	if (GetDlgItemText(this->GetMyPage(), IDC_COMBOGAINSETTING, szGain, MAX_PATH) > 0)
	{
		if (this->GetDetectorInfo(&pdisp))
		{
			detInfo.doInit(pdisp);
			detInfo.SetgainSetting(szGain);
			pdisp->Release();
		}
	}
	this->DisplayGainSetting();
}
